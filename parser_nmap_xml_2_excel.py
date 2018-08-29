# -*- coding: utf-8 -*-

"""
@author:随时静听
@file: parserXML.py
@time: 2018/08/23
@email:wang_di@topsec.com.cn
"""
# http://blog.51cto.com/maoyao/1772102
# https://xlsxwriter.readthedocs.io/format.html
import re
try:
    import xml.etree.cElementTree as ET
except:
    import xml.etree.ElementTree as ET
import glob

import os
import time

# import xlwt
import xlsxwriter

import argparse

from multiprocessing import Process,Pool,Lock

XMLPATH='report'

DEFAULT_STYLE={
        'font_size': 12,  # 字体大小
        'bold': False,  # 是否粗体
        # 'bg_color': '#101010',  # 表格背景颜色
        'font_color': 'black',  # 字体颜色
        'align': 'left',  # 居中对齐
        'valign':'vcenter',
        'font_name':'Courier New',
        'top': 2,  # 上边框
        # 后面参数是线条宽度
        'left': 2,  # 左边框
        'right': 2,  # 右边框
        'bottom': 2  # 底边框
}
TITLE=[
    (u'序号',8),
    ('IP',22),
    (u'端口',10),
    (u'服务',18),
    (u'开放状态',15),
]



def get_xml(filepath=XMLPATH):
    try:
        return map(lambda x:os.path.join(filepath,x),glob.glob1(filepath,'*.xml'))
    except:
        return []


def test_use_xlsxwriter():
    #创建 excel
    book=xlsxwriter.Workbook('hello.xlsx')
    #创建工作簿
    sheet=book.add_worksheet('FIRST')
    # 添加样式
    ItemStyle = book.add_format({
        'font_size': 10,  # 字体大小
        'bold': True,  # 是否粗体
        'bg_color': '#101010',  # 表格背景颜色
        'font_color': '#FEFEFE',  # 字体颜色
        'align': 'center',  # 居中对齐
        'valign':'vcenter',
        'font_name':'Courier New',
        'top': 2,  # 上边框
        # 后面参数是线条宽度
        'left': 2,  # 左边框
        'right': 2,  # 右边框
        'font_size':16,
        'bottom': 2  # 底边框
    })

    #指定位置写入数据
    sheet.write('A1',u'测试')
    #批量数据写入
    expenses = (['Rent', 1000],
                ['Gas', 100],
                ['Food', 300],
                ['Gym', 50],
                )
    row = 1
    col = 0

    # .write方法  write（行,列,写入的内容,样式）
    for item, cost in (expenses):
        sheet.write(row, col, item)  # 在第一列的地方写入item
        sheet.write(row, col + 1, cost)  # 在第二列的地方写入cost
        row=row + 1  # 每次循环行数发生改变
    sheet.write(row, 0, 'Total')
    sheet.write(row, 1, '=SUM(B2:B4)')  # 写入公式

    #或者使用下面的样式
    # ItemStyle.set_font_size(10)
    # ItemStyle.set_bold()
    # ItemStyle.set_bg_color('#101010')
    # ItemStyle.set_font_color('#FEFEFE')
    # ItemStyle.set_align('center')
    # ItemStyle.set_align('vcenter')
    # ItemStyle.set_bottom(2)
    # ItemStyle.set_top(2)
    # ItemStyle.set_left(2)
    # ItemStyle.set_right(2)

    sheet.write(row+1,0,u'样式设置',ItemStyle)
    sheet.set_column('C:D',12)
    sheet.set_row(3,30)
    sheet.merge_range('C4:D4',u'合并单元格内容',ItemStyle)

    #关闭
    book.close()

# test_use_xlsxwriter()

#获取写入格式
def get_style(default=DEFAULT_STYLE,**kw):
    return  default.update(**kw)



def parseNmap(filename):
    try:
        tree=ET.parse(filename)
        root=tree.getroot()
    except Exception as e:

        print e
        return {}
    data_lst=[]
    for host in root.iter('host'):
        if host.find('status').get('state') == 'down':
            continue
        address=host.find('address').get('addr',None)
        if not address:
            continue
        ports=[]
        for port in host.iter('port'):
            state=port.find('state').get('state','')
            port_num= port.get('portid',None)
            service=port.find('service')
            service= service.get('name','') if service else ""
            ports.append([port_num,service,state])
            data_lst.append({address:ports})
    return data_lst

# filename='report/111.26.138.28.xml'
#
#
# print parseNmap(filename)


def reportEXCEL(filename,datalst,title=TITLE,style=DEFAULT_STYLE,**kwargs):
    if not datalst:
        return ''
    if  os.path.exists(filename):
        print "%s 文件已经存在" % filename
        path,name=os.path.split(filename)
        filename=os.path.splitext(name)[0]
        filename=filename+str(time.strftime("%Y%m%d%H%M%S",time.localtime()))+'.xlsx'
        filename=os.path.join(path,filename)
        print 'data will save as new file named :%s ' % filename

    book=xlsxwriter.Workbook(filename)
    title_style= style if not kwargs.get('title',None) else kwargs.get('title')

    row_hight=[20,16] if not kwargs.get('row_set',None) else kwargs.get('row_set')    #标题题和常规的高度
    # col_width=[8,22] if not kwargs.get('col_set',None) else kwargs.get('col_set') #序号，其他宽度
    sheet_name= 'sheet' if not kwargs.get('sheet_name',None) else kwargs.get('sheet_name')
    sheet=book.add_worksheet(sheet_name)

    row_hight=row_hight+(2000-len(row_hight))*[row_hight[-1]]
    for row , h in enumerate(row_hight):
        sheet.set_row(row,h)
    col_width=map(lambda x:x[1],title)
    for col , w in enumerate(col_width):
        sheet.set_column(col,col,w)
    title_style = book.add_format(title_style)
    for index,t in enumerate(title):

        sheet.write(0,index,t[0],title_style)




    #
    row=1
    col=0
    style=book.add_format(style)
    index2=0
    for index,item in enumerate(datalst):
        for ip,ports in item.items():
            port_num=len(ports)
            if not ports:
                continue
            index2=index2+1
            for  i,data in enumerate(ports):


                sheet.write(row,2,data[0],style)
                sheet.write(row,3,data[1],style)
                sheet.write(row,4,data[2],style)
                row = row + 1
            if row-port_num+1 != row:
                sheet.merge_range('B'+str(row-port_num+1)+':B'+str(row),ip,style)
                sheet.merge_range('A'+str(row-port_num+1)+':A'+str(row),index2,style)
            else:
                print index2
                sheet.write(row-1,0,index2,style)
                sheet.write(row-1,1,ip,style)
    print  'Reprot result of xml parser to file: %s' % filename
    book.close()




def main(XMLPATH,REPORTFILENAME):

    data_lst=[]
    for xml in get_xml(XMLPATH):

        data=parseNmap(xml)
        if data:
            data_lst.extend(data)
            # print data


    reportEXCEL(REPORTFILENAME,data_lst)


















if __name__ == '__main__':
    import sys
    if len(sys.argv)<3:
        print '[!] Usage: parserXML.py XMLPATH [reportfilename]'
        print '[!] Demo: parserXML.py  xmldir  result.xlsx'
    else:

        XMLPATH=sys.argv[1]
        REPORTFILENAME = sys.argv[2]
        print '[-] set parser XML file dir: %s' % XMLPATH
        print '[-] set report Excel file name: %s' % REPORTFILENAME

        if not os.path.exists(XMLPATH):
                print "[!] '%s' path does not exists!" % XMLPATH
                exit(1)
        main(XMLPATH,REPORTFILENAME)




    # main()
    pass