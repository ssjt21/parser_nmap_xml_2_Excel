
## 脚本功能

- 批量解析指定文件夹内的所有nmap 扫描结果的xml文件
- 将解析结果导出为excel文件
- 解析内容主要针对存活主机端口，端口开放状态以及端口开放的服务

> 注意1： 文件夹中的XML文件必须以 .xml作为后缀

> 注意2：指定导出文件名尽量以.xlsx作为后缀

## 依赖包

__python 2的环境__

```angular2html
pip install xlsxwriter

```

> 使用教程参见官方文档：https://xlsxwriter.readthedocs.io/format.html

> 用过那么多python的excel模块，感觉这个相对来说学起来和用起来更加的实惠，公式，格式设置，图标，单元格合并应有尽有。


## 使用方法

```pyhton

>> python 脚本.py  指定解析的xml文件夹 指定生成Excel文件的名字

#Demo
>> python parser_nmap_xml_2_excel.py xmldir result.xlsx
```

## 导出Execl展示

![导出Excel的格式]( images/Demoshow.png "导出Excel的格式")


