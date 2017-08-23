# python-excel-count

考勤表格筛选统计

说明：
使用前先把考勤表格放到 sourceXls 文件夹中，
然后再运行程序，避免出错。
运行后生成的文件放在 destinationXls 文件夹中。

注意：
请保证 sourceXls 文件夹中存在Excel文件，且表格格式与考勤表格相同。
表格格式可参看 templates 文件夹中的Excel模板。

因为暂时只能读取一个表格,
所以请保持 sourceXls 文件夹中只存在一个你要统计的Excel表格。
如果 sourceXls 文件夹中存在多个表格，也只会读取最新的表格。

因为脚本只统计Excel表格中的第一个表,
所以请把统计的数据放在Excel表格中的第一个表格。

要先把表格按名称排好顺序,否则统计有问题。

支持的Excel格式为 .xls 和 .xlsx

最后更新：ver1.0 2017-08-21
