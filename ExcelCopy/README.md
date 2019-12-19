# 基于Python 3.6实现的对Excel操作功能.
主要是解决了自己的一个经常要重复的工作

## 1.主要功能:
利用Python进行对Excel表格的复制，写入，加入公式，保持原样式
##  2.实现
使用xlrd xlwt xlutils
## 当中遇到的坑
- 写入文件会导致失去原有样式
解决:复制两份原表格。1份用来操作，一份用来对照

```
# rb = xlrd.open_workbook(name, formatting_info=True, on_demand=True)
# wb, s = copy2(rb)
# wbs = wb.get_sheet(0)
# rbs = rb.get_sheet(0)
# styles = s[rbs.cell_xf_index(0, 0)]
# rb.release_resources()  #关闭模板文件
# wbs.write(0, 0, 'aa', styles)
# wb.save("2.xls")
```
复制时formatting_info=True会连同样式问题一起复制过来，直接复制会没有样式
利用.cell_xf_index的属性去获取样式，写入时也加入同样的样式
- 公式写入问题

```
Newexcelsheet.write(hang, 11, xlwt.Formula(gongshi),Copyexcel[Copyexcels.cell_xf_index(hang, 11)])
```
有的地方需要使用此代码置入公式
- 操作必须用.xls格式的Excel 其他的支持性不好会报错
- 导出时注意包是否会一同导出 Pycharm的包可能会找不到位置所以导出时要加入python目录下的第三方库模板目录路径 site-packages。

```
pyinstaller example.py -F -p C:/python/lib/site-packages  例子

pyinstaller ExcelCopy.py -F -p  F:\text\Lib\site-packages  实例

```


