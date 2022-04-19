# 使用说明
# `方法一：(需要会python或java)`<br>
鉴于当前没有方便提取财报PDF中的财务数据的工具，于是研究了一下各方面资料写了两种语言提取的小工具，即：<br>
- [java版](https://github.com/ARTAvrilLavigne/ExtractFinancialStatement/tree/main/java/ParsePDF)
- [python版](https://github.com/ARTAvrilLavigne/ExtractFinancialStatement/tree/main/python/parsePDF)

备注：若需要提取10页以上的PDF转为excel，可以自行修改代码for循环使用`spire.pdf-3.8.5.jar`提供的方法即可(免费API限制使用10页)<br>

========================================================================<br>
# 更新2021-11-19  
# `方法二：(个人推荐，不用写代码)`<br>
找到一款超级好用，更适合小白的开源PDF提取表格转化excel工具，下载安装即可。刚刚使用一下该工具对PDF中表格提取并转化为excel文件的准确率达到100%<br>
- 使用条件：首先需要安装Java环境，然后下载windows的`tabula-win.zip`安装包解压后双击`tabula.exe`即可~<br>
备注：安装java环境可以自行百度，操作教程太多了。实在不会，我附上一个参考教程链接吧：[win10安装java8](https://blog.csdn.net/JunLeon/article/details/122623465)<br>
* ### [Windows](https://aegis4048.github.io/parse-pdf-files-while-retaining-structure-with-tabula-py)  
  1. Windows & Linux users will need a copy of Java installed. You can download [Java](https://www.java.com/zh-CN/download/) here. (Java is included in the Mac version.)
  
  2. Download `tabula-win.zip` from https://tabula.technology/. Unzip the whole thing
  and open the `tabula.exe` file inside. A browser should automatically open
  to http://127.0.0.1:8080/ . If not, open your web browser of choice and
  visit that link.

  To close Tabula, just go back to the console window and press "Control-C"
  (as if to copy).

========================================================================<br>
# 更新2022-03-24  
# `方法三：(需要会python)`<br>
对于复杂的表格，使用tabula工具提取表格时也会有部分格式混乱。所以找到一款基于tabula-java工具包装的`tabula-py`依赖库<br>
- Github地址: https://github.com/chezou/tabula-py

python环境安装依赖库：`pip install tabula-py`<br>

通过tabula-py依赖库提供的API进行读取PDF提取表格数据，然后按照自己的要求进行清洗即可，开发环境要求如下：<br>
- Java 8+
- Python 3.7+

### Example

tabula-py enables you to extract tables from a PDF into a DataFrame, or a JSON. It can also extract tables from a PDF and save the file as a CSV, a TSV, or a JSON.  

```py
import tabula

# Read pdf into list of DataFrame
dfs = tabula.read_pdf("test.pdf", pages='all')

# Read remote pdf into list of DataFrame
dfs2 = tabula.read_pdf("https://github.com/tabulapdf/tabula-java/raw/master/src/test/resources/technology/tabula/arabic.pdf")

# convert PDF into CSV file
tabula.convert_into("test.pdf", "output.csv", output_format="csv", pages='all')

# convert all PDFs in a directory
tabula.convert_into_by_batch("input_directory", output_format='csv', pages='all')
```

See [example notebook](https://nbviewer.jupyter.org/github/chezou/tabula-py/blob/master/examples/tabula_example.ipynb) for more details. I also recommend to read [the tutorial article](https://aegis4048.github.io/parse-pdf-files-while-retaining-structure-with-tabula-py).
