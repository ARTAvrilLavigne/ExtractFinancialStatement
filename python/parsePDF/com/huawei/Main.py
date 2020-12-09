# -*- coding:utf-8 -*-
from PyPDF2 import PdfFileWriter, PdfFileReader
import os
import re
import xlrd
import xlwt
import time
import subprocess


# 定义全局变量出存放最终的提取数值结果
result_data = []


# 解析excel文件 pip3 install xlrd
def parse_excel(xlsx_path, resultPath):
    # 打开此地址下的excel文档
    data_xsls = xlrd.open_workbook(xlsx_path)
    sheet_list = data_xsls.sheet_names()
    for page_index in range(0, len(sheet_list)):
        # 定义本方法的局部变量
        data = []        # 提取该sheet页全部行列表
        valid_data = []  # 提取>=3符合条件的各行列表

        # 进入第page_index张表
        sheet_name = data_xsls.sheets()[page_index]
        # 获取总行数
        count_nrows = sheet_name.nrows
        # 获得总列数
        count_nocls = sheet_name.ncols
        for i in range(1, count_nrows):
            # 初始化每行的数值列表
            row_data = []
            for j in range(1, count_nocls):
                # 根据行数来取对应列的值，并添加到列表中
                s = sheet_name.cell(i, j).value
                if (isinstance(s, str)) and (s.strip() != ''):
                    row_data.append(s)  # 添加数据到每行列表中
                # print(data_1) # 打印每行
            data.append(row_data)  # 添加每行数据到全局列表中

        # 打印提取该页的所有结果
        # for i in range(len(data)):
        #     print(data[i])
        #
        # print('===============================')

        # 遍历data，删除小于等于长度3的数据，因为这种明显是不符合的
        for index, value in enumerate(data):
            if len(value) >= 3:
                valid_data.append(data[index])
        print('===============================')

        # 打印清洗后的数据
        # for i in range(len(valid_data)):
        #     print(valid_data[i])
        # print('===============================')

        # 处理不同长度的行数值
        for i in range(len(valid_data)):
            # 长度为3的行
            if len(valid_data[i]) == 3:
                # 检验每列属性并保存
                checkColumn(valid_data[i])
            # 长度为4的行
            elif len(valid_data[i]) == 4:
                second_value = str(valid_data[i][1])
                # 第二列数值为数字或者包含字母
                if is_number(second_value) or isCharacter(second_value):
                    valid_data[i].remove(valid_data[i][1])
                    # 检验每列属性并保存
                    checkColumn(valid_data[i])
                else:
                    # 长度为4，第二列不为附注值的情况
                    print('该行长度=4，并且第二行不为字母或数字组合的附注，该值为：', second_value)
            # 长度大于4的行
            else:
                # TODO 后续遇到其余财报此类特殊情况再相应特殊处理即可
                # ['稀释每股收益 /(亏损)', '55', '人民币', '1.22', '元', '人民币(1.67)元']
                special_extract(valid_data[i])

    # 去除首尾的空格以及金额的括号
    deal_result()

    # 打印清洗后的提取数据
    for i in range(len(result_data)):
        print(result_data[i])
    print('In finally,extract valid data count is:', len(result_data))

    # 生成最终结果的excel文件
    write_excel(resultPath)
    print('生成提取数据结果的Excel文件成功!!!')

    # 删除临时转化生成的excel文件
    print('准备删除临时中转解析生成的Excel文件中·······')
    delete_file(xlsx_path)
    print('删除Excel文件成功啦~~~~~~')
    print("$$$$$ complete sophia の magic $$$$$")


# 删除任意一个文件
def delete_file(path):
    # 如果文件存在
    if os.path.exists(path):
        # 删除文件
        os.remove(path)
    else:
        # 文件不存在打印报错信息
        print('no such file exist, please check your excel path!')


# 生成excel文件 pip3 install xlwt
def write_excel(resultPath):
    # 创建工作簿
    workbook = xlwt.Workbook()
    # 创建sheet
    data_sheet = workbook.add_sheet('result')
    # 设置第一列宽度
    first_col = data_sheet.col(0)
    first_col.width = 256*45
    # 设置第二列宽度
    second_col = data_sheet.col(1)
    second_col.width = 256*30
    # 设置第三列宽度
    third_col = data_sheet.col(2)
    third_col.width = 256*30
    # 遍历result_data
    row_index = 0
    for i in range(len(result_data)):
        # 设置第row_index行0/1/2列值
        data_sheet.write(row_index, 0, result_data[i][0], set_style('Times New Roman', 220, True))
        data_sheet.write(row_index, 1, result_data[i][1], set_style('Times New Roman', 220, True))
        data_sheet.write(row_index, 2, result_data[i][2], set_style('Times New Roman', 220, True))
        row_index = row_index + 1

    # 保存文件
    sec_time = int(time.time())  # 增加时间戳
    workbook.save(resultPath + '_' + str(sec_time) + '.xls')


# 设置生成excel样式
def set_style(name, height, bold=False):
    style = xlwt.XFStyle()  # 初始化样式
    font = xlwt.Font()  # 为样式创建字体
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    return style


# 去除多余的空格以及金额的括号
def deal_result():
    for i in range(0, len(result_data)):
        result_data[i][0] = str(result_data[i][0]).replace(' ', '')
        result_data[i][1] = str(result_data[i][1]).replace(' ', '').replace('(', '').replace(')', '')
        result_data[i][2] = str(result_data[i][2]).replace(' ', '').replace('(', '').replace(')', '')


# 特殊处理长度大于4的有效数据行
def special_extract(row_data):
    tmp_data = []
    second_boolean = is_number(row_data[1]) or isCharacter(row_data[1])
    if is_chinese(row_data[0]) and second_boolean:
        tmp_data.append(row_data[0])
        # 后续列包含数字则抽取出来再处理
        for i in range(2, len(row_data)):
            # 判断是否包含特殊字符或者包含数字
            if isSpecialCharacter(row_data[i]) or bool(re.search(r'\d', row_data[i])):
                # 清理多余的中文字符
                new_char = ''
                for ch in row_data[i]:
                    if u'\u4e00' >= ch or ch >= u'\u9fff':
                        # 非中文字符
                        new_char = new_char + ch
                tmp_data.append(new_char)
    if tmp_data:
        # 不为空存在则保存
        result_data.append(tmp_data)


# 校验长度为3的每一列是否符合
def checkColumn(row_data):
    if is_chinese(row_data[0]):
        if (isSpecialCharacter(row_data[1]) or bool(re.search(r'\d', row_data[1]))) and not(is_chinese(row_data[1])):
            if (isSpecialCharacter(row_data[2]) or bool(re.search(r'\d', row_data[2]))) and not(is_chinese(row_data[2])):
                # 保存有效的三列值
                result_data.append(row_data)
    else:
        print("此行不符合要求，row=", row_data)


# 判断是否有特殊字符 逗号与点与负号
def isSpecialCharacter(str):
    # string = "~!@#$%^&*()_+-*/<>,.[]\/"
    string = ",.-"
    for i in string:
        if i in str:
            return True
    return False


# 判断字符串是否包含字母
def isCharacter(str):
    my_re = re.compile(r'[A-Za-z]', re.S)
    res = re.findall(my_re, str)
    if len(res):
        return True
    else:
        return False


# 判断字符串是否包含汉字
def is_chinese(string):
    for ch in string:
        if u'\u4e00' <= ch <= u'\u9fff':
            return True

    return False


# 判断字符串是否为数字
def is_number(s):
    try:
        float(s)
        return True
    except ValueError:
        pass

    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass

    return False


# 实现截取PDF文件指定页数 pip3 install PyPDF2
def split_pdf(start_page, end_page, originPDFPath, tmpPDFPath):
    output = PdfFileWriter()
    pdf_file = PdfFileReader(open(originPDFPath, "rb"))
    pdf_pages_len = pdf_file.getNumPages()

    # 保存input.pdf中的1-5页到output.pdf
    for i in range(start_page, end_page):
        output.addPage(pdf_file.getPage(i))

    outputStream = open(tmpPDFPath, "wb")
    output.write(outputStream)
    return


# 实现解析PDF转化为excel格式文件 TODO 当前API只支持最多十页的转化。若超过数量则可以循环调用！！
# 使用我单独封装的java API接口进行解析转化excel，因为试了多种python开源三方件解析效果极差
# 故使用我封装的解析jar包即可，屏蔽掉本部分java代码解析具体操作对使用者sophia不可见
def parse_pdf(jarPath, parsePDFPath, saveExcelPath):
    command = "java -jar " + jarPath
    arg0 = parsePDFPath
    arg1 = saveExcelPath
    cmd = [command, arg0, arg1]
    new_cmd = " ".join(cmd)
    subprocess.Popen(new_cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE).communicate()
    print('调用java API执行完毕，准备删除临时中转解析生成的pdf文件中·······')

    # 删除临时截取的pdf文件
    delete_file(parsePDFPath)
    print('删除pdf文件成功啦~~~~~~')


# 函数入口放在最后
if __name__ == '__main__':
    # pdf截取开始页
    start_page = 111

    # pdf截取结束页
    end_page = 116

    # 原始PDF财报文件存放路径
    originPDFPath = "C:\\Users\\ThinkPad\\Desktop\\ZTE2019.pdf"

    # 截取PDF临时文件存放路径
    parsePDFPath = "C:\\Users\\ThinkPad\\Desktop\\tmp.pdf"

    # 解析后excel临时文件存放路径
    saveExcelPath = "C:\\Users\\ThinkPad\\Desktop\\tmp.xlsx"

    # java的jar包存放路径
    jarPath = "C:\\Users\\ThinkPad\\Desktop\\ParsePDF.jar"

    # 提取数据保存成excel存放路径 TODO 注意不要带后缀.xlsx或者.xls，只带文件名即可
    # 提取结果excel命名定义规则：文件名_时间戳.xls格式
    resultPath = "C:\\Users\\ThinkPad\\Desktop\\sophiaMagic"

    # 实现截取PDF有效页数
    split_pdf(start_page, end_page, originPDFPath, parsePDFPath)
    # 实现PDF转化excel
    parse_pdf(jarPath, parsePDFPath, saveExcelPath)
    # 实现解析提取excel并生成
    parse_excel(saveExcelPath, resultPath)

