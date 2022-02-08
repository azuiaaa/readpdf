import re
import sys
import importlib

import openpyxl
import xlwt
import pandas as pd
from openpyxl.styles import PatternFill

importlib.reload(sys)

from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import PDFPageAggregator
from pdfminer.layout import LTTextBoxHorizontal, LAParams
from pdfminer.pdfpage import PDFTextExtractionNotAllowed


# 对本地保存的pdf文件进行读取和写入到txt文件当中


# 定义解析函数
def pdftotxt(path, new_name):
    # 创建一个文档分析器
    parser = PDFParser(path)
    # 创建一个PDF文档对象存储文档结构
    document = PDFDocument(parser)
    # 判断文件是否允许文本提取
    if not document.is_extractable:
        raise PDFTextExtractionNotAllowed
    else:
        # 创建一个PDF资源管理器对象来存储资源
        resmag = PDFResourceManager()
        # 设定参数进行分析
        laparams = LAParams()
        # 创建一个PDF设备对象
        # device=PDFDevice(resmag)
        device = PDFPageAggregator(resmag, laparams=laparams)
        # 创建一个PDF解释器对象
        interpreter = PDFPageInterpreter(resmag, device)
        # 处理每一页
        for page in PDFPage.create_pages(document):
            interpreter.process_page(page)
            # 接受该页面的LTPage对象
            layout = device.get_result()
            for y in layout:
                if (isinstance(y, LTTextBoxHorizontal)):
                    with open("%s" % (new_name), 'a', encoding="utf-8") as f:
                        f.write(y.get_text() + "\n")

def readtxt():
    book = xlwt.Workbook(encoding = 'utf-8')
    book.add_sheet('sheet1', cell_overwrite_ok=True)
    with open("pdfminer.txt", "r") as f:
        add_test_case = ''
        revised_test_case = ''
        removed_test_case = ''
        flag_add = 0
        flag_revise = 0
        flag_remove = 0
        for line in f:
            a = re.search('Added:', line)
            b = re.search('Revised:', line)
            c = re.search('Removed:', line)
            if a is not None or flag_add == 1:
                add_test_case = add_test_case + str(line)
                if str(line).strip().endswith(','):
                    flag_add = 1
                else:
                    flag_add = 0
                print(flag_add)
            if b is not None or flag_revise == 1:
                revised_test_case = revised_test_case + str(line)
                if str(line).strip().endswith(','):
                    flag_revise = 1
                else:
                    flag_revise = 0
            if c is not None or flag_remove == 1:
                removed_test_case = removed_test_case + str(line)
                if str(line).strip().endswith(','):
                    flag_remove = 1
                else:
                    flag_remove = 0
        return {"added": add_test_case, "revised": revised_test_case, "removed": removed_test_case}
        # print(add_test_case,'rrrrrr', removed_test_case, 'rrrrrr',revised_test_case)

def highlight(select=None):
    df = pd.read_excel('HomeKit用例.xlsx', 'R11.1', index_col=None, header=0, parse_dates=True)  # herder=1：从第2行开始读取
    wb = openpyxl.load_workbook(r'HomeKit用例.xlsx')
    arr = select['added'].split(':')[1].split(',')
    arr.extend(select['revised'].split(':')[1].split(','))
    arr.extend(select['removed'].split(':')[1].split(','))
    arr = [str(x).strip().replace('\\n', '') for x in arr]
    print(wb)
    testcase = wb['R11.1']['D']
    # print(set(wb.active.firstHeader))
    for cellobj in testcase:
        if cellobj.value in arr:
            print(cellobj)
            cellobj.fill = PatternFill('solid', 'fff000')
    wb.save('Homekit用例_1.xlsx')
    for tc in select:

        result_dataframe = df.loc[df['用例编号'].isin(arr)]






if __name__ == '__main__':
    # 获取文件的路径
    # path = open("/Users/zaochuan/Downloads/HomeKit\ Certification\ Test\ Cases\ R11.1.pdf", 'rb')
    # pdftotxt(path, "pdfminer.txt")
    t = readtxt()
    highlight(t)

