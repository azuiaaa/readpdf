import os
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
        return new_name


def readtxt(txt_name=None):
    book = xlwt.Workbook(encoding='utf-8')
    book.add_sheet('sheet1', cell_overwrite_ok=True)
    with open(txt_name, "r") as f:
        # add_test_case = ''
        revised_test_case = ''
        removed_test_case = ''
        add_test_case = []
        # revised_test_case = []
        # removed_test_case = []
        flag_add = 0
        flag_revise = 0
        flag_remove = 0
        lines = f.readlines()
        for line in lines:
            print(line)
            b = re.search('Revised:', line)
            c = re.search('Removed:', line)
            if re.search('Added:', line) is not None or flag_add == 1:
                t = str(line).split(',')
                add_test_case.extend(t)
                # add_test_case.app = add_test_case + str(line)
                if str(line).strip().endswith(','):
                    flag_add = 1
                else:
                    flag_add = 0
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
        # print(add_test_case)
        revised_json = {"added": add_test_case, "revised": revised_test_case, "removed": removed_test_case}
        arr = []
        if revised_json['added']:
            arr = revised_json['added'].split(':')[1].split(',')
        if revised_json['revised']:
            arr.extend(revised_json['revised'].split(':')[1].split(','))
        if revised_json['removed']:
            arr.extend(revised_json['removed'].split(':')[1].split(','))
        arr = [str(x).strip().replace('\\n', '') for x in arr]
        return revised_json


def read_tc_title(txt_name=None, revised_json=None):
    print(revised_json)
    arr = revised_json['revised'].split(',')
    with open(txt_name, "r") as f:
        flag_revise = 0
        revised_tc = {}
        for line in f:
            a = re.search(revised_json['revised'] + ':', line)
            if a is not None or flag_revise == 1:
                title = title + str(line)
                if line.isspace():
                    flag_revise = 0
                    # revised_tc[]
                else:
                    flag_revise = 1

def highlight(select=None, sheet=None):
    df = pd.read_excel('HomeKit用例.xlsx', sheet_name=sheet, index_col=None, header=0,
                       parse_dates=True)  # herder=1：从第2行开始读取
    wb = openpyxl.load_workbook(r'HomeKit用例.xlsx')
    arr = []
    if select['added']:
        arr = select['added'].split(':')[1].split(',')
    if select['revised']:
        arr.extend(select['revised'].split(':')[1].split(','))
    if select['removed']:
        arr.extend(select['removed'].split(':')[1].split(','))
    arr = [str(x).strip().replace('\\n', '') for x in arr]
    testcase = wb[sheet]['D']
    for cellobj in testcase:
        if cellobj.value in arr:
            cellobj.fill = PatternFill('solid', 'fff000')
    wb.save('Homekit用例_1.xlsx')
    for tc in select:
        result_dataframe = df.loc[df['用例编号'].isin(arr)]



if __name__ == '__main__':
    # 获取文件的路径
    # path = open("/Users/zaochuan/Downloads/HomeKit\ Certification\ Test\ Cases\ R11.1.pdf", 'rb')
    path = open('/Users/zaochuan/Downloads/HomeKit Certification Test Cases R11.2.pdf', 'rb')
    # path = path.replace(r'\/'.replace(os.sep, ''), os.sep)
    path_txt = pdftotxt(path, "hk testcase.txt")
    t = readtxt(path_txt)
    read_tc_title(t)
    # highlight(t, "R11.2")
