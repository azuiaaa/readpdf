import os
import re
from io import StringIO

# import numpy as np
import openpyxl
from openpyxl import load_workbook

from copy import copy
import pandas as pd
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.utils import get_column_letter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfinterp import PDFResourceManager, process_pdf


def read_pdf(pdf):
    # resource manager
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    laparams = LAParams()
    # device
    device = TextConverter(rsrcmgr, retstr, laparams=laparams)
    process_pdf(rsrcmgr, device, pdf)
    device.close()
    content = retstr.getvalue()
    retstr.close()
    # 获取所有行
    lines = str(content).split("\n")
    new_name = 'testcase.txt'
    # if os.path.exists(new_name): os.remove(new_name)
    with open("%s" % (new_name), 'w', encoding="utf-8") as f:

        for line in lines:
            """移除版权行"""
            pattern = re.compile(r'\d+.*copyright.*\.', re.I)  # 注意用4个\\\\来替换\
            if pattern.match(line):
                lines.remove(line)
                continue
            """移除页码行"""
            if re.compile(r'^\d+$').match(line):
                lines.remove(line)
                continue
            """将每条用例结尾处加上感叹号方面"""
            if re.compile(r'TC\S+\d+\s').match(line) or re.compile(r'chapter', re.I).match(line):
                print(line)
                f.write('!!!\n')

            f.write(line + '\n')
    return lines


def find_revised(lines):
    flag_add = 0
    added = []
    flag_revise = 0
    revised = []
    flag_remove = 0
    removed = []
    for line in lines:

        if re.search('Added:', line) is not None or flag_add == 1:
            added.extend(
                [x.strip().split(' ')[2] if len(x.strip().split(' ')) > 1 else x.strip() for x in line.split(',')])
            flag_add = 1 if line.strip().endswith(',') else 0
            continue

        if re.search('Revised:', line) is not None or flag_revise == 1:
            revised.extend(
                [x.strip().split(' ')[2] if len(x.strip().split(' ')) > 1 else x.strip() for x in line.split(',')])
            flag_revise = 1 if line.strip().endswith(',') else 0
            continue

        if re.search('Removed:', line) is not None or flag_remove == 1:
            removed.extend(
                [x.strip().split(' ')[2] if len(x.strip().split(' ')) > 1 else x.strip() for x in line.split(',')])
            flag_remove = 1 if line.strip().endswith(',') else 0
            continue

    return {'added': added, 'revised': revised, 'removed': removed}


def read_tc_title(excel_path=None, old_sheet=None, sheet_name=None, revised_json=None, testcase_column_name="用例编号"):
    global f, excel_obj, writer
    txt_name = 'testcase.txt'

    # sheet = data.active
    revised_tc = []
    try:
        # """创建一个复制表"""
        new_sheet_name = sheet_name + "（更新中）"
        #
        # book = openpyxl.load_workbook(excel_path)  # 打开工作簿
        # if new_sheet_name in book.get_sheet_names():
        #     book.remove(book[new_sheet_name])
        # copy_sheet = book.copy_worksheet(book[sheet_name])
        # copy_sheet.title = new_sheet_name
        #
        # book.save(excel_path)
        # book.close()

        """需要用pandas去进行用例部分内容的替换，首先需要获取到一个用例内容的dataframe"""
        excel_obj = pd.ExcelFile(excel_path, engine='openpyxl')
        df = pd.read_excel(excel_path, sheet_name=old_sheet)

        # 从解析出来的txt中获取到所有行
        f = open(txt_name, "r", encoding="UTF-8", errors="ignore")
        lines = f.readlines()

        """更新增加的用例"""
        for tc in revised_json["added"]:
            result = analyze_test_case(lines, tc)
            pattern = re.match(r"(\D+)(\d+)", tc)
            print("{}, {}".format(pattern.group(1), pattern.group(2)))
            # process = df[df["用例编号"].str.contains(pattern.group(1) + "\d+", regex=True) == pattern.group(1)]
            # row_tc = process[df["用例编号"] > tc]

            flag = df["用例编号"].str.contains(pattern.group(1)+"\d+", regex=True)
            process = df[flag]
            print("!!!!!!!!!!!!!!!,{}".format(df.index[flag]))
            num = df.index[flag]
            if len(num) > 0:
                insert_id = num[-1]
                row_tc = df.index[process[tc < process["用例编号"]]]
                if len(row_tc) > 0:
                    insert_id = row_tc[0]
            else:
                insert_id = 0
            df_add = pd.DataFrame({"用例编号": tc, "用例标题": result["title"], "适用设备": result["applies"], "英文步骤": result["content"], "用例更新点(Revised)": sheet_name + "新增"})
            df1 = df.iloc[:insert_id, :]
            df2 = df.iloc[insert_id:, :]
            df = pd.concat([df1, df_add, df2], ignore_index=True)

            print(row_tc)

            # df.apply()

        """更新更新的用例"""
        # f = open(txt_name, "r", encoding="gb18030", errors="ignore")  # 会导致出现中文乱码
        for tc in revised_json['revised']:
            result = analyze_test_case(lines, tc)

            """使用pandas获取单元格行列，并替换dataframe的值"""
            row_tc = df.index[df["用例编号"] == tc].tolist()  # this will only contain 2,4,6 rows
            if len(row_tc) > 0:
                df.at[row_tc, "用例标题"] = result.get('title')
                df.at[row_tc, "适用设备"] = result.get('applies')
                df.at[row_tc, "英文步骤"] = result.get('content')
                df.at[row_tc, "用例更新点(Revised)"] = sheet_name + "更新"

        """更新移除的用例"""
        for tc in revised_json['removed']:
            """使用pandas获取单元格行列，并替换dataframe的值"""
            row_tc = df.index[df["用例编号"] == tc].tolist()  # this will only contain 2,4,6 rows
            if len(row_tc) > 0:
                df.at[row_tc, "用例更新点(Revised)"] = sheet_name + "移除"

        # nrows = sheet.max_row  # 获得行数
        # ncolumns = sheet.max_column  # 获得列数
        # for row in df_tc.iterrows():
        #     row
        # sheet.cell(nrows + 1, 1).value = tc
        # sheet.cell(nrows + 1, 2).value = title
        # sheet.cell(nrows + 1, 3).value = case_content
        # # data.save('Homekit用例_1.xlsx')
        book = openpyxl.load_workbook(excel_path)
        writer = pd.ExcelWriter(excel_path, engine="openpyxl")
        print('{}, {}'.format(excel_path, type(excel_path)))
        writer.book = book
        if new_sheet_name in writer.book.sheetnames:
            writer.book.remove(writer.book[new_sheet_name])
        df.to_excel(excel_writer=writer, sheet_name=new_sheet_name, index=None)
        writer.save()
        book.close()


    finally:
        f.close()
        # writer.close()
        excel_obj.close()
    return revised_tc


def analyze_test_case(lines, testcase_name) -> dict:
    result = dict()

    flag_title = flag_tc = flag_applies = flag_empty = 0
    title = ''
    case_content = ''
    applies = ''

    it = iter(lines)
    for line in it:
        if line.find(testcase_name + ' ') > -1 or flag_title == 1 or flag_applies == 1 or flag_tc == 1:

            if line.find(testcase_name + ' ') > -1:
                flag_title = 1
                title = title + line.replace(testcase_name + ' ', '')
                continue
            if flag_title == 1:
                title = title + line
                if line.strip() == '':
                    flag_applies = 1
                    flag_title = 0
                    continue
                # it.__next__()
            # elif line.find('Applies to') > -1:
            if flag_applies == 1:
                if line.strip() == '':
                    flag_applies = 0
                    flag_tc = 1
                    continue
                else:
                    applies = applies + line
                    continue
            if flag_tc == 1:
                if line.strip() == '':
                    flag_empty = flag_empty + 1
                    continue
                if line.startswith('!!!') or flag_empty > 1:
                    flag_tc = 0
                    break
                else:
                    case_content = case_content + line
                    flag_empty = 0

    """将部分不是常用字符的内容删除"""
    title = ILLEGAL_CHARACTERS_RE.sub(r'', title)
    case_content = ILLEGAL_CHARACTERS_RE.sub(r'', case_content)
    applies = ILLEGAL_CHARACTERS_RE.sub(r'', applies)

    return {'tc': testcase_name, 'title': title, 'applies': applies, 'content': case_content}


if __name__ == '__main__':
    with open('./HomeKit Certification Test Cases R11.1.pdf', "rb") as my_pdf:
        read_tc_title("./HomeKit用例.xlsx", "R11", "R11.1", revised_json=find_revised(read_pdf(my_pdf)))
