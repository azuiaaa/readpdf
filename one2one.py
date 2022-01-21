# -*- coding: utf-8 -*
import re
import sys

import openpyxl
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd
import numpy as np
import argparse
import os
import csv


class ReadFrom:
    def __init__(self, version, path_1, path_2):
        self.path = path_1
        # self.data = xlrd.open_workbook(path)
        self.df_1 = pd.ExcelFile(path_1).parse("R11.1")
        self.df_2 = pd.ExcelFile(path_2)

        # self.sheet_name_list = self.df_1.sheet_names
        self.df_1 = self.df_1[self.df_1['用例更新点(Revised)'].notnull()]
        self.select_data_1 = self.df_1.loc[self.df_1['用例更新点(Revised)'].str.contains(version, regex=False)]

        for index, it in self.select_data_1.iterrows():
            # print(item)
            # print('-------------\n')
            for i in self.df_2.sheet_names:
                print(i)
                self.select_data_2 = self.df_2.parse(i)
                print(it["用例编号"])
                arr = self.select_data_2[self.select_data_2["目录层级"].str.contains(it["用例编号"])].index.tolist()
                if len(arr) == 0:
                    self.df_2.parse(i)
                # if index is not None:
                #     row = rows[0]
                #     self.df_2["目录层级"] = it
        self.wb = Workbook()


if __name__ == '__main__':
    path_1 = "./HomeKit用例.xlsx"
    path_2 = "./WWHK用例合集.xlsx"
    version = 'R11.1'
    ReadFrom(version, path_1, path_2)
