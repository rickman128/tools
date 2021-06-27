# -*- coding: utf-8 -*-

import tempfile
import xlrd
import msoffcrypto
import pandas as pd

def open_workbook(workbook_file_path, password=None):
    if password is None:
        wb = xlrd.open_workbook(workbook_file_path)
    else:
        with open(workbook_file_path, 'rb') as fin,\
                tempfile.TemporaryFile() as tfp:
            encrypted = msoffcrypto.OfficeFile(fin)
            encrypted.load_key(password=password)
            encrypted.decrypt(tfp)
            tfp.seek(0)
			# xlsxがサポートされていない？
			# xlrdを2.0.1→1.2.0にすれば動く
            wb = xlrd.open_workbook(file_contents=tfp.read())
    return wb

wb_a = open_workbook('34.xlsx', '31093109')
wb_b = open_workbook('35.xlsx', '31093109')
df_a = pd.read_excel(wb_a)
df_b = pd.read_excel(wb_b)
# mergeを調べる
df_dif = pd.merge(df_a, df_b, on='A', how='outer', indicator=True)
print(df_dif)
