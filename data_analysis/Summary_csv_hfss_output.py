# -*- coding: UTF-8 -*-
import openpyxl
# import Excel_openpyxl
import pandas as pd
import os
import re


def find_file(path, ext, file_list=[]):
    dir = os.listdir(path)
    for i in dir:
        i = os.path.join(path, i)
        if os.path.isdir(i):
            find_file(i, ext, file_list)
        else:
            if ext == os.path.splitext(i)[1].lower(
            ) or ext == os.path.splitext(i)[1].upper():  #文件后缀名不区分大小写
                file_list.append(i)
    return file_list  #返回文件列表

def get_file_name(path_string):
    """获取path_string路径字符串中的文件名称"""
    pattern = re.compile(r'([^<>/\\\|:""\*\?]+)\.\w+$')
    data = pattern.findall(path_string)
    if data:
        return data[0]


path = r"X:\Work\COC_HHI_NEO\Result\0128"
Excelname = os.path.join(path, "Summary.xlsx")
ext = '.csv'

# xls = Excel_openpyxl.PyEasyExcel(r'C:\Users\huahu.shi\OneDrive - Irixi Photonics, Inc\Program\Python\Data_summary\Summary.xlsx')     

filelist = find_file(path, ext, file_list=[])

TDR_time_flag = True
s11_freq_flag = True
s21_freq_flag = True
s22_freq_flag = True

for dataname in filelist:
    columnname = get_file_name(dataname)
    if 'TDR' in columnname:
        df = pd.read_csv(dataname)
        for colname in  df.columns:
            if 'Time' not in colname:
                df.drop(colname,axis=1,inplace=True)
            break
        row,col = df.shape
        if col == 2:    #one columnn for Frequency another for data
            df.rename(columns={df.columns[1]:columnname}, inplace = True)   #列名替换为文件名
        if TDR_time_flag:
            df_tdr = df.copy()
            TDR_time_flag = False
        else:
            df_tdr = pd.merge(df_tdr,df,how = 'left', on=df_tdr.columns[0])
            # print(df_tdr)
            # df_tdr = pd.concat(df_tdr,df,axis=1)

    if 'S11' in columnname:
        df = pd.read_csv(dataname)
        for colname in  df.columns:
            if 'Freq' not in colname:
                df.drop(colname,axis=1,inplace=True)
            break
        row,col = df.shape
        if col == 2:    #one columnn for Frequency another for data
            df.rename(columns={df.columns[1]:columnname}, inplace = True)   #列名替换为文件名
        if s11_freq_flag:
            df_s11 = df.copy()
            s11_freq_flag = False
        else:
            df_s11 = pd.merge(df_s11,df,on=df.columns[0])
    
    if 'S21' in columnname:
        df = pd.read_csv(dataname)
        for colname in  df.columns:
            if 'Freq' not in colname:
                df.drop(colname,axis=1,inplace=True)
            break
        row,col = df.shape
        if col == 2:    #one columnn for Frequency another for data
            df.rename(columns={df.columns[1]:columnname}, inplace = True)   #列名替换为文件名
        if s21_freq_flag:
            df_s21 = df.copy()
            s21_freq_flag = False
        else:
            df_s21 = pd.merge(df_s21,df,on=df.columns[0])
    
    if 'S22' in columnname:
        df = pd.read_csv(dataname)
        for colname in  df.columns:
            if 'Freq' not in colname:
                df.drop(colname,axis=1,inplace=True)
            break
        row,col = df.shape
        if col == 2:    #one columnn for Frequency another for data
            df.rename(columns={df.columns[1]:columnname}, inplace = True)   #列名替换为文件名
        if s22_freq_flag:
            df_s22 = df.copy()
            s22_freq_flag = False
        else:
            df_s22 = pd.merge(df_s22,df,on=df.columns[0])

writer = pd.ExcelWriter(Excelname)
df_tdr.to_excel(writer,'TDR',index=None)
df_s11.to_excel(writer,'S11',index=None)
df_s21.to_excel(writer,'S21',index=None)
df_s22.to_excel(writer,'S22',index=None)
writer.save()

pass

