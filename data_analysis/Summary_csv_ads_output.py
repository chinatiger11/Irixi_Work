# -*- coding: UTF-8 -*-
import openpyxl
# import Excel_openpyxl
import pandas as pd
import numpy as np
import os
import re


def find_file(path, ext, file_list=[]):
    dir = os.listdir(path)
    for i in dir:
        i = os.path.join(path, i)
        if os.path.isdir(i):
            find_file(i, ext, file_list)
        else:
            if ext == os.path.splitext(i)[-1].lower(
            ) or ext == os.path.splitext(i)[-1].upper():  #文件后缀名不区分大小写
                file_list.append(i)
    return file_list  #返回文件列表

def get_file_name(path_string):
    """获取path_string路径字符串中的文件名称"""
    pattern = re.compile(r'([^<>/\\\|:""\*\?]+)\.\w+$')
    data = pattern.findall(path_string)
    if data:
        return data[0]


path = r"D:\Work\ROSA\100G_PAM4_ROSA\Result\0609"
Excelname = os.path.join(path, "Summary.xlsx")
ext = '.csv'


filelist = find_file(path, ext, file_list=[])

for dataname in filelist:
    columnname = get_file_name(dataname)
    csvname = os.path.join(path, columnname+".csv")
    df = pd.read_csv(dataname)
    row,col = df.shape
    if col == 2:
        tempdf= df.iloc[:,0]    #取dataframe第一列，该列为频率值
        for i in range(col-1):
            aser = df.iloc[:,i+1].str.replace('/',',')
            aser2 = aser.str.split(',',expand = True)   #将dB/Degrees 用','分割成两个数据
            aser2.columns = ['dB','Degrees']    
            tempdf=pd.concat([tempdf,aser2],axis =1)    #将处理后的数据合并到新的dataframe中
        tempdf.to_csv(csvname, index=False,quoting=0)   #重新写回csv文件
    elif col == 3:
        df_index = df.iloc[:,0]
        column_new = df_index.drop_duplicates() #去除重复的指示变量
        tempdf= df.iloc[:column_new.index[1],[1]]   #取出原dataframe第二列，为频率分量
        aser = df.iloc[:,col-1].str.replace('/',',')    #原始dataframe最后一列为数据
        aser2 = aser.str.split(',',expand = True)   #将dB/Degrees 用','分割成两个数据
        tempdf1_mat = aser2.values      #DataFrame转换成numpy array，方便reshape
        tempdf2_mat = tempdf1_mat[:,0].reshape(column_new.count(),-1)   #取出dB数据进行时reshape处理
        tempdf3_mat = tempdf1_mat[:,1].reshape(column_new.count(),-1)   #取出degree数据进行时reshape处理
        # print(np.hstack((tempdf2_mat.T, tempdf3_mat.T)))
        datadf = pd.DataFrame(np.hstack((tempdf2_mat.T, tempdf3_mat.T)))     #根据变量的扫描次数，reshape矩形，按列进行排列数值
        # print(datadf)
        tempdf=pd.concat([tempdf,datadf],axis =1)    #将处理后的数据合并到新的dataframe中
        db_list = []
        degree_list = []
        for name in column_new:
            db_list.append('dB_'+str(name))
            degree_list.append('degree_'+str(name))
        tempdf.columns = ['Freq'] + db_list + degree_list  #数据列名用 Freq, dB, Degrees 进行指示
        tempdf.to_csv(csvname, index=False,quoting=0)   #重新写回csv文件
        

pass

