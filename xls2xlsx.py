# -*- coding: utf-8 -*-
"""
Created on Fri May 24 15:23:31 2019

@author: qxp7153
"""

import os
import os.path
import win32com.client as win32
import time
import zipfile
import pandas as pd

#extract zip files from PuMA
#xls to xlsx
#concat the xlsx and save as open/created/modified
def xls2xlsx(week_number):

    #桌面路径
    root = 'C:\\Users\\qxp7153\\Desktop\\'   
    l = ['open', 'modified', 'created']
    print('-'*50)
    print("MAKE SURE ALL ZIP FILES ARE DOWNLOADED AND PUT IN RIGHT FILEFOLDERS!")
    print('-'*50)
    
    #解压缩包
    for i in l:
        rootdir = root + i
        for parent, dirnames, filenames in os.walk(rootdir):
            for fn in filenames:
                filedir = os.path.join(parent, fn)
                if filedir.endswith('.zip'):
                    with zipfile.ZipFile(filedir) as zfile:
                        zfile.extractall(path = parent)
    print("zip files were extracted!")
    print('-'*50)
    print("xls files are changing to xlsx!")
    print('-'*50)
    print('44.8s needed for 10MB xls document')
    print('-'*50)
    
    #xls转xlsx
    for i in l:
        rootdir = root + i
        for parent, dirnames, filenames in os.walk(rootdir):
            for fn in filenames:
                starttime = time.time()
                filedir = os.path.join(parent, fn)
                if filedir.endswith('.xls'):
                    print(filedir)
                    excel = win32.gencache.EnsureDispatch('Excel.Application') 
                    wb = excel.Workbooks.Open(filedir) # xlsx: FileFormat=51 # xls: FileFormat=56,
                    wb.SaveAs(filedir.replace('xls', 'xlsx'), FileFormat=51)
                    wb.Close()
                    excel.Application.Quit()
                    endtime = time.time()
                    print(endtime-starttime)
                    print('-'*50)
    print('Finish all xls to xlsx!')
    print('-'*50)
    print("xlsx files are under concat!")
    print('-'*50)
    
    #xlsx合并   
    for i in l:
        file_name = 'CW' + week_number + ' ' + i + '.xlsx'
        a = {}
        df = pd.DataFrame(a) 
        rootdir = root + i
        for parent, dirnames, filenames in os.walk(rootdir):
            for fn in filenames:
                filedir = os.path.join(parent, fn)
                if filedir.endswith('.xlsx'):
                    temp_df = pd.read_excel(filedir)
                    temp_df = temp_df.iloc[:, 0:30]
                    df = pd.concat([df, temp_df], axis=0)
        df.to_excel(file_name)
        
        #打印合并后的数据个数，以确定文件位置以及文件合并是否正确
        data_no = df.shape[0]
        print("%s is saved!"%file_name)
        print("The number of data in %s is %d"%(file_name,data_no))
        print("-"*50)
        
def main():  
    xls2xlsx('23')
    
if __name__ == '__main__':
    main()