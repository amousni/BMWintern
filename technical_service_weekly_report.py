# -*- coding: utf-8 -*-
"""
Created on Wed Jun  5 14:31:02 2019

@author: qxp7153
"""

import pandas as pd

def technical_service(week_number):
    file_name = "Technical Service weekly Report_CW%s.xlsx"%(week_number,)
    df = pd.read_excel(file_name, sheet_name = 'PT')
    df['CW23'] = df['CW22']
    df = pd.DataFrame(df, columns=[' NO. OF CASE BY TEAM','CW23','CW22','CW21','CW20','CW19','CW18'])
    df.loc[9,'CW23'] = 'CW23'
    df.loc[16,'CW23'] = 'CW22'
    df.loc[22,'CW23'] = 'CW23'
    df.loc[28,'CW23'] = 'CW23'
    print(df)
    writer = pd.ExcelWriter(file_name)
    df.to_excel(writer, sheet_name = 'PT')
    writer.save()

def main():
    technical_service('22')
    
if __name__ == '__main__':
    main()