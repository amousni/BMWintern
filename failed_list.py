# -*- coding: utf-8 -*-
"""
Created on Fri May 31 11:07:50 2019

@author: qxp7153
"""

import pandas as pd
import datetime
import time
import warnings

warnings.simplefilter(action='ignore', category=FutureWarning)

#读取response time数据并生成failed list文件
def failed_list(week_number):
    
    #输入上周假期信息
    #number_of_holiday表示上周假期天数
    #holiday_no_list储存假期为星期几以便于筛选failed list
    number_of_holiday = int(input("Input the number of holidays in last week, 0 for none:"))
    holiday_no_list = []
    if number_of_holiday != 0:
        print('Input holiday in last week(eg: 0501):')
        for i in range(number_of_holiday):
            holiday_str = input('Holiday string:')
            holiday = datetime.date(2019, int(holiday_str[0:2]), int(holiday_str[2:4]))
            holiday_no = holiday.strftime("%w")
            holiday_no_list.append(holiday_no)
    print('-'*50)
    print("Program is reading Response time data...")
    print('-'*50)
    
    #打开response time文件
    #从summary sheet中取出tc/urgent/rr的相应列数据信息
    #从Run8sheet中提取时间信息
    file_name = 'Response time 2019-CW%s.xlsm'%(week_number,)
    summary_df = pd.read_excel(file_name, sheet_name = 'Summary')
    tc = pd.DataFrame(summary_df, columns=['Case no.', 'Team', 'Result.1', 'Urgency'])
    urgent = pd.DataFrame(summary_df, columns=['Case no.', 'Team', 'Result.1', 'Urgency'])
    rr = pd.DataFrame(summary_df, columns=['Case no.', 'Team', 'Result.3', 'Urgency'])
    whole_time_df = pd.read_excel(file_name, sheet_name = 'Run8')
    time_df = pd.DataFrame(whole_time_df, columns = ['Case No.', 'Submit Time'])
    time_df = time_df.rename(columns={'Case No.':'Case no.'})
    print('Finish reading data!')
    print('-'*50)
    
    #将tc/urgent/rr信息与时间信息合并
    tc = tc[(tc['Result.1']=='failed') & (tc['Urgency']=='Preventive')]
    tc = pd.merge(tc, time_df, on='Case no.', how='left')
    urgent = urgent[(urgent['Result.1']=='failed') & (urgent['Urgency']=='Urgent')]
    urgent = pd.merge(urgent, time_df, on='Case no.', how='left')
    rr = rr[(rr['Result.3']=='failed') & (rr['Urgency']=='Reply requested')]
    rr = pd.merge(rr, time_df, on='Case no.', how='left')
    time.sleep(1)
    print('Finish time integration!')
    print('-'*50)
    
    #将tc/urgent/rr中的时间信息以星期几的格式储存在Week Number列中
    #剔除tc/urgent/rr中的周末时间
    df_list = [tc, urgent, rr]
    for i in df_list:
        number = []
        for j in range(i.shape[0]):
            number.append(i.loc[j, 'Submit Time'].strftime("%w"))
        i['Week Number'] = number 
    tc = tc[(tc['Week Number'] != '6')]
    tc = tc[(tc['Week Number'] != '0')]
    urgent = urgent[(urgent['Week Number'] != '6')]
    urgent = urgent[(urgent['Week Number'] != '0')]
    rr = rr[(rr['Week Number'] != '6')]
    rr = rr[(rr['Week Number'] != '0')]
    
    #剔除之前输入的假期时间数据
    #仅保留case no.与team信息
    #更改df index
    if number_of_holiday != 0:
        for i in range(len(holiday_no_list)):
            tc = tc[(tc['Week Number'] != holiday_no_list[i])]
            urgent = urgent[(urgent['Week Number'] != holiday_no_list[i])]
            rr = rr[(rr['Week Number'] != holiday_no_list[i])]  
    tc = pd.DataFrame(tc, columns=['Case no.', 'Team'])
    tc = tc.rename(columns={'Case no.':'TC<2h'})
    tc = tc.set_index(['TC<2h'])
    urgent = pd.DataFrame(urgent, columns=['Case no.', 'Team'])
    urgent = urgent.rename(columns={'Case no.':'URGENT<2h'})
    urgent = urgent.set_index(['URGENT<2h'])
    rr = pd.DataFrame(rr, columns=['Case no.', 'Team'])
    rr = rr.rename(columns={'Case no.':'RR<4h'})
    rr = rr.set_index(['RR<4h'])
    
    #储存failed list数据
    failed_list_file_name = 'response time failed list CW%s.xlsx'%(week_number,)
    writer = pd.ExcelWriter(failed_list_file_name)
    tc.to_excel(writer)
    urgent.to_excel(writer, startcol = 2)
    rr.to_excel(writer, startcol = 4)
    
    #保存
    writer.save()
    print("Finsh failed list analyzing!")
    

def main():
    failed_list('22')

if __name__ == '__main__':
    main()