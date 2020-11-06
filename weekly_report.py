# -*- coding: utf-8 -*-
"""
Created on Fri May 24 14:22:40 2019

@author: qxp7153
"""

from escalation_email import escalation_email as ee
from modified_4_pingpong_analyzing import modified_4_pingpong as pp
from xls2xlsx import xls2xlsx as xx
from open_case import open_case as oc
from failed_list import failed_list as fl
import time

#打印基本信息
def print_card_infors_basic_menu():
    title = "For Weekly Report, For Mother Russia!"
    option1 = "1、PuMA open/created/modified zipfiles Process"
    option2 = "2、Modified Cases for Pingpong Analyze"
    option3 = "3、Open Case from Open and Cognos Files"
    option4 = "4、Escalation Email Analyse"
    option5 = "5、Failed List from Response Time"
    option6 = "6、Other options still under programming..."
    option0 = "0、Quit"
    print("="*50)
    print(title.center(50))
    print("-"*50)
    print(option1.center(50))
    print(option2.center(50))
    print(option3.center(50))
    print(option4.center(50))
    print(option5.center(50))
    print(option6.center(50))
    print(option0.center(50))
    print("="*50)

def main():
    while True:
        print_card_infors_basic_menu()
        flag = int(input("Input your option："))
        print('-'*50)
        if flag == 1:
            week_number = input('input week number:')
            xx(week_number)     
        elif flag == 2:
            week_number = input('input week number:')
            print('-'*50)
            print('Please be waiting for about 10.80s...')
            pp(week_number)  
        elif flag == 3:
            week_number = input("input week number:")
            end_day = input("input end day of open case(eg: 0526):")
            oc(week_number, end_day)
        elif flag == 4:
            week_number = input("input week number:")
            ee(week_number)    
        elif flag == 5:
            week_number = input("input week number:")
            print('-'*50)
            print("Make sure Response time file is completed before you analyze failed list!")
            print('-'*50)
            fl(week_number)
        elif flag == 6:
            print('This part has not finished yet...')
            time.sleep(1)
        elif flag == 0:
            break
        else:
            print('Wrong number, once again')
            time.sleep(1)

if __name__ == '__main__':
    main()