# -*- coding: utf-8 -*-
"""
Created on Fri May 24 17:05:41 2019

@author: qxp7153
"""

import pandas as pd
import re
import datetime

#读取人名信息
def load_name_list():
    df = pd.read_csv('name_list.csv').set_index('name')
    return [list(df.index), df.T.to_dict()]

#department_dic = {'Yang Levi':'DT','Li Keats':'DT','Huang Jeffrey':'PT','Zeng Justin':'DT','qian guoqing':'0','Yang Lei':'DT','Zou Chun':'PT','Yang Jintu':'EE','Guo Galen':'DT','Liu Dongsheng':'PT','hu yuxian':'PT','Zhang Chunhu':'EE','Yu Hailong':'EV','You Dragon':'TT','Xu Wilson':'DT','Si Shengnan':'DT','Wu Hanan':'PT','Zhang Harry':'PT','Cong Huifeng':'EE','Kou Zutao':'DT','Li Wenhao':'DT','China Alpine':'EE','li Yuping':'PT','Li Yin':'DT','Wang Yu':'EE','Liu Junyan':'DT','Wu Steven':'PT','Huang Weiliang':'EE','Ma Marc':'DT','Chen Zhiqiang':'DT','Yang Lei':'DT','Zhang Wayne':'EE','Dong Shixin':'PT','Chen Shawn':'PT','Xiang Baolin':'EE','Lu Sandy':'PT','Jiang Steven':'EE','xing bin':'EE','Wen Zenan':'EE','Li Gang':'EE','Yan Joshua':'PT','Song Colin':'DT', 'Yang Bo':'DT', 'Gao Yongyuan':'DT','Song Min':'DT', 'Li Zhaojun':'DT', 'Lu kejie':'DT', 'Geng Charlie':'DT', 'Tan Lunan':'DT', 'Jin Mingjun':'DT', 'Liu Luno':'DT', 'Wei Yunqian':'EE', 'Fu Rich':'EE', 'Meng Fan Bo':'EE', 'Ren Fei':'EE', 'You Marshal':'EE', 'Wang David':'EE', 'Li Feixue':'EE', 'Qu Zuoqi':'EE', 'Zhou Ke':'EE', 'Zhang Yuan':'PT', 'Shi Wei':'PT', 'Feng Jianquan':'PT', 'Wang Frank':'PT', 'Huang Sean':'PT', 'Ma Ben':'PT', 'Shu Yongcheng':'PT', 'Zhao Shelina':'TT', 'Zheng Annamaria':'TT','Chu Zhaoxian':'TT'}
#name_list = ['guoqing qian','Lei Yang','Dongsheng Liu','yuxian hu','Hailong Yu','Kun Xu', 'Annamaria Zheng', 'xing bin', 'bin xing', 'Chunhu Zhang', 'Don Wang', 'Weiyu Hou', 'Kejie Lu', 'Wenjie Li', 'Hailong Zhao', 'Qi Yang', 'Yongyuan Gao', 'Ruoyun Wang', 'Frank Wang', 'Jian Xu', 'Min Song', 'Ben Ma', 'Jianquan Feng', 'Hu Sun', 'Joshua Yan', 'Raul Bernal', 'Yali Jiang', 'Kai Wang', 'Chengyan Sun', 'Xin Tian', 'Xin Tian', 'Alice Yang', 'Fady Bawatna', 'Li Wang', 'Peter Chen', 'Angela Feng', 'Pengcheng Gao', 'lei Zhang', 'Dongli Liu', 'Lesley lan', 'Zhenhua Li', 'Alexander Lang', 'Zhi Zhang', 'Shuai Ni', 'Donghui Zhang', 'Rui Miao', 'Benzy Qin', 'Ping Hao', 'TB GZ', 'hongxing zhao', 'Rocky Chen', 'Becky Zhang', 'yuan zhou3', 'juan zhang', 'Xin Tian', 'Frank Zheng', 'Ivan Wang', 'Qiang Zhong', 'Guoxiang Deng', 'Alan Zhang', 'Test1 Puma', 'Reader Power User', 'Peter Loertzing', 'Carsten Gruetzmacher', 'Christoph Brettle', 'Tao Zhao', 'Wendy Wu', 'Baolin Xiang', 'Bin Fu', 'Joanna Zhong', 'Mahathi Murthy', 'Rishika Menon', 'Yali Jiang', 'Xiaomao Wu', 'Ting Xu', 'Yali Jiang', 'ruiqing Zhang', 'Yang Li', 'David Li', 'Yu Hou', 'Jing Dai', 'Joanna Zhong', 'Fengjie Lu', 'Jie Tan', 'Sean Huang', 'Dan Wu', 'Fan Bo Meng', 'Anna Yang', 'Fernando Ferrer', 'Rocky Chen', 'Becky Zhang', 'Fei Ren', 'yuan zhou3', 'juan zhang', 'Jie Yang', 'Xin Tian', 'Xin Tian', 'Jintu Yang', 'Andrew Fletcher', 'Yunqian Wei', 'Xiaoxi Liu', 'Alice Yang', 'Bo Yang', 'Fady Bawatna', 'Huifeng Cong', 'Li Wang', 'Peter Chen', 'Marshal You', 'Zuoqi Qu', 'Graham Boyd', 'Jamie Zhang', 'Angela Feng', 'Yanjun Gu', 'Pengcheng Gao', 'Fangchao Xu', 'Pengcheng Gao', 'Yan Yang', 'Lei Li', 'Leo Li', 'Juan Zhang', 'Shixin Dong', 'juan zhang', 'Zhaojun Li', 'Jie Tan', 'David Wang', 'Marc Ma', 'Dawei Wang', 'Dongsheng Liu ', 'Zhengyun Lu', 'Lipeng Wang', 'Lei Liu', 'Yuping li', 'lei Zhang', 'Quan Huang', 'Christian Schoenmeyer', 'Maple Huo', 'Laney Zhang', 'Joachim Fluegge', 'Gordon Zhao', 'Chris Dong', 'Artem Lobanow', 'Zhe Yang', 'Gary Tian', 'Zhicong Xie', 'Dongli Liu', 'Lesley lan', 'Steven Miao', 'Alex Li', 'Zhenhua Li', 'Hao Ling', 'Ao Shen', 'Frank Brandel', 'Mareike Volkmer', 'Arpad Toth', 'Xuqiang Yin', 'Qiang Zhong', 'Andreas Sedlak', 'Ting Jin', 'Zhihai Wu', 'Chun Zou', 'Yuan Zhang', 'Collie Wang', 'Yiding Zhang', 'Guoxiang Deng', 'Season Liu', 'Jonathan Zhu', 'Billy Xiao', 'Yin Li', 'Alan Zhang', 'Ke Chen', 'Justin Jiang', 'Alexander Lang', 'Weiming Hou', 'Gabrielle Ge', 'Billy Fang', 'Wilson Xu', 'Dawei Zhang', 'Zhi Zhang', 'Shuai Ni', 'Joao Martins', 'Wei Shi', 'Gavin Cai', 'dan zhang', 'Justin Zeng', 'Jimmy Ji', 'Spike Zhao', 'Juliet Zhang', 'Jiangtao Chi', 'Steven Jiang', 'Weiliang Huang', 'Ruifeng Liu', 'Ming Liu', 'Yan Ke', 'Vicente Liu', 'Chris Liu', 'Mandy Du', 'Fred Zhang', 'Qinran Luo', 'Harry Zhang', 'Ming Yuan', 'Benzy Qin', 'Jian Ma', 'Walter Gassner', 'Galen Guo', 'Feal Wang', 'Colin Song', 'Fred Feng', 'Ping Hao', 'Jianmin Xin', 'Jarod Yang', 'TB GZ', 'Tony Xu', 'Mingjia Ma', 'Tianshu Li', 'Rich Fu', 'Wayne Zhang', 'Alexander Lang', 'hongxing zhao', 'Christian Kluth', 'Friedrich Mursch', 'Alexander Fischer', 'John White', 'Richard Brown', 'Test1 Puma', 'Reader Power User', 'Thomas Riedl', 'Mark Werner', 'Harald Kammerl', 'Jean Claude Chauveau', 'Frank Schoenrock', 'Stefan Rosse', 'Horst Redner', 'Regan Chen', 'Friedrich Mursch', 'Ian Weston', 'Andreas Herkert', 'Sebastien Beck', 'Thomas Drexler', 'Markus Gamisch', 'Wenjin Chen', 'Richard Wimmer', 'Huashan Wang', 'Yingxue Peng', 'Stefan Hofmann', 'Bob White', 'Christian Mueck', 'Wendi Yi', 'Feng Fred', 'Di Lai', 'Armin Meyer', 'Maoqi Cui', 'Hebe He', 'zhihai wu', 'Sinan Oeztan', 'Yingxue Peng', 'Rex Jiang', 'Frank Zintl', 'Sherry Lu', 'Junli Liang', 'Tim Shi', 'Yuhuo Ma', 'lilong zhao', 'Allen Mao', 'Hailong Li', 'Shawn Chen', 'Leiming He', 'Ermera Setyadi', 'Kimreol Chen', 'Stefan Krott', 'Xiaohui Li', 'Liping Hong', 'Jay Lee', 'Christian Klenovsky', 'Feixue Li', 'Jeffrey Jiang', 'Shelina Zhao', 'Zhang Zhe', 'Jian Guo', 'weiqiang mu', 'Sabrina Fischer', 'Zhi Cui', 'Ryan Qian', 'Zenan Wen', 'Yinbo Shen', 'Zhaoxian Chu', 'Dragon You', 'Wenhao Li', 'Zutao Kou', 'Lunan Tan', 'Charlie Geng', 'Sandy Lu', 'Ke Zhou', 'Haoxin Liu', 'Gang Li', 'Mingjun Jin', 'Levi Yang', 'Junyan Liu', 'Yu Wang', 'Zhiqiang Chen', 'Yongcheng Shu', 'Shengnan Si', 'Lei Yang ', 'Bin Xing', 'Hanan Wu', 'Jack Zhang']


def recom_pair_for_opencase(recom_str):

    unfinished_name_list = load_name_list()[0]
    name_list = []
    for i in unfinished_name_list:
        last_first = i.split(' ')
        name = last_first[-1] + ' ' + last_first[0]
        name_list.append(name)
    
    #时间节点，当天至半年前
    end_day = datetime.date.today()
    start_day = end_day - datetime.timedelta(days=365)
    
    #正则匹配回复时间与回复者人名
    pattern = re.compile('Recommendation   (.*?)   .*?   (.*? .*?) ')
    
    #匹配信息，一个元组列表，元组中第一个元素为回复时间，第二个元素为回复人名
    s = pattern.findall(recom_str)
    
    #人名列表
    recom_name_list = []
    
    #时间匹配计数
    flag = 0
    for i in s:
        mdy_str = i[0].split('/')
        mdy = datetime.date(int(str(20) + mdy_str[2]),int(mdy_str[0]),int(mdy_str[1]))
        if mdy <= end_day and mdy >= start_day:
            recom_name_list.append(i[1])
            flag += 1
            
    #Additional information正则匹配
    pattern_a = re.compile('Additional information   (.*?)   .*?   (.*? .*?) ')
    s_a = pattern_a.findall(recom_str)
    for i in s_a:
        mdy_str = i[0].split('/')
        mdy = datetime.date(int(str(20) + mdy_str[2]),int(mdy_str[0]),int(mdy_str[1]))
        if mdy <= end_day and mdy >= start_day:
            if i[1] in name_list:
                recom_name_list.append(i[1])
                flag += 1
                
    #正则匹配到的人名是名+姓，department字典中是姓+名，需要调换一下顺序
    #如果正则匹配到人名，则取正则中的第一个匹配对象和最后一个匹配对象，否则name取0
    name = []
    if len(recom_name_list)>0:
        name0list = recom_name_list[0].split(' ')
        name0 = name0list[-1] + ' ' + name0list[0]
        name.append(name0)
        name1list = recom_name_list[0].split(' ')
        name1 = name1list[-1] + ' ' + name1list[0]
        name.append(name1)        
    else:
        name = [0]
    
    #返回两个人名的意义在于，department字典中的人名可能不全
    #如果第一个不能在字典中匹配，则匹配最后一个，如果正则中只匹配到一个，则name中的两个人名相同
    #返回的name为一个列表，如果匹配到人名，则name的长度为2，如果未匹配到，则长度为1，元素为0            
    return name

#open case analyzing
def open_case(week_number, end_day):
    
    department_dic = load_name_list()[-1]  
    
    #file_name为输出的open case结果
    #cognos为从1月1号到end_day的cognos数据
    #open_file为本周的open case
    file_name = 'Open Cases from 20190101-2019' + end_day + '.xlsx'
    cognos = 'Case and Vehicle Details open 20190101-2019' + end_day + '.xlsx'
    open_file = 'CW' + week_number + ' open.xlsx'
    
    #读取cognos数据中的case id和urgency两列
    cognos_df_whole = pd.read_excel(cognos)
    cognos_df = pd.DataFrame(cognos_df_whole, columns=['Case Id', 'Urgency'])
    
    #打开open case
    open_df_whole = pd.read_excel(open_file)
    
    #提取open case中case no.和recommendation两列
    #将recommendation列改名
    #将open_df中的Case no.改名为和Cognos中一样的Case Id，目的是将open_df与cognos_df按相同的列合并
    open_df = pd.DataFrame(open_df_whole, columns=['Case no.', 'Previous recommendations/queries/additional information', 'Remarks'])
    open_df = open_df.rename(columns={'Case no.':'Case Id', 'Previous recommendations/queries/additional information':'Recommendation'})
    
    #将open case中的空值数据填为0
    open_df = open_df.fillna(0)
    
    #open case数据个数
    row = open_df.shape[0]
    
    #department列表，也可以改进为open_df['depart'] = []
    depart = []
    
    #对open case数据进行循环
    for i in range(row):
        
        #判断是否由recommendation数据，没有depart添加0
        if open_df.loc[i, 'Recommendation'] != 0:
            
            #正则匹配recommendation中的回复者信息
            name = recom_pair_for_opencase(open_df.loc[i, 'Recommendation'])
            
            #匹配到人名则len(name)>1
            if len(name) > 1:
                
                #判断第一个匹配到的人名是否在department字典中
                #如果不在，则用最后一个匹配到的人名
                #如果两个人名都不在，则depart添加-1
                if name[0] in department_dic.keys():
                    depart.append(department_dic[name[0]]['department'])
                else:
                    if name[1] in department_dic.keys():
                        depart.append(department_dic[name[1]]['department'])
                    else:
                        depart.append('-1')
            else:
                depart.append('0')
        else:
            depart.append('0')
            
    #把depart赋值给open_df
    open_df['team'] = depart
    
    #筛选remarks不为B1的数据，remarks为B1表示不需要回复，需要剔除
    open_df = open_df[open_df.Remarks != 'B1']
    open_df = pd.DataFrame(open_df, columns=['Case Id', 'Recommendation', 'team'])
    open_df = open_df.rename(columns={'Urgency.':'urgency'})
    
    #将open_df与按Case Id合并
    df = pd.merge(open_df,cognos_df,on='Case Id',how='left')
    
    #df中出现空值则填0
    df = df.fillna(0)
    
    #保存文件
    df.to_excel(file_name)
    print('-'*50)
    print('Finish Open case analyzing!')

def main():
    open_case('17', '0421')

if __name__ == '__main__':
    main()