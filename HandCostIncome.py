#!/usr/local/bin/python3
# -*- coding:utf-8 -*-

"""处理分地区收入和成本，生成回报率周报表或者月报表"""

__author__="汤永军"

import sys
import os
import pandas as pd
import re
import json

pd.set_option('display.expand_frame_repr', False)


class Income(object):
    """
    收入类

    从后台导出的分地区新增用户收入表一般格式规范，直接使用pandas导入到类中就可以，将所有sheet数据合并到一个字典中，便于使用

    """
    __income_data={}
    #构造函数，将文件路径引入
    def __init__(self):
        list_files = os.listdir(".")
        path = []
        for file in list_files:
            if re.match(r"^[^.].+?用户收入.+",file):
                path.append(file)
                if not os.path.isfile(file):
                    print("income file:",file,"不是文件!")
                    exit()
        if not path:
            print("没有匹配<<用户收入>>关键字的文件")
            exit()
        self.__path = path

    # 按文件路径和sheet名读取文件内容
    def readFile(self):
        # 将整个excel文件转为对象，存入excel变量
        for file in self.__path:
            try:
                excel = pd.ExcelFile(file)
            except :
                print(self.__path,"打开失败")
                exit()
            # 获取sheet名,构建sheet name与其hash值的映射字典
            for sheet_name in excel.sheet_names:
                game_id = matchGameName2Id(sheet_name)
                if not game_id:
                    print(sheet_name,"匹配不到游戏id")
                    exit()
                channel_id = matchChannelId(sheet_name)
                df = excel.parse(sheet_name)
                df.drop(columns=['date'],inplace=True)
                df['Other'] = 0
                other = df.loc[0,'ALL']-df.loc[0,'DE':'xiyu_other'].sum()
                df.loc[0,'Other'] = other
                if game_id not in self.__income_data:
                    self.__income_data[game_id] = {}
                self.__income_data[game_id][channel_id] = df.iloc[0]
        return self.__income_data

    #返回汇率转换之前和之后的收入数据
    def getResult(self):
        origin_data = self.readFile()
        finish_data = {}
        for game_id,channel_data in origin_data.items():
            if game_id not in finish_data:
                finish_data[game_id]={}
            for channel_id,income_data in channel_data.items():
                if channel_id not in finish_data[game_id]:
                    finish_data[game_id][channel_id] = {}
                for country,data in income_data.items():
                    #income_country = country+'收入'
                    income_country = country
                    if currencies[game_id] == "EUR":
                        exchange_rate = exchange_rate_dict['EUR']
                        finish_data[game_id][channel_id][income_country] = float('{:.4f}'.format(data * exchange_rate))
                    else:
                        finish_data[game_id][channel_id][income_country] = float(data)
        return origin_data,finish_data

class Cost(object):
    """成本类
    从当前目录中搜索所有excel文件，包含游戏名的文件当做有效文件

    因为成本表不规范，容易出现脏数据，需要在使用前做清理工作
    将数据导入到类中，处理数据中的脏数据，并返回格式化后的数据

    """

    __cost_data = {}

    def __init__(self):
        file_list = os.listdir(".")
        self.__valid_files = []
        for file in file_list:
            #打开excel文件时，会出现隐藏文件.~文件名这种格式，所以过滤掉
            if not file.startswith(".") and file.endswith(".xlsx") and matchGameName2Id(file) and os.path.isfile(file):
                self.__valid_files.append(file)

    #遍历各个游戏文件
    def readFiles(self):
        for file in self.__valid_files:
            game_id = matchGameName2Id(file)
            if not game_id:
                print(file,"不能匹配到游戏ID")
                exit()
            self.__cost_data[game_id] = {}
            try:
                excel = pd.ExcelFile(file)
            except:
                print("error:",file,"读取失败")
                exit()
            channels = excel.sheet_names
            for channel in channels:
                df = excel.parse(channel)
                df.rename(columns={"USA":"US"},inplace=True)
                #所有值都为空的列删掉
                df = df.dropna(axis=1, how='all')
                #存在值为空的行删掉
                df = df.dropna(axis=0,how='all')
                #检查数据，清除脏数据
                total_index_count,total_col_count = df.shape
                #删除列脏数据
                del_cols = []
                for column in df.columns:
                    column_count = df[column].count()
                    if column_count/total_index_count< 0.1:
                        del_cols.append(column)
                df.drop(columns=del_cols,inplace=True)
                #删除行脏数据
                del_indexes = []
                for index in df.index:
                    cell_df = df.loc[[index]].count(axis=1)
                    index_count = cell_df.values[0]
                    if index_count/total_col_count < 0.1:
                        del_indexes.append(index)
                df.drop(index=del_indexes,inplace=True)
                #匹配sheet name,重命名
                if re.match(r'.*?and.*',channel,flags=re.I):
                    channel = 'Android'
                elif re.match(r'.*?ios.*',channel,flags=re.I):
                    channel = 'IOS'
                else:
                    print("匹配不到安卓或者ios")
                    exit()
                channel_map = {v:k for k,v in channelname_map.items()}
                channel_id = channel_map[channel]
                #保留最后的“合计”行
                df = df.drop_duplicates(df.columns[0],keep="last")
                df.set_index(df.columns[0],inplace=True)
                #获取最后的合计行数据
                ds = df.loc["总计"]
                #将无效列名改为有效列名
                ill_cols= []
                valid_cols = []
                for country in ds.index:
                    if country.find("Unnamed") != -1:
                        ill_cols.append(country)
                    else:
                        valid_cols.append(country)
                ds = ds.drop(labels=valid_cols)
                for key,name in enumerate(valid_cols):
                    ds = ds.rename({ds.index[key]:name})
                #将单次投放的成本之和列改为ALL
                ds = ds.rename({ds.index[len(ds)-1]:"ALL"})
                try:
                    if channel_id == 1:
                        ds['xiyu_other'] = ds.loc['UY':'hn'].sum()
                    else:
                        ds['xiyu_other'] = ds.loc['UY':'ec'].sum()
                except KeyError:
                    ds['xiyu_other'] = 0
                self.__cost_data[game_id][channel_id] = ds

    #获取结果
    def getResult(self):
        self.readFiles()
        return self.__cost_data

#正则匹配游戏名和id
def matchGameName2Id(name):
    for gameName,ids in gamename_map.items():
        if re.match(r'.*?'+gameName+r'.*',name,re.I):
            return ids
    return None

#匹配渠道id
def matchChannelId(channel_str):
    if not channel_str.find('-'):
        return None
    else:
        return int(channel_str.split('-')[1])



#各个游戏的默认货币
currencies = {10001: "EUR", 10002: "EUR", 10003: "EUR", 10004: "USD", 10005: "USD", 10006: "EUR", 10007: "USD",
              10008: "EUR", 10009: "USD", 10011: "USD", 10014: "EUR", 10015: "USD"}
#exchange_rate
exchange_rate_dict = {"EUR":1.174}
#gamename_map = {'VIP': (10003,), '舰队': (10001, 10002), '战舰': (10001, 10002), '沙漠': (10007,), '现代': (10006,),
#                '生物': (10009,), '坦克风云2': (10014,), "黑帮": (10015,), "罪恶": (10015,), "坦克": (10008,)}
gamename_map = {'VIP': 10003, '舰队': 10001, '战舰': 10001, '沙漠': 10007, '现代': 10006,
                '生物': 10009, '坦克风云2': 10014, "黑帮": 10015, "罪恶": 10015, "坦克": 10008}
channelname_map = {1: "Android", 2: "IOS"}

gameid_map = {10001:"超级舰队",10002:"超级舰队",10003:"超级舰队VIP",10006:"现代战争欧洲版",10007:"现代战争北美版",10008:"坦克",10009:"生物",
              10014:"坦克2",10015:"黑帮"}

#主逻辑处理函数
def main():
    income = Income()
    ori_income_data,exchange_income_data = income.getResult()
    #print(ori_income_data,exchange_income_data)
    #exit()

    cost = Cost()
    cost_data = cost.getResult()
    #print(cost_data)
    #exit()
    cost_income_rate = {}
    for game_id,channel_data in exchange_income_data.items():
        if game_id not in cost_income_rate:
            cost_income_rate[game_id] = {}
        for channel_id,income_data in channel_data.items():
            if channel_id not in cost_income_rate[game_id]:
                cost_income_rate[game_id][channel_id] = {}
            for country,data in income_data.items():
                rate_country = country
                try:
                    cost_income_rate[game_id][channel_id][rate_country] = float("{:.4f}".format(data / cost_data[game_id][channel_id][country]))
                except KeyError:
                    cost_income_rate[game_id][channel_id][rate_country] = '='+str(data)+"/0"
                except ZeroDivisionError:
                    cost_income_rate[game_id][channel_id][rate_country] = '='+str(data)+"/0"
    #将原始收入数据格式转换
    ori_income_df = {}
    for game_id,channel_data in ori_income_data.items():
        if game_id not in ori_income_df:
            ori_income_df[game_id] = {}
        for channel_id,income_data in channel_data.items():
            if channel_id not in ori_income_df[game_id]:
                ori_income_df[game_id][channel_id] = {}
            for country,data in income_data.items():
                if country not in ori_income_df[game_id][channel_id]:
                    ori_income_df[game_id][channel_id][country] = []
                ori_income_df[game_id][channel_id][country].append(data)

    returnData = {}
    #为了符合最后输出的格式规范，对ALL和Other做了特殊处理
    for game_id,channel_data in ori_income_df.items():
        if game_id not in returnData:
            returnData[game_id] = {}
        for channel_id,income_data in channel_data.items():
            if channel_id not in returnData[game_id]:
                returnData[game_id][channel_id] = {}
            returnData[game_id][channel_id] = income_data.copy()
            returnData[game_id][channel_id]['欧元美元汇率'] = exchange_rate_dict['EUR']
            for country,data in income_data.items():
                income_country = country+'收入'
                cost_country = country+'成本'
                rate_country = country+'回报率'
                if income_country not in returnData[game_id][channel_id]:
                    returnData[game_id][channel_id][income_country] = []
                if cost_country not in returnData[game_id][channel_id] and country != 'ALL' and country !='Other':
                    returnData[game_id][channel_id][cost_country] = []
                if rate_country not in returnData[game_id][channel_id] and country != 'ALL' and country !='Other':
                    returnData[game_id][channel_id][rate_country] = []
                returnData[game_id][channel_id][income_country].append(exchange_income_data[game_id][channel_id][country])

                if country == 'ALL' or country == 'Other':
                    continue
                try:
                    returnData[game_id][channel_id][cost_country].append(cost_data[game_id][channel_id][country])
                except KeyError:
                    returnData[game_id][channel_id][cost_country].append('0成本')
                returnData[game_id][channel_id][rate_country].append(cost_income_rate[game_id][channel_id][country])
            if 'ALL成本' not in returnData[game_id][channel_id]:
                returnData[game_id][channel_id]['ALL成本']=[]
            if 'ALL回报率' not in returnData[game_id][channel_id]:
                returnData[game_id][channel_id]['ALL回报率']=[]
            try:
                returnData[game_id][channel_id]['ALL成本'].append(cost_data[game_id][channel_id]['ALL'])
            except KeyError:
                returnData[game_id][channel_id]['ALL成本'].append('0成本')
            returnData[game_id][channel_id]['ALL回报率'].append(cost_income_rate[game_id][channel_id]['ALL'])

    writer = pd.ExcelWriter('最新回报率.xlsx',engine='xlsxwriter')
    workbook = writer.book
    cell_format = workbook.add_format({'align': 'center'})
    if 'month' in sys.argv :
        cell_format.set_bg_color('yellow')
        cell_format.set_align('center')
    for game_id,channel_data in returnData.items():
        for channel_id,income_data in channel_data.items():
            df = pd.DataFrame.from_dict(income_data)
            df.to_excel(writer,sheet_name=gameid_map[game_id] + channelname_map[channel_id])
            worksheet = writer.sheets[gameid_map[game_id] + channelname_map[channel_id]]
            worksheet.conditional_format('A2:DF2', {'type': 'cell',
                                                    'criteria': '!=',
                                                    'value': -1,
                                                    'format': cell_format})
            #for country,data in income_data.items():
    writer.save()

#主模块入口
if __name__ == "__main__":
    main()

