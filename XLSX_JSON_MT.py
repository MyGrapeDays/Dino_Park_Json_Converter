#!/usr/bin/env python
# coding: utf-8

# In[28]:


import numpy as np
import pandas as pd
import sys
import time

import json
import os
import math
import re


# In[29]:


LtoN = {'a': 1,'b': 2,'c': 3,'d': 4,'e': 5,'f': 6,'g': 7,'h': 8,'i': 9,'j': 10,'k': 11,'l': 12,'m': 13,'n': 14,'o': 15,'p': 16,'q': 17,'r': 18,'s': 19,'t': 20,'u': 21,'v': 22,'w': 23,'x': 24,'y': 25,'z': 26,'A': 1,'B': 2,'C': 3,'D': 4,'E': 5,'F': 6,'G': 7,'H': 8,'I': 9,'J': 10,'K': 11,'L': 12,'M': 13,'N': 14,'O': 15,'P': 16,'Q': 17,'R': 18,'S': 19,'T': 20,'U': 21,'V': 22,'W': 23,'X': 24,'Y': 25,'Z': 26}


def letters_to_number(letter):
    sum = 0
    for i in range(len(letter)):
        sum = sum + LtoN[letter[len(letter) - i - 1]]*math.pow(26, i)
    return int(sum - 1)

#переводим A1:A1 нотацию в кортеж(начало, конец)
def get_rows_from_a1(A1_request):
    try:
        a1_notation = re.match("[a-zA-Z]{1,}:[a-zA-Z]{1,}", A1_request).group(0)
    except:
        print("Bad A1_notation request")
    
    first_row = a1_notation[:re.search(":",a1_notation).start()]
    last_row = a1_notation[re.search(":",a1_notation).start()+1:] 
    return(letters_to_number(first_row), letters_to_number(last_row))

#переводим A1нотацию в кортеж(номер столбца, номер строки)
def get_cell_from_a1(A1_request):
    try:
        a1_notation = re.match("[a-zA-Z]{1,}([1-9][0-9]{1,}|[0-9])", A1_request).group()
        col_num = letters_to_number(a1_notation[:re.search("[a-zA-Z]{1,}",a1_notation).end()])
        row_num = int(a1_notation[re.search("[a-zA-Z]{1,}",a1_notation).end():])
    except:
        print('Bad A1_notation request')
    return (row_num, col_num)

#проверяем любой объект на nan
def is_float_nan(obj):
    if type(obj) != float: 
        return False
    return np.isnan(obj)

#делаем форматированный df из df, взятого из excel
# df_worksheet - неформатированный df, A1_request — диапазон столбцов в A:A нотации, offset — смещение по строкам
def get_df_from_worksheet(df_worksheet, A1_request, offset = 1, is_float = True): 
    col_range = get_rows_from_a1(A1_request)
    df_result = df_worksheet.iloc[offset:, col_range[0]:col_range[1]+1]
    df_result = df_result.loc[ df_result.iloc[0:,0].apply( lambda x: not(is_float_nan(x))) ]
    
    df_result.columns = df_result.iloc[0]
    df_result = df_result[1:]
    df_result.index = range(0, len(df_result))
    
    if is_float:
        df_result = df_result.astype('float64')
    return df_result

#берем содержимое клетки из df, взятого из excel
# df_worksheet - неформатированный df, A1_request — номер клеки в A1 нотации, offset — смещение по строкам
def get_cell_from_worksheet(df_worksheet, A1_request, offset = -2):
    return df_worksheet.iloc[offset + get_cell_from_a1(A1_request)[0]][get_cell_from_a1(A1_request)[1]]


# In[30]:

# ФУНКЦИЯ ЗАПОЛНЯЕТ .json МОНСТР-БИЗНЕСОВ 
def businesses_json(df_worksheet, businesses_name):
    df1 = get_df_from_worksheet(df_worksheet, "X:Z", offset = 1)
    df2 = get_df_from_worksheet(df_worksheet, "AD:AH", offset = 1)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/Businesses/' + businesses_name + '/' + businesses_name + 'GradesLevels.json', orient = 'records', indent = 4)
    
    
    df = get_df_from_worksheet(df_worksheet, "AP:AV", offset = 1)
    df.to_json('JSON/GameData/Businesses/' + businesses_name + '/' + businesses_name + 'IslandsReqs.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "I:I", offset = 1)
    df2 = get_df_from_worksheet(df_worksheet, "M:O", offset = 1)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/Businesses/' + businesses_name + '/' + businesses_name + 'Levels.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "G:H", offset = 1)
    df2 = pd.DataFrame( [[get_cell_from_worksheet(df_worksheet, "C4"), get_cell_from_worksheet(df_worksheet, "C5"), get_cell_from_worksheet(df_worksheet, "C6"),
                          get_cell_from_worksheet(df_worksheet, "C7"), get_cell_from_worksheet(df_worksheet, "C8"), get_cell_from_worksheet(df_worksheet, "C9"), 
                          get_cell_from_worksheet(df_worksheet, "C10"),
                          get_cell_from_worksheet(df_worksheet, "BD4"), get_cell_from_worksheet(df_worksheet, "BE4"), get_cell_from_worksheet(df_worksheet, "BF4"), 
                          get_cell_from_worksheet(df_worksheet, "BH4"), get_cell_from_worksheet(df_worksheet, "BI4"),
                          get_cell_from_worksheet(df_worksheet, "AL4"), get_cell_from_worksheet(df_worksheet, "AL5"), get_cell_from_worksheet(df_worksheet, "AL6"),
                          get_cell_from_worksheet(df_worksheet, "AL7"), get_cell_from_worksheet(df_worksheet, "AL8"), get_cell_from_worksheet(df_worksheet, "AL9")]],
                          
                          columns = ['BusinessScaleChange0', 'BusinessScaleChange1', 'BusinessScaleChange2', 
                                     'BusinessScaleChange3', 'BusinessScaleChange4', 'BusinessScaleChange5', 
                                     'BusinessScaleChange6',
                                     "BusinessObservationDeck1", "BusinessObservationDeck2", "BusinessObservationDeck3", 
                                     "BusinessServiceManagerWalkTimeKitchen", "BusinessServiceManagerWalkTimeClean",
                                     "BusinessIsland1LevelLimit", "BusinessIsland2LevelLimit", "BusinessIsland3LevelLimit",
                                     "BusinessIsland4LevelLimit", "BusinessIsland5LevelLimit", "BusinessIsland6LevelLimit"])
    
    
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/Businesses/' + businesses_name + '/' + businesses_name + 'Params.json', orient = 'records', indent = 4)
              
        
    df1 = get_df_from_worksheet(df_worksheet, "T:V", offset = 1)
    df2 = get_df_from_worksheet(df_worksheet, "S:S", offset = 1)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/Businesses/' + businesses_name + '/' + businesses_name + 'QueueLevels.json', orient = 'records', indent = 4)
    
    
    df = get_df_from_worksheet(df_worksheet, "P:R", offset = 1)
    df.to_json('JSON/GameData/Businesses/' + businesses_name + '/' + businesses_name + 'SeatsLevels.json', orient = 'records', indent = 4)


# ФУНКЦИЯ ЗАПОЛНЯЕТ .json КАСС
def boothes_json(df_worksheet): 
    df = get_df_from_worksheet(df_worksheet, "C:E", offset = 4)
    df.to_json('JSON/GameData/GatesTicketsBooth/GatesTicketsBooth1Levels.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "C:C", offset = 4)
    df2 = get_df_from_worksheet(df_worksheet, "F:G", offset = 4)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/GatesTicketsBooth/GatesTicketsBooth2Levels.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "C:C", offset = 4)
    df2 = get_df_from_worksheet(df_worksheet, "H:I", offset = 4)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/GatesTicketsBooth/GatesTicketsBooth3Levels.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "C:C", offset = 4)
    df2 = get_df_from_worksheet(df_worksheet, "J:K", offset = 4)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/GatesTicketsBooth/GatesTicketsBooth4Levels.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "C:C", offset = 4)
    df2 = get_df_from_worksheet(df_worksheet, "L:M", offset = 4)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/GatesTicketsBooth/GatesTicketsBooth5Levels.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "C:C", offset = 4)
    df2 = get_df_from_worksheet(df_worksheet, "N:O", offset = 4)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/GatesTicketsBooth/GatesTicketsBooth6Levels.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "C:C", offset = 4)
    df2 = get_df_from_worksheet(df_worksheet, "P:Q", offset = 4)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/GatesTicketsBooth/GatesTicketsBooth7Levels.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "C:C", offset = 4)
    df2 = get_df_from_worksheet(df_worksheet, "R:S", offset = 4)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/GatesTicketsBooth/GatesTicketsBooth8Levels.json', orient = 'records', indent = 4)
    
    
    df = get_df_from_worksheet(df_worksheet, "Y:AE", offset = 1)
    df.to_json('JSON/GameData/GatesTicketsBooth/GatesTicketsBoothIslandsReqs.json', orient = 'records', indent = 4)
    
    
    df = pd.DataFrame([[get_cell_from_worksheet(df_worksheet, "D4"), get_cell_from_worksheet(df_worksheet, "E4")],
                       [get_cell_from_worksheet(df_worksheet, "F4"), get_cell_from_worksheet(df_worksheet, "G4")],
                       [get_cell_from_worksheet(df_worksheet, "H4"), get_cell_from_worksheet(df_worksheet, "I4")],
                       [get_cell_from_worksheet(df_worksheet, "J4"), get_cell_from_worksheet(df_worksheet, "K4")],
                       [get_cell_from_worksheet(df_worksheet, "L4"), get_cell_from_worksheet(df_worksheet, "M4")],
                       [get_cell_from_worksheet(df_worksheet, "N4"), get_cell_from_worksheet(df_worksheet, "O4")],
                       [get_cell_from_worksheet(df_worksheet, "P4"), get_cell_from_worksheet(df_worksheet, "Q4")],
                       [get_cell_from_worksheet(df_worksheet, "R4"), get_cell_from_worksheet(df_worksheet, "S4")]],
                       columns = ['GatesTicketsBoothID', 'GatesTicketsBoothPrice'])
    df.to_json('JSON/GameData/GatesTicketsBooth/GatesTicketsBoothPrices.json', orient = 'records', indent = 4)


# In[32]:


# ФУНКЦИЯ ЗАПОЛНЯЕТ .json СЕРВИСОВ
# service_name IN ("Air","Kithen","Clean")
def service_json(sheet_name, service_name): 
    df = get_df_from_worksheet(df_worksheet, "A:C", offset = 4)
    df.to_json('JSON/GameData/ServiceBusiness/'+service_name+'/'+service_name+'1Levels.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "A:A", offset = 4)
    df2 = get_df_from_worksheet(df_worksheet, "D:E", offset = 4)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/ServiceBusiness/'+service_name+'/'+service_name+'2Levels.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "A:A", offset = 4)
    df2 = get_df_from_worksheet(df_worksheet, "F:G", offset = 4)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/ServiceBusiness/'+service_name+'/'+service_name+'3Levels.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "A:A", offset = 4)
    df2 = get_df_from_worksheet(df_worksheet, "H:I", offset = 4)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/ServiceBusiness/'+service_name+'/'+service_name+'4Levels.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "A:A", offset = 4)
    df2 = get_df_from_worksheet(df_worksheet, "J:K", offset = 4)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/ServiceBusiness/'+service_name+'/'+service_name+'5Levels.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "A:A", offset = 4)
    df2 = get_df_from_worksheet(df_worksheet, "L:M", offset = 4)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/ServiceBusiness/'+service_name+'/'+service_name+'6Levels.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "A:A", offset = 4)
    df2 = get_df_from_worksheet(df_worksheet, "N:O", offset = 4)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/ServiceBusiness/'+service_name+'/'+service_name+'7Levels.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "A:A", offset = 4)
    df2 = get_df_from_worksheet(df_worksheet, "P:Q", offset = 4)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/ServiceBusiness/'+service_name+'/'+service_name+'8Levels.json', orient = 'records', indent = 4)
    
    
    df = get_df_from_worksheet(df_worksheet, "S:Y", offset = 1)
    df.to_json('JSON/GameData/ServiceBusiness/'+service_name+'/'+service_name+'IslandsReqs.json', orient = 'records', indent = 4)
    
    
    df = pd.DataFrame([[get_cell_from_worksheet(df_worksheet, "B4"), get_cell_from_worksheet(df_worksheet, "C4")],
                       [get_cell_from_worksheet(df_worksheet, "D4"), get_cell_from_worksheet(df_worksheet, "I4")],
                       [get_cell_from_worksheet(df_worksheet, "F4"), get_cell_from_worksheet(df_worksheet, "G4")],
                       [get_cell_from_worksheet(df_worksheet, "H4"), get_cell_from_worksheet(df_worksheet, "I4")],
                       [get_cell_from_worksheet(df_worksheet, "J4"), get_cell_from_worksheet(df_worksheet, "K4")],
                       [get_cell_from_worksheet(df_worksheet, "L4"), get_cell_from_worksheet(df_worksheet, "M4")],
                       [get_cell_from_worksheet(df_worksheet, "N4"), get_cell_from_worksheet(df_worksheet, "O4")],
                       [get_cell_from_worksheet(df_worksheet, "P4"), get_cell_from_worksheet(df_worksheet, "Q4")]],
                       columns = ['ServiceBusinessID', 'ServiceBusinessPrice'])
    df.to_json('JSON/GameData/ServiceBusiness/'+service_name+'/'+service_name+'Prices.json', orient = 'records', indent = 4)


# In[33]:


# ФУНКЦИЯ ЗАПОЛНЯЕТ .json СУВЕНИРНОГО МАГАЗИНА
def souvenir_json(sheet_name): 
    df = get_df_from_worksheet(df_worksheet, "D:F", offset = 4)
    df.to_json('JSON/GameData/SouvenirGatesTicketsBooth/SouvenirGatesTicketsBooth1Levels.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "D:D", offset = 4)
    df2 = get_df_from_worksheet(df_worksheet, "G:H", offset = 4)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/SouvenirGatesTicketsBooth/SouvenirGatesTicketsBooth2Levels.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "D:D", offset = 4)
    df2 = get_df_from_worksheet(df_worksheet, "I:J", offset = 4)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/SouvenirGatesTicketsBooth/SouvenirGatesTicketsBooth3Levels.json', orient = 'records', indent = 4)
    
    
    df1 = get_df_from_worksheet(df_worksheet, "D:D", offset = 4)
    df2 = get_df_from_worksheet(df_worksheet, "K:L", offset = 4)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/SouvenirGatesTicketsBooth/SouvenirGatesTicketsBooth4Levels.json', orient = 'records', indent = 4)
    
    
    df = get_df_from_worksheet(df_worksheet, "R:X", offset = 1)
    df.to_json('JSON/GameData/SouvenirGatesTicketsBooth/SouvenirGatesTicketsBoothIslandsReqs.json', orient = 'records', indent = 4)
    
    df = pd.DataFrame([[get_cell_from_worksheet(df_worksheet, "C4"), 
                        get_cell_from_worksheet(df_worksheet, "AG4"),
                        get_cell_from_worksheet(df_worksheet, "AG5"),
                        get_cell_from_worksheet(df_worksheet, "AG6"),
                        get_cell_from_worksheet(df_worksheet, "AG7"),
                        get_cell_from_worksheet(df_worksheet, "AG8"),
                        get_cell_from_worksheet(df_worksheet, "AG9"),
                        get_cell_from_worksheet(df_worksheet, "B4"),]],
                       columns = ['GatesTicketsBoothIslandIncome', 
                                  'GatesTicketsBoothThresholdIsland1',
                                  'GatesTicketsBoothThresholdIsland2',
                                  'GatesTicketsBoothThresholdIsland3',
                                  'GatesTicketsBoothThresholdIsland4',
                                  'GatesTicketsBoothThresholdIsland5',
                                  'GatesTicketsBoothThresholdIsland6',
                                  'GatesTicketsBoothBuyBusinessPrice'])
    df.to_json('JSON/GameData/SouvenirGatesTicketsBooth/SouvenirGatesTicketsBoothParams.json', orient = 'records', indent = 4)
    
    df = pd.DataFrame([[get_cell_from_worksheet(df_worksheet, "E4"), get_cell_from_worksheet(df_worksheet, "F4")],
                       [get_cell_from_worksheet(df_worksheet, "G4"), get_cell_from_worksheet(df_worksheet, "H4")],
                       [get_cell_from_worksheet(df_worksheet, "I4"), get_cell_from_worksheet(df_worksheet, "J4")],
                       [get_cell_from_worksheet(df_worksheet, "K4"), get_cell_from_worksheet(df_worksheet, "L4")]],
                       columns = ['GatesTicketsBoothID', 'GatesTicketsBoothPrice'])
    df.to_json('JSON/GameData/SouvenirGatesTicketsBooth/SouvenirGatesTicketsBoothPrices.json', orient = 'records', indent = 4)


# In[34]:


# ФУНКЦИЯ ЗАПОЛНЯЕТ .json РАСПИСАНИЯ
def schedule_json(sheet_name): 
    df = get_df_from_worksheet(df_worksheet, "A:A", offset = 1)
    df.to_json('JSON/GameData/DayCycle/DayCycleConfig.json', orient = 'records', indent = 4)
    
    df1 = get_df_from_worksheet(df_worksheet, "C:E", offset = 1)
    df2 = get_df_from_worksheet(df_worksheet, "F:F", offset = 1, is_float = False)
    df = pd.concat([df1, df2], axis = 1)
    df.to_json('JSON/GameData/DayCycle/DayCycleSchedule.json', orient = 'records', force_ascii=False, indent = 4)

# ФУНКЦИЯ ЗАПОЛНЯЕТ .json МИРОВ
def worlds_json(sheet_name): 
    df1 = get_df_from_worksheet(df_worksheet, "A:G", offset = 1)
    df2 = get_df_from_worksheet(df_worksheet, "H:H", offset = 1, is_float = False)
    df3 = get_df_from_worksheet(df_worksheet, "I:L", offset = 1)
    df = pd.concat([df1, df2, df3], axis = 1)
    df.to_json('JSON/GameData/Islands/IslandBase.json', orient = 'records', indent = 4)
    
    df = get_df_from_worksheet(df_worksheet, "N:R", offset = 1)
    df.to_json('JSON/GameData/Islands/IslandRatingRewards.json', orient = 'records', force_ascii=False, indent = 4)
    
try:
    excel_file = pd.ExcelFile('Monster Tycoon 1.0.xlsx')
except:
    print("Failing to load file Monster Tycoon 1.0.xlsx")
    time.sleep(10)
    sys.exit()


# In[36]:


try:
    df_worksheet = excel_file.parse(sheet_name = 'Монстр 1')
    businesses_json(df_worksheet, 'Business1')
        
    df_worksheet = excel_file.parse(sheet_name = 'Монстр 3')
    businesses_json(df_worksheet, 'Business3')
    
    df_worksheet = excel_file.parse(sheet_name = 'Монстр 4')
    businesses_json(df_worksheet, 'Business4')
    
    df_worksheet = excel_file.parse(sheet_name = 'Монстр 5')
    businesses_json(df_worksheet, 'Business5')
    
    df_worksheet = excel_file.parse(sheet_name = 'Монстр 6')
    businesses_json(df_worksheet, 'Business6')
    
    df_worksheet = excel_file.parse(sheet_name = 'Монстр 7')
    businesses_json(df_worksheet, 'Business7')
    
    df_worksheet = excel_file.parse(sheet_name = 'Монстр 8')
    businesses_json(df_worksheet, 'Business8')
    
    df_worksheet = excel_file.parse(sheet_name = 'Монстр 9')
    businesses_json(df_worksheet, 'Business9')
except:
    print("Failing to create Monster json")
    time.sleep(10)
    sys.exit()
print("Monster json created")


# In[37]:


try:
    df_worksheet = excel_file.parse(sheet_name = 'Кассы')
    boothes_json(df_worksheet)
except:
    print("Failing to create Booth json")
    time.sleep(10)
    sys.exit()
print("Booth json created")


# In[38]:


try:
    df_worksheet = excel_file.parse(sheet_name = 'Обслуживающий бизнес 1')
    service_json(df_worksheet, 'Clean')
    
    df_worksheet = excel_file.parse(sheet_name = 'Обслуживающий бизнес 2')
    service_json(df_worksheet, 'Kitchen')
except:
    print("Failing to create Service json")
    time.sleep(10)
    sys.exit()
print("Service json created")


# In[39]:


try:
    df_worksheet = excel_file.parse(sheet_name = 'Сувенирный магазин')
    souvenir_json(df_worksheet)
except:
    print("Failing to create Souvenir jsons")
    time.sleep(10)
    sys.exit()
print("Souvenir jsons created")


# In[40]:


try:
    df_worksheet = excel_file.parse(sheet_name = 'Расписание')
    schedule_json('Расписание')
except:
    print("Failing to create Schedule")
    time.sleep(10)
    sys.exit()
print("Schedule jsons created")


try:
    df_worksheet = excel_file.parse(sheet_name = 'Миры')
    worlds_json(df_worksheet)
except:
    print("Failing to create worlds json")
    time.sleep(10)
    sys.exit()
print("Worlds json created")

