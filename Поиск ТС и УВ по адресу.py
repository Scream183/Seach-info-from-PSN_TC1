#!/usr/bin/env python
# coding: utf-8

# In[225]:


import pandas as pd
from datetime import datetime


# In[226]:


b = pd.read_excel("Z:\\ТС\\Аналитика\\Сотрудники\\Медведев Р.А\\EXPORT_DATA\\TC1_BACKUP\\TC1_18_02_2022.xlsx")


# In[227]:


t = []
replace_list1 = [['Тысяча Девятьсот', '1905'], ['Десятилетия', '10'], ['10-летия', '10'], ['Восьмисотлетия', '800'], ['800-летия', '800'], ['Двадцати Шести', '26'], ['26-ти', '26'], ['Тысяча Восемьсот', '1812'], ['1 Мая', '1'], ['Сорок', '40'], ['Пятьдесят', '50'], ['60-летия', '60'], ['Шестидесятилетия', '60'], ['Восьмого', '8'], ['Девятого', '9']]
for index, row in b.iterrows():
    addr = row['C_STREET']
    for p in replace_list1:
        if not pd.isna(addr):      
            addr = addr.replace(p[0], p[1])
    t.append(addr)
b['w'] = t
street = []
for index, row in b.iterrows():
    addr = str(row['w']) +' '+ str(row['C_ADMINISTRATIVE_DISTRICT']) +' '+ str(row['C_ROOM']) + str(row['C_CITY']) + str(row['C_LOCALITY']) + str(row['C_REGION'])+ str(row['C_HOUSE']) + str(row['C_BUILDING'])
    street.append(addr)
b['Улица'] = street
b['Улица'] = b['Улица'].str.upper() #замена маленьких букв на большие
b['Улица'] = b['Улица'].str.replace('Ё', 'Е') #замена букв в улицах
b['ANNULEMENT'] = b['C_IGNORING_TYPE'].str.len()
b['date_stop'] = pd.to_datetime(b['C_STOP_USING_DATE'],dayfirst = True) #переформатируем столбец дата прекращения из строки в дату
b['date_begin'] = pd.to_datetime(b['C_USE_OBJECT_EMERGENCE_DATE'], format='%d.%m.%Y', errors='coerce') #переформатируем столбец дата возникновения из строки в дату
b = b[~b['date_begin'].isna()]


# In[ ]:


s = pd.read_excel("Z:\\ТС\\Аналитика\\Сотрудники\\Медведев Р.А\\EXPORT_DATA\\PATENT_BACKUP\\PATENT_10_02_2022.xlsx", dtype = str) #
t_psn = []
replace_list1 = [['Тысяча Девятьсот', '1905'], ['Десятилетия', '10'], ['10-летия', '10'], ['Восьмисотлетия', '800'], ['800-летия', '800'], ['Двадцати Шести', '26'], ['26-ти', '26'], ['Тысяча Восемьсот', '1812'], ['1 Мая', '1'], ['Сорок', '40'], ['Пятьдесят', '50'], ['60-летия', '60'], ['Шестидесятилетия', '60'], ['Восьмого', '8'], ['Девятого', '9']]
for index, row in s.iterrows():
    addr = row['STREET']
    for p in replace_list1:
        if not pd.isna(addr):      
            addr = addr.replace(p[0], p[1])
    t_psn.append(addr)
s['Улица'] = t_psn
s['Улица'] = s['Улица'].str.upper() #замена маленьких букв на большие
s['Улица'] = s['Улица'].str.replace('Ё', 'Е') #замена букв в улицах
s['date_finish'] = pd.to_datetime(s['DATE_STOP_PATENT'],dayfirst = True) #переформатируем столбец дата прекращения из строки в дату
s['date_begin'] = pd.to_datetime(s['DATE_START_PATENT'],dayfirst = True)
s['date_loss'] = pd.to_datetime(s['DATE_LOSS_PATENT'], dayfirst = True)
s['date_cessation'] = pd.to_datetime(s['DATE_CESSATION_PATENT'], dayfirst = True)
s['stop_use_day'] = pd.to_datetime(s['DATE_STOP_USE_PATENT'], dayfirst = True)


# In[324]:


def get_yv(df, go, street, house = None, building = None):
    result = df[df['Улица'].str.contains(street, na = False)]
    if house == None:
        pass
    else:
        result = result[result['C_HOUSE'].str.contains(house, na = False)]
    if building == None:
        pass
    else:
        result = result[result['C_BUILDING'].str.contains(building, na = False) | result['C_HOUSE'].str.contains(building, na = False) ]
    result = result[result['C_MARK_NOTICE'] == 1]
    result = result[(result['date_begin'] <= go) | ((result['date_begin'].dt.year == go.year) & (result['date_begin'].dt.quarter == go.quarter))]
    result = result[(result['date_stop'] >= go) | result['date_stop'].isna() | ((result['date_stop'].dt.year == go.year) & (result['date_stop'].dt.quarter == go.quarter))]
    result = result[result['ANNULEMENT'].isna()]
    ven = result[result['C_OBJECT_TYPE'] == 9]
    return result, len(result), len(ven)


# In[325]:


def get_psn(df, go, street, house = None):
    result = df[df['Улица'].str.contains(street, na = False)]
    if house == None:
        pass
    else:
        result = result[result['HOUSE'].str.contains(house, na = False)]
    result = result[(result['date_begin'] <= go)]
    result = result[(result['date_finish'] >= go)]
    result = result[(result['date_loss'] > go) | result['date_loss'].isna()]
    result = result[(result['date_cessation'] > go) | result['date_cessation'].isna()]
    result = result[(result['stop_use_day'] > go) | result['stop_use_day'].isna()]
    return result, len(result)


# # Ищем уведомление

# In[326]:


result_objcts_yv = get_yv(b, pd.to_datetime("today"), 'ОСТАНКИН')[0]
total_yv = get_yv(b, pd.to_datetime("today"), 'ОСТАНКИН')[1]
ven = get_yv(b, pd.to_datetime("today"), 'ОСТАНКИН')[2]


# # Ищем ПСН

# In[327]:


result_objcts_psn = get_psn(s, pd.to_datetime("today"), 'ОСТАНКИН', '53')[0]
total_psn = get_psn(s, pd.to_datetime("today"), 'ОСТАНКИН', '53')[1]


# In[328]:


table = pd.DataFrame({"Всего ТС":[total_yv], "Всего плательщиков ТС": [result_objcts_yv['C_INN'].nunique()],
                      "Всего псн":[total_psn], "Всего плательщиков псн": [result_objcts_psn['INN'].nunique()],
                     "Вендинг": [ven]})


# In[329]:


table


# In[330]:


result_objcts_yv.to_excel("C:\\Users\\yav\\Desktop\\Рапира_ТС_1.xlsx", index = False)


# In[207]:


result_objcts_psn.to_excel("C:\\Users\\yav\\Desktop\\Рапира_PSN.xlsx", index = False)


# In[ ]:




