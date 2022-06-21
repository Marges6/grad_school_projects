#!/usr/bin/env python
# coding: utf-8

# In[2]:


import requests
import pandas as pd
#import seaborn as sns
#import matplotlib.pyplot as plt
import json
from datetime import datetime
import os
import xlsxwriter
import warnings
warnings.filterwarnings("ignore")
def fxn():
    warnings.warn("deprecated", DeprecationWarning)

with warnings.catch_warnings():
    warnings.simplefilter("ignore")
    fxn()


# #### Functions

# In[3]:


def read_csv_xlsx(path):
    try:
        if '.csv' in path:
            file = pd.read_csv(path, names = ['Waluty'], header = None)
            print(f'\nPlik csv załadowany')
        elif '.xls' in path:
            file = pd.read_excel(path, names = ['Waluty'], header = None)
            print(f'\nPlik xlsx załadowany')
        else:
            print(f'\nNiepoprawna ścieżka - nie znaleziono pliku - '+ path.rsplit('\\',1)[1])
    except ValueError:
        print(f'\nW pliku źródłowym jest więcej niż 1 kolumna!')
    else:
        return file


# In[4]:


def nbp_api_request(requested_currency_abv):
    currency_table = requests.get(f'http://api.nbp.pl/api/exchangerates/rates/A/{requested_currency_abv}/last/30/')
    if currency_table.status_code == 404:
        print('Błąd połączenia - nie znaleziono waluty ' + str(requested_currency_abv))
    elif currency_table.status_code != 200:
        print('Błąd połączenia - error:' + currency_table.status_code)
    else:
        return  currency_table


# In[5]:


def get_path(file):
    for i in os.listdir():
        y = None
        y = i.split('.')[0]
        file = file.split('.')[0]
        if file.lower() == y.lower():
            full_path = curr_dir + i
            return full_path


# In[6]:


# def jprint(obj):
#     # create a formatted string of the Python JSON object
#     text = json.dumps(obj, sort_keys=True, indent=4)
#     print(text)


# ### Workbench:

# In[ ]:


user_response = input('Wpisz jeden kod waluty np. EUR, CHF, DKK lub wpisz CSV dla listy kodów zawartej w pliku .csv/.xls/.xlsx: ')
curr_dir = os.getcwd() + '\\'
if user_response.lower() != 'csv':
    #zamien jeden symbol waluty na DataFrame
    currencies = pd.DataFrame([user_response.upper()],columns=['Waluty'])
else:
    #plik źródłowy z symbolami walut
    
    print(f'\n')
    for i in os.listdir():
        if any(x in i for x in ['csv','xls']):
            print(i)
    full_path = None
    while full_path == None:
        full_path = None
        plik = input(f'\nWybierz plik z powyższej listy skoroszytów znajdujących się w aktualnym folderze. \nPlik powinien zawierać tylko kolumnę z listą wybranych kodów!\n')
        full_path = get_path(plik)
        if full_path == None: print(f'\nNie znaleziono "'+ plik +'" w folderze.')
    currencies = read_csv_xlsx(full_path)
# SET OPTION
choice_type = len(currencies)
#print(choice_type)
# XLSX FILE SET UP
if choice_type == 1:
    wb_name = 'Tabela_'+ str(currencies['Waluty'][0])+ '_' + datetime.today().strftime(('%Y_%m_%d'))
else:
    wb_name = 'Tabele_walut_' + datetime.today().strftime(('%Y_%m_%d'))
path_wb = os.path.join(curr_dir,wb_name + '.xlsx')
#workbook = xlsxwriter.Workbook(path_wb)
writer_engine = pd.ExcelWriter(path_wb, date_format = 'YYYY-MM-DD', datetime_format='YYYY-MM-DD', engine='xlsxwriter')
workbook = writer_engine.book
#try:
for currency in currencies['Waluty']: 
### API REQUEST
    nbp_request = nbp_api_request(currency)
    #print(nbp_request)
    if nbp_request == None: continue
### DATAFRAME
    df = pd.DataFrame(nbp_request.json()['rates'])
    df.rename(columns = {'no':'Tabela_NBP','effectiveDate':'Data_publikacji','mid':'Kurs_średni'}, inplace = True)
    df.Data_publikacji = pd.to_datetime(df.Data_publikacji,format = '%Y-%m-%d')
    df.sort_values('Data_publikacji', ascending = False,inplace = True)
    df.reset_index(drop = True, inplace = True)
### WORKSHEET
    max_row = len(df)
    max_kurs = max(df['Kurs_średni'])
    min_kurs = min(df['Kurs_średni'])
#    max_range_graph = round( max_kurs + 0.01,2)
#    min_range_graph = round( min_kurs - 0.01,2)
    sheet = 'Tabela '+ currency
    
    df.to_excel(writer_engine,sheet_name = sheet, index = False)
    curr_sheet = writer_engine.sheets[sheet]
    curr_sheet.set_column('A:C', 18) #format option
 ### CHART
     #przygotuj obiekt
    chart = workbook.add_chart({'type': 'line'})
     #Seria danych
    chart.add_series({
            'name':'Kurs ' + currency +'/PLN',
            'categories':[sheet, 1, 1,   max_row, 1],
             'values':[sheet, 1, 2, max_row, 2],
             })
    chart.set_x_axis({'name': 'Data'})
     # formatowanie zakresu osi na podstawie wartości kursu
     # mniejsze zakresy min/max wykresu oraz dystansa pomiędzy liniami siatki dla walut o niskiej wartości w stosunku do PLN
    if max_kurs -min_kurs > 0.02:
        max_range_graph = round( max_kurs + 0.01,2)
        min_range_graph = round( min_kurs - 0.01,2)
        chart.set_y_axis({'name': 'Kurs średni '+ currency + '/PLN' ,
                      'min':min_range_graph,'max':max_range_graph,
                      'major_gridlines': {'visible': True}, 'major_unit':0.01})
    elif max_kurs-min_kurs > 0.01:
        max_range_graph = round( max_kurs + 0.002,3)
        min_range_graph = round( min_kurs - 0.002,3)
        chart.set_y_axis({'name': 'Kurs średni '+ currency + '/PLN' ,
                      'min':min_range_graph,'max':max_range_graph,
                      'major_gridlines': {'visible': True}, 'major_unit':0.001})
    else:
        max_range_graph = round( max_kurs + 0.0006,4)
        min_range_graph = round( min_kurs - 0.0006,4)
        chart.set_y_axis({'name': 'Kurs średni '+ currency + '/PLN' ,
                      'min':min_range_graph,'max':max_range_graph,
                      'major_gridlines': {'visible': True}, 'major_unit':0.0005})
    chart.set_legend({'none':True})
    #wstaw obiekt - graf
    curr_sheet.insert_chart('G3', chart, {'x_scale': 2.1, 'y_scale': 2.1} )
    #
    print(currency + ' added')
print('Finished')
writer_engine.save()
workbook.close()
writer_engine.close()
writer_engine.handles = None

#except:
 #print('Wystąpił błąd podczas przetwarzania pliku')


# In[ ]:




