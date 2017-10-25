# -*- coding: utf-8 -*-
"""
Created on Tue Oct 24 13:21:05 2017

@author: Nehal
"""

import requests
import json
from xlwt import Workbook

r = requests.get('https://bittrex.com/api/v1.1/public/getmarketsummaries') 
parsed = json.dumps(r.json())
data = json.loads(parsed)

#print(json.dumps(parsed, indent=4, sort_keys=True))


import xlwt

from openpyxl import load_workbook


def saveWorkSpace(fields):
    wb = load_workbook('C://Users//Nehal//Desktop//Machine Learning//accounts2.xlsx')
    ws = wb.active
    st = xlwt.easyxf('pattern: pattern solid;')
    st.pattern.pattern_fore_colour = 20
    for i in fields:
        print(i)
        ws.append(i)
        wb.save("accounts2.xlsx")

def walk_dict(data):
    Combined = []
    count = 2
    for k,v in data.items():
        if isinstance(v, dict):
            walk_dict(v)
        else:
            if isinstance(v, list):
                for i in v:
                     sublist = [ i['MarketName'] , i['High'] , i['Low'] , i['Ask'] , i['Last'] , i['OpenSellOrders'] , i['Volume']   
                                 , i['BaseVolume'], i['PrevDay'] ,  ((i['Last'] -  i['PrevDay']) * 100)/  i['PrevDay'] , '=IF(J'+str(count)+'<-10,"DUMPING",IF(AND(J'+str(count)+'>=-10,J'+str(count)+'<0),"DOWFALL",IF(AND(J'+str(count)+'>=0,J'+str(count)+'<=7),"RISING","BOOMING")))'   ]
                     count = count + 1
                     Combined.append(sublist)                                   
                saveWorkSpace(Combined)
                
walk_dict(data)

