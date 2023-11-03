#!/usr/bin/env python
# coding: utf-8

# In[6]:


import pandas as pd
import FinanceDataReader as fdr # 이거설치는 pip install -U finance-datareader 로 해야함  pip install FinanceDataReader하면 안됨
import xlwings as wx     #파일을 생성하는게 아닌 파일이 있는상태에서 엑셀을 다루는 것
from pykrx import stock # 코스피 종목의 주가 관련 정보를 얻는 API 입니다. https://github.com/sharebook-kr/pykrx
import datetime
import numpy as np


# In[7]:


def Vol_Stock(file_name,sheet_name,wb):
    #wb=wx.Book(file_name)
    #variable_sheet=sheet_name
    Company_Name=wb.sheets[sheet_name].range('F56').expand('right').value 
    Start_Date=wb.sheets[sheet_name].range('F57').expand('right').value 
    End_Date=wb.sheets[sheet_name].range('F58').expand('right').value 
    #financeDataReader에서 주가 data 받음
    code_data=fdr.StockListing('KRX')
    Company_symbol=[]
    Company_stock_Price_1=[]
    for i,name in enumerate(Company_Name):
        Company_symbol.append(code_data[code_data['Name']==name]['Code'])  ##회사 종목코드 가져오기 코드
        Company_stock_Price_1.append(fdr.DataReader(Company_symbol[i].values[0], Start_Date[i], End_Date[i]))
    #financeDataReader에서 받지못하는 정보도 있으니 krxㅇ서 한번더 받음

    Company_stock_Price_2=[]

    for i,symbol in enumerate(Company_symbol):
        inf=stock.get_market_cap(Start_Date[i].strftime('%Y%m%d'),End_Date[i].strftime('%Y%m%d'),symbol.tolist()[0]) #자료형을 맞춰야함
        Company_stock_Price_2.append(inf)
    
    ################### 내려받은 자료에 필드 추가하여 원하는 자료 형태 Table만들기#####
    original_std=[]
    adj_std=[]
    for i, letter in enumerate(Company_Name):
        Tables_1=pd.DataFrame(Company_stock_Price_1[i])
        Tables_2=pd.DataFrame(Company_stock_Price_2[i])
        Table=pd.concat([Tables_1,Tables_2],axis=1)
        Table['수정주가']           =np.array(Table['Close'])*np.array(Table['상장주식수'])/np.array(Table['상장주식수'][0])
        Table['원주가변동성']=0
        Table['원주가변동성'][:-1]  =np.log(np.array(Table['Close'][:-1])/np.array(Table['Close'][1:]))
        Table['수정주가변동성']=0
        Table['수정주가변동성'][:-1]=np.log(np.array(Table['수정주가'][:-1])/np.array(Table['수정주가'][1:]))
        
        Add_sheet_name=letter
        if Add_sheet_name in wb.sheet_names:
            wb.sheets[Add_sheet_name].delete()
        else:
            pass
        
        Vol_sheet=wb.sheets.add(name=Add_sheet_name,after=wb.sheet_names[-1])
        Vol_sheet.range('C5').value=Table
        
        original_std.append(np.std(np.array(wb.sheets[letter].range('O7').expand('down').value)[:-1]))
        adj_std.append(np.std(np.array(wb.sheets[letter].range('P7').expand('down').value)[:-1]))
    
    wb.sheets[sheet_name].range('Q56').value=np.array(original_std).mean()
    wb.sheets[sheet_name].range('R56').value=np.array(adj_std).mean()


# In[ ]:




