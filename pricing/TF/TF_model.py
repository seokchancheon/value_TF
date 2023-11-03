# -*- coding: utf-8 -*-
"""
Created on Wed Nov  1 23:57:44 2023

@author: cjstj
"""

#!/usr/bin/env python
# coding: utf-8

# # 0.1 부스트래핑

# In[1]:


import xlwings as wx
import numpy as np


# file_name='TF(BDT).xlsx'

#wb=wx.Book(file_name) #워크북 열기      >>view.py에 기재
#sheet_name='TF모형(BDT)_python'        >>view.py에 기재

# In[25]:


def dx(S_2,j,ADJ_Kd_F,dt):####4기간부터 구하는 함수...
    S_3=S_2[:-1] * 0.5 * 1 / ((1 + ADJ_Kd_F[:j - 2, j - 2] / 100) ** dt)+\
        S_2[1:]  * 0.5 * 1 / ((1 + ADJ_Kd_F[:j - 2, j - 2] / 100) ** dt)
    j=j-1
    if j>=3:
        return(dx(S_3,j,ADJ_Kd_F,dt))
    else:
        return(S_3)


# In[26]:


def BDT_interest_adj(x,j,interest_binomial, ADJ_Kd_F,B_PV_by_Kd_S,dt): #할인율 트리를 기준으로 생각해야함 !! 마지막에만 미지수있음
    #global j
    '''S_1이 미지수가 있어서 함수를 못태우니, S_2까지 만들고 함수를 태우는것'''
    if j==2:
        S_1=0.5* 10_000 / ((1 + (interest_binomial[:j, j]+x) / 100) ** dt)+\
            0.5 * 10_000 / ((1 + (interest_binomial[:j,j]+x) / 100) ** dt)#
        S_2=S_1[:-1]*0.5 * 1 / ((1 + ADJ_Kd_F[:j - 1, j - 1] / 100) ** dt)+\
            S_1[1:]*0.5 * 1 / ((1 + ADJ_Kd_F[:j - 1, j - 1] / 100) ** dt)
        return(S_2-B_PV_by_Kd_S[j])
    else:
        S_1=0.5* 10_000 / ((1 + (interest_binomial[:j, j]+x) / 100) ** dt)+\
            0.5 * 10_000 / ((1 + (interest_binomial[:j,j]+x) / 100) ** dt)#
        S_2=S_1[:-1]*0.5 * 1 / ((1 + ADJ_Kd_F[:j - 1, j - 1] / 100) ** dt)+\
            S_1[1:]*0.5 * 1 / ((1 + ADJ_Kd_F[:j - 1, j - 1] / 100) ** dt)        
        return(dx(S_2,j,ADJ_Kd_F,dt)-B_PV_by_Kd_S[j])
    








def TF_model(wb,sheet_name,file_name):
    # In[2]:
    import numpy as np

    np.set_printoptions(precision=3)

    #### 변수받기
    
    n=wb.sheets[sheet_name].range('F6').value 
    freq=wb.sheets[sheet_name].range('L19').value
    
    n=int(n)     # 'float' object cannot be interpreted as an integer 이런오류가 남
    freq=int(freq)
    
    Matuality=wb.sheets[sheet_name].range('C30').expand('right').value
    YTM_Rf=wb.sheets[sheet_name].range('C31').expand('right').value
    YTM_Kd=wb.sheets[sheet_name].range('C32').expand('right').value
    
    
    # In[3]:
    
    import pricing.TF.Bootstraping_basic as Bootstraping_basic
 
    
    
    # In[4]:
    
    
    # 부스트래핑 계산
    Rf_F, Rf_Y, Rf_S=Bootstraping_basic.FYS_Rate(Matuality,YTM_Rf,n,freq)
    Kd_F, Kd_Y, Kd_S=Bootstraping_basic.FYS_Rate(Matuality,YTM_Kd,n,freq)
    
    
    # In[5]:
    
    
    # 부스트래핑 출력
    wb.sheets[sheet_name].range('C41').expand('table').clear()
    wb.sheets[sheet_name].range('C41').value=np.arange(0,n+1)
    wb.sheets[sheet_name].range('C42').value=Rf_Y
    wb.sheets[sheet_name].range('C43').value=Rf_S
    wb.sheets[sheet_name].range('C44').value=Rf_F
    wb.sheets[sheet_name].range('C45').value=Kd_Y
    wb.sheets[sheet_name].range('C46').value=Kd_S
    wb.sheets[sheet_name].range('C47').value=Kd_F
    
    
    # In[ ]:
    
    
    
    
    
    # # 0.2 변동성 산출
    
    # In[6]:
    
    import pricing.TF.Vol_Stock as Vol_Stock

    Vol_Stock.Vol_Stock(file_name,sheet_name,wb)
    
    
    # 
    
    # 
    
    # # 1. 주가 이항트리작성
    
    # In[7]:
    
    
    S0,sigma,dt,S_U,S_D=wb.sheets[sheet_name].range('C55').expand('down').value
    
    
    # In[8]:
    
    
    import numpy as np
    np.set_printoptions(precision=3)
    
    
    # In[9]:
    
    
    S_Tree=np.zeros((n+1,n+1))
    S_Tree[0,0]=S0
    for i in np.arange(0,n+1):
        S_Tree[i,:]=S0*(S_U**np.arange(0,n+1))*(S_D**(2*i))
    S_Tree=np.triu(S_Tree)
    
    
    # **S_Tree 시트생성**
    
    # In[10]:
    
    
    Add_sheet_name='1. S_Tree'
    
    if Add_sheet_name in wb.sheet_names:
        wb.sheets[Add_sheet_name].delete()
    else:
        pass
    
    S_Tree_Table_sheet=wb.sheets.add(name=Add_sheet_name,after=wb.sheet_names[-1])
    
    S_Tree_Table_sheet.range("C2").value=Add_sheet_name
    S_Tree_Table_sheet.range("C4").value="Node"
    S_Tree_Table_sheet.range("D4").value=np.arange(0,n+1)
    S_Tree_Table_sheet.range("D5").value=S_Tree
    
    
    # # 2. 리픽싱
    
    # In[11]:
    
    
    Issue_price,convert_price,refix_price=wb.sheets[sheet_name].range('C65').expand('down').value
    original_r, refix_r=wb.sheets[sheet_name].range('D66').expand('down').value
    
    
    # In[12]:
    
    
    R_Tree=np.zeros((n+1,n+1))
    refix_Tree=(S_Tree>=refix_price)*original_r+(S_Tree<refix_price)*refix_r
    refix_Tree=np.triu(refix_Tree)
    
    
    # **리픽싱 시트생성**
    
    # In[13]:
    
    
    Add_sheet_name='2. Refix_Tree'
    
    
    # In[14]:
    
    
    Add_sheet_name='2. Refix_Tree'
    
    if Add_sheet_name in wb.sheet_names:
        wb.sheets[Add_sheet_name].delete()
    else:
        pass
    
    S_Tree_Table_sheet=wb.sheets.add(name=Add_sheet_name,after=wb.sheet_names[-1])
    
    S_Tree_Table_sheet.range("C2").value=Add_sheet_name
    S_Tree_Table_sheet.range("C4").value="Node"
    S_Tree_Table_sheet.range("D4").value=np.arange(0,n+1)
    S_Tree_Table_sheet.range("D5").value=refix_Tree
    
    
    # # 3. 희석주가 Tree
    
    # In[15]:
    
    
    Stock_n,Rcps_n,Bond_face=wb.sheets[sheet_name].range('C72').expand('down').value
    
    
    # In[16]:
    
    
    Diluted_S_Tree=S_Tree*(Stock_n+Rcps_n)/(Stock_n+Rcps_n*refix_Tree)
    
    
    # **희석주가 시트생성**
    
    # In[17]:
    
    
    Add_sheet_name='3. Diluted_S_Tree'
    
    if Add_sheet_name in wb.sheet_names:
        wb.sheets[Add_sheet_name].delete()
    else:
        pass
    
    S_Tree_Table_sheet=wb.sheets.add(name=Add_sheet_name,after=wb.sheet_names[-1])
    
    S_Tree_Table_sheet.range("C2").value=Add_sheet_name
    S_Tree_Table_sheet.range("C4").value="Node"
    S_Tree_Table_sheet.range("D4").value=np.arange(0,n+1)
    S_Tree_Table_sheet.range("D5").value=Diluted_S_Tree
    
    
    # # 4. BDT
    
    # In[18]:
    
    
    # BDT변수 받아오기
    
    int_info_sheet="기간별Kd_YTM"
    Matuality=wb.sheets[int_info_sheet].range('B3').expand('right').value
    Eachday_Kd=wb.sheets[int_info_sheet].range('B4').expand('table').value #일수갯수 행, 1열로 담김
    
    
    # **일자별 부스트래핑**
    
    # In[19]:
    
    from pricing.TF.Bootstraping_basic import FYS_Rate

    Forward_Rate_Table  =[]
    YTM_Table           =[]
    Spot_rate_Table     =[]
    
    
    # In[20]:
    
    
    for kd in Eachday_Kd:
        Forward_Rate,YTM,Spot_rate =FYS_Rate(Matuality,kd, n,freq)
        Forward_Rate_Table.append(Forward_Rate)
        YTM_Table.append(YTM)
        Spot_rate_Table.append(Spot_rate)#리스트 상태
    
    #array로 변환    
    Forward_Rate_Table=np.array(Forward_Rate_Table)
    YTM_Table=np.array(YTM_Table)
    Spot_rate_Table=np.array(Spot_rate_Table)
    
    
    # In[21]:
    
    
    Forward_Vol=Forward_Rate_Table.std(axis=0)/100 # 각 열별 변동성
    
    
    # In[22]:
    
    
    BDT_U=np.exp(np.sqrt(dt)*Forward_Vol)      # U, D산출
    BDT_D=1/BDT_U
    
    
    # **이자율이항트리**
    
    # In[23]:
    
    
    #1행 setting
    interest_binomial=np.zeros((n + 1,n + 1))
    interest_binomial[0,1]=Kd_F[1]
    interest_binomial[0,2:]=interest_binomial[0,1]*np.cumprod(BDT_U[1:n])
    #2행부터 반복문
    for i in np.arange(1,n+1):
        interest_binomial[i,i+1:]=interest_binomial[i-1,i:n]*BDT_D[i:n]
    
    
    # In[24]:
    
    
    np.cumprod(BDT_U[1:n])
    
    

    
    # In[27]:
    
    
    '''위함수를 밖으로 빼서하니 에러남... 밖으로 빼는 방법좀...'''
    from scipy.optimize import fsolve
    
    ADJ_Kd_F=np.zeros_like(interest_binomial)
    ADJ_Kd_F[0,1]=interest_binomial[0,1]
    B_PV_by_Kd_S=10_000/(1+Kd_S/100)**(dt*np.arange(0,n+1))  ##SPOT RATE로 할인한 채권의 현재가치
    
    for j in range(2, n + 1):
        a=fsolve(BDT_interest_adj, 0.5,args=(j,interest_binomial,ADJ_Kd_F,B_PV_by_Kd_S,dt))
        ADJ_Kd_F[:j, j] = interest_binomial[:j, j]+a
        #print(a)
    
    
    # In[28]:
    
    
    B_list=[1]
    for j in np.arange(1,n+1):
        B=np.ones(j+1)
        if j==1:
            B=1 / ((1 + ADJ_Kd_F[0, 1 ] / 100) ** dt)
        if j>=2:
            while j>=2:
                B=B[:-1]*0.5 * 1 / ((1 + ADJ_Kd_F[:j, j ] / 100) ** dt)+\
                  B[1:] *0.5 * 1 / ((1 + ADJ_Kd_F[:j , j ] / 100) ** dt)
                j=j-1
                if j==1:
                    B=B.sum()*0.5 / ((1 + ADJ_Kd_F[0, 1 ] / 100) ** dt)
        B_list.append(B)
    
    #wb.sheets[sheet_name].range('F92').value=np.array(B_list)*10_000
    
    
    # In[29]:
    
    
    #wb.sheets[sheet_name].range('F93').value=B_PV_by_Kd_S
    
    
    # **이자율이항트리 시트생성**
    
    # In[30]:
    
    
    Add_sheet_name='4. interest_binomial'
    
    if Add_sheet_name in wb.sheet_names:
        wb.sheets[Add_sheet_name].delete()
    else:
        pass
    
    S_Tree_Table_sheet=wb.sheets.add(name=Add_sheet_name,after=wb.sheet_names[-1])
    
    S_Tree_Table_sheet.range("C2").value=Add_sheet_name
    S_Tree_Table_sheet.range("C4").value="Node"
    S_Tree_Table_sheet.range("D4").value=np.arange(0,n+1)
    S_Tree_Table_sheet.range("D5").value=interest_binomial
    
    
    # **조정이자율 이자율이항트리 시트생성**
    
    # In[31]:
    
    
    Add_sheet_name='5. ADJ_interest_binomial'
    
    if Add_sheet_name in wb.sheet_names:
        wb.sheets[Add_sheet_name].delete()
    else:
        pass
    
    S_Tree_Table_sheet=wb.sheets.add(name=Add_sheet_name,after=wb.sheet_names[-1])
    
    S_Tree_Table_sheet.range("C2").value=Add_sheet_name
    S_Tree_Table_sheet.range("C4").value="Node"
    S_Tree_Table_sheet.range("D4").value=np.arange(0,n+1)
    S_Tree_Table_sheet.range("D5").value=ADJ_Kd_F
    
    
    # # 4. 채권상환가치
    
    # In[32]:
    
    
    Gf,Redeem_Start=wb.sheets[sheet_name].range('C84').expand('down').value
    
    Redeem_array=(np.arange(0,n+1)>=Redeem_Start)*1                 
    Redeem_Amount=Bond_face*(1+Gf)**(dt*np.arange(0,n+1))
    Bond_V=Redeem_Amount*Redeem_array  #  채권가치 Array를 만들고
    B_int_V=Bond_V*Redeem_array*np.triu(np.ones((n+1,n+1)))
    
    
    # In[33]:
    
    
    for i in np.arange(n,0,-1):
        B=B_int_V[ :i,  i] *0.5 * 1 / ((1 + ADJ_Kd_F[:i, i ] / 100) ** dt)+\
          B_int_V[1:i+1,i] *0.5 * 1 / ((1 + ADJ_Kd_F[:i, i ] / 100) ** dt)
        B_int_V[:i,i-1]=np.maximum(B_int_V[:i,i-1],B) # np.max는 배열내 최댓값을 구하는 것
    
    
    # **채권상환가치 시트생성**
    
    # In[34]:
    
    
    Add_sheet_name='6. 채권상환가치'
    
    if Add_sheet_name in wb.sheet_names:
        wb.sheets[Add_sheet_name].delete()
    else:
        pass
    
    S_Tree_Table_sheet=wb.sheets.add(name=Add_sheet_name,after=wb.sheet_names[-1])
    
    S_Tree_Table_sheet.range("C2").value=Add_sheet_name
    S_Tree_Table_sheet.range("C4").value="Node"
    S_Tree_Table_sheet.range("D4").value=np.arange(0,n+1)
    S_Tree_Table_sheet.range("D5").value=B_int_V
    
    
    # In[ ]:
    
    
    
    
    
    # # 5. 주식전환가치
    
    # In[35]:
    
    
    Conversion_Start=wb.sheets[sheet_name].range('C93').value
    
    
    Conversion_array=(np.arange(0,n+1)>=Conversion_Start)*1
    S_int_V=Conversion_array*Diluted_S_Tree
    
    
    # **주식전환가치 시트생성**
    
    # In[36]:
    
    
    Add_sheet_name='7. 주식전환가치'
    
    if Add_sheet_name in wb.sheet_names:
        wb.sheets[Add_sheet_name].delete()
    else:
        pass
    
    S_Tree_Table_sheet=wb.sheets.add(name=Add_sheet_name,after=wb.sheet_names[-1])
    
    S_Tree_Table_sheet.range("C2").value=Add_sheet_name
    S_Tree_Table_sheet.range("C4").value="Node"
    S_Tree_Table_sheet.range("D4").value=np.arange(0,n+1)
    S_Tree_Table_sheet.range("D5").value=S_int_V
    
    
    # In[ ]:
    
    
    
    
    
    # # 6. 전환사채 내재가치(Intrinsic Value)
    
    # In[37]:
    
    
    RCPS_int_V=(S_int_V > B_int_V)*S_int_V+(S_int_V<B_int_V)*B_int_V
    
    
    # **전환사채내재가치 시트생성**
    
    # In[38]:
    
    
    Add_sheet_name='8. 전환사채 내재가치'
    
    if Add_sheet_name in wb.sheet_names:
        wb.sheets[Add_sheet_name].delete()
    else:
        pass
    
    S_Tree_Table_sheet=wb.sheets.add(name=Add_sheet_name,after=wb.sheet_names[-1])
    
    S_Tree_Table_sheet.range("C2").value=Add_sheet_name
    S_Tree_Table_sheet.range("C4").value="Node"
    S_Tree_Table_sheet.range("D4").value=np.arange(0,n+1)
    S_Tree_Table_sheet.range("D5").value=RCPS_int_V
    
    
    # # 전환사채 가치 계산
    
    # In[39]:
    
    
    p=(np.exp(Rf_F/100*dt)-S_D)/(S_U-S_D)
    q=1-p
    
    
    # **수의상환 Array 생성**
    
    # In[40]:
    
    
    Period_of_call, Reverse_C_Price=wb.sheets[sheet_name].range('C107').expand('down').value
    
    Reverse_Call_array=(np.arange(0,n+1)>=Period_of_call)*1*Reverse_C_Price
    
    
    # **각 테이블 생성**
    
    # In[41]:
    
    
    RCPS_V=np.zeros_like(S_Tree)
    D_Table     =np.zeros_like(S_Tree)
    RCPS_H_V=np.zeros_like(S_Tree)
    S_H_V   =np.zeros_like(S_Tree)
    S_H_V_1 =np.zeros_like(S_Tree)
    B_H_V   =np.zeros_like(S_Tree)
    B_H_V_1 =np.zeros_like(S_Tree)
    
    
    # **n 번째 마지막노드 값 채우기**
    
    # In[42]:
    
    
    RCPS_V[:,-1]     = RCPS_int_V[:,-1]
    D_Table[:,-1]    = (RCPS_V[:,-1]==S_int_V[:,-1])*10+(RCPS_V[:,-1]==B_int_V[:,-1])*20+((RCPS_V[:,-1]!=S_int_V[:,-1])*(RCPS_V[:,-1]!=B_int_V[:,-1]))*30    # 10 : 전환 , 20 : 상환 , 30 : 보유
    
    S_H_V  [:,-1]    = (D_Table[:,-1]==10)*S_int_V[:,-1]
    S_H_V_1[:,-1]    = (D_Table[:,-1]==10)*S_int_V[:,-1]
    B_H_V  [:,-1]    = (D_Table[:,-1]==20)*B_int_V[:,-1]
    B_H_V_1[:,-1]    = (D_Table[:,-1]==20)*B_int_V[:,-1]
    
    RCPS_H_V[:,-1] = S_H_V[:,-1]+B_H_V[:,-1]
    
    
    # In[43]:
    
    
    # 마지막노드 전값부터 처음값까지 채우기
    
    Rf_discount=np.exp(-Rf_F/100*dt) #한구간 할인이니 dt만큼만 할인하면 됨
    Kd_discount=np.exp(-ADJ_Kd_F/100*dt)*(ADJ_Kd_F>0)
    Reverse_Call_table=np.triu(np.tile(Reverse_Call_array,(n+1,1)))
    
    for i in np.arange(n,0,-1):
        S_H_V[:i,i-1]    = (S_H_V_1[:i,i]*p[i]+S_H_V_1[1:i+1,i]*q[i])*Rf_discount[i]
        B_H_V[:i,i-1]    = (B_H_V_1[:i,i]*0.5+B_H_V_1[1:i+1,i]*0.5)*Kd_discount[:i,i]
    
        RCPS_H_V[:i,i-1] = S_H_V[:i,i-1]+B_H_V[:i,i-1]
        
        if i>=int(Period_of_call):
            RCPS_V[:i,i-1]     = np.maximum(RCPS_int_V[:i,i-1],np.minimum(RCPS_H_V[:i,i-1],Reverse_Call_table[:i,i-1]))
        else:
            RCPS_V[:i,i-1]     = np.maximum(RCPS_int_V[:i,i-1],RCPS_H_V[:i,i-1])
        
        D_Table[:i,i-1]  = (RCPS_V[:i,i-1]==S_int_V[:i,i-1])*10+\
                           (RCPS_V[:i,i-1]==B_int_V[:i,i-1])*20+\
                           (RCPS_V[:i,i-1]==Reverse_Call_array[i-1])*20+\
                           ((RCPS_V[:i,i-1]!=S_int_V[:i,i-1])*(RCPS_V[:i,i-1]!=B_int_V[:i,i-1])*(RCPS_V[:i,i-1]!=Reverse_Call_array[i-1]))*30    # 10 : 전환 , 20 : 상환 , 30 : 보유
    
        S_H_V_1[:i,i-1]  = (D_Table[:i,i-1]==10)*S_int_V[:i,i-1]+((D_Table[:i,i-1]!=10)*(D_Table[:i,i-1]!=20))*(S_H_V_1[:i,i]*p[i]+S_H_V_1[1:i+1,i]*q[i])*Rf_discount[i]
        B_H_V_1[:i,i-1]  = (D_Table[:i,i-1]==20)*B_int_V[:i,i-1]+((D_Table[:i,i-1]!=10)*(D_Table[:i,i-1]!=20))*(B_H_V_1[:i,i]*p[i]+B_H_V_1[1:i+1,i]*q[i])*Kd_discount[:i,i]
    
    
    
        
    wb.sheets[sheet_name].range('M130').value =RCPS_V[0,0]    
    
    
    # #### 각 Table sheet 출력
    
    # In[44]:
    
    
    Table_list_name=wb.sheets[sheet_name].range('B114').expand('down').value
    
    Table_list=[RCPS_V,
                D_Table,
                RCPS_H_V,
                S_H_V,
                S_H_V_1,
                B_H_V,
                B_H_V_1]
    
    Table_result=dict(zip(Table_list_name, Table_list))  # 시트이름과 내용 묶기
    
    
    # In[45]:
    
    
    for key, value in Table_result.items():#key가 시트이름, value가 내용
       
       Add_sheet_name=key
       
       if Add_sheet_name in wb.sheet_names:
           wb.sheets[Add_sheet_name].delete()
       else:
           pass
    
       Result_Table_sheet=wb.sheets.add(name=Add_sheet_name,after=wb.sheet_names[-1])
    
       Result_Table_sheet.range("C2").value=Add_sheet_name
       Result_Table_sheet.range("C4").value="Node"
       Result_Table_sheet.range("D4").value=np.arange(0,n+1)
       Result_Table_sheet.range("D5").value=value
    
    
    # In[46]:
    
    
    wb.sheets[sheet_name].range('B130').value=RCPS_V[0,0]
    
    
    # In[ ]:
    
    
    
    
    
    # # 수의상환권이 없는 RCPS가치 구하기
    # <span style='background-color:#fff5b1'>**수의상환 Array 생성 >>이것만 변경**  </span>
    
    # # 4. 채권상환가치
    
    # In[47]:
    
    
    Gf,Redeem_Start=wb.sheets[sheet_name].range('C84').expand('down').value
    
    Redeem_array=(np.arange(0,n+1)>=Redeem_Start)*1                 
    Redeem_Amount=Bond_face*(1+Gf)**(dt*np.arange(0,n+1))
    Bond_V=Redeem_Amount*Redeem_array  #  채권가치 Array를 만들고
    B_int_V=Bond_V*Redeem_array*np.triu(np.ones((n+1,n+1)))
    
    
    # In[48]:
    
    
    Gf,Redeem_Start
    
    
    # In[49]:
    
    
    for i in np.arange(n,0,-1):
        B=B_int_V[ :i,  i] *0.5 * 1 / ((1 + ADJ_Kd_F[:i, i ] / 100) ** dt)+\
          B_int_V[1:i+1,i] *0.5 * 1 / ((1 + ADJ_Kd_F[:i, i ] / 100) ** dt)
        B_int_V[:i,i-1]=np.maximum(B_int_V[:i,i-1],B) # np.max는 배열내 최댓값을 구하는 것
    
    
    # # 5. 주식전환가치
    
    # In[50]:
    
    
    Conversion_Start=wb.sheets[sheet_name].range('C93').value
    
    
    Conversion_array=(np.arange(0,n+1)>=Conversion_Start)*1
    S_int_V=Conversion_array*Diluted_S_Tree
    
    
    # # 6. 전환사채 내재가치(Intrinsic Value)
    
    # In[51]:
    
    
    RCPS_int_V=(S_int_V > B_int_V)*S_int_V+(S_int_V<B_int_V)*B_int_V
    
    
    # # 전환사채 가치 계산
    
    # <span style='background-color:#fff5b1'>**수의상환 Array 생성 >> 변경**  </span>
    
    # In[52]:
    
    
    Period_of_call, Reverse_C_Price=wb.sheets[sheet_name].range('C107').expand('down').value
    Period_of_call=n+1  #  변경
    Reverse_Call_array=(np.arange(0,n+1)>=Period_of_call)*1*Reverse_C_Price
    
    
    # **각 테이블 생성**
    
    # In[53]:
    
    
    RCPS_V=np.zeros_like(S_Tree)
    D_Table     =np.zeros_like(S_Tree)
    RCPS_H_V=np.zeros_like(S_Tree)
    S_H_V   =np.zeros_like(S_Tree)
    S_H_V_1 =np.zeros_like(S_Tree)
    B_H_V   =np.zeros_like(S_Tree)
    B_H_V_1 =np.zeros_like(S_Tree)
    
    
    # **n 번째 마지막노드 값 채우기**
    
    # In[54]:
    
    
    RCPS_V[:,-1]     = RCPS_int_V[:,-1]
    D_Table[:,-1]    = (RCPS_V[:,-1]==S_int_V[:,-1])*10+(RCPS_V[:,-1]==B_int_V[:,-1])*20+((RCPS_V[:,-1]!=S_int_V[:,-1])*(RCPS_V[:,-1]!=B_int_V[:,-1]))*30    # 10 : 전환 , 20 : 상환 , 30 : 보유
    
    S_H_V  [:,-1]    = (D_Table[:,-1]==10)*S_int_V[:,-1]
    S_H_V_1[:,-1]    = (D_Table[:,-1]==10)*S_int_V[:,-1]
    B_H_V  [:,-1]    = (D_Table[:,-1]==20)*B_int_V[:,-1]
    B_H_V_1[:,-1]    = (D_Table[:,-1]==20)*B_int_V[:,-1]
    
    RCPS_H_V[:,-1] = S_H_V[:,-1]+B_H_V[:,-1]
    
    
    # In[55]:
    
    
    # 마지막노드 전값부터 처음값까지 채우기
    
    Rf_discount=np.exp(-Rf_F/100*dt) #한구간 할인이니 dt만큼만 할인하면 됨
    Kd_discount=np.exp(-ADJ_Kd_F/100*dt)*(ADJ_Kd_F>0)
    Reverse_Call_table=np.triu(np.tile(Reverse_Call_array,(n+1,1)))
    
    for i in np.arange(n,0,-1):
        S_H_V[:i,i-1]    = (S_H_V_1[:i,i]*p[i]+S_H_V_1[1:i+1,i]*q[i])*Rf_discount[i]
        B_H_V[:i,i-1]    = (B_H_V_1[:i,i]*0.5+B_H_V_1[1:i+1,i]*0.5)*Kd_discount[:i,i]
    
        RCPS_H_V[:i,i-1] = S_H_V[:i,i-1]+B_H_V[:i,i-1]
        
        if i>=int(Period_of_call):
            RCPS_V[:i,i-1]     = np.maximum(RCPS_int_V[:i,i-1],np.minimum(RCPS_H_V[:i,i-1],Reverse_Call_table[:i,i-1]))
        else:
            RCPS_V[:i,i-1]     = np.maximum(RCPS_int_V[:i,i-1],RCPS_H_V[:i,i-1])
        
        D_Table[:i,i-1]  = (RCPS_V[:i,i-1]==S_int_V[:i,i-1])*10+\
                           (RCPS_V[:i,i-1]==B_int_V[:i,i-1])*20+\
                           (RCPS_V[:i,i-1]==Reverse_Call_array[i-1])*20+\
                           ((RCPS_V[:i,i-1]!=S_int_V[:i,i-1])*(RCPS_V[:i,i-1]!=B_int_V[:i,i-1])*(RCPS_V[:i,i-1]!=Reverse_Call_array[i-1]))*30    # 10 : 전환 , 20 : 상환 , 30 : 보유
    
        S_H_V_1[:i,i-1]  = (D_Table[:i,i-1]==10)*S_int_V[:i,i-1]+((D_Table[:i,i-1]!=10)*(D_Table[:i,i-1]!=20))*(S_H_V_1[:i,i]*p[i]+S_H_V_1[1:i+1,i]*q[i])*Rf_discount[i]
        B_H_V_1[:i,i-1]  = (D_Table[:i,i-1]==20)*B_int_V[:i,i-1]+((D_Table[:i,i-1]!=10)*(D_Table[:i,i-1]!=20))*(B_H_V_1[:i,i]*p[i]+B_H_V_1[1:i+1,i]*q[i])*Kd_discount[:i,i]
    
    
    
        
    wb.sheets[sheet_name].range('C130').value =RCPS_V[0,0]    
    
    
    # In[ ]:
    
    
    
    
    
    # # 수의상환권+상환권이 없는 RCPS가치 구하기
    # <span style='background-color:#fff5b1'>**수의상환 Array 변경상태에서 채권상환 가치부분 변경**  </span>
    
    # # 4. 채권상환가치  >>정보변경
    
    # In[56]:
    
    
    Gf,Redeem_Start=wb.sheets[sheet_name].range('C84').expand('down').value
    
    Redeem_Start=n  ### >>변경
    
    Redeem_array=(np.arange(0,n+1)>=Redeem_Start)*1                 
    Redeem_Amount=Bond_face*(1+Gf)**(dt*np.arange(0,n+1))
    Bond_V=Redeem_Amount*Redeem_array  #  채권가치 Array를 만들고
    B_int_V=Bond_V*Redeem_array*np.triu(np.ones((n+1,n+1)))
    
    
    # In[57]:
    
    
    for i in np.arange(n,0,-1):
        B=B_int_V[ :i,  i] *0.5 * 1 / ((1 + ADJ_Kd_F[:i, i ] / 100) ** dt)+\
          B_int_V[1:i+1,i] *0.5 * 1 / ((1 + ADJ_Kd_F[:i, i ] / 100) ** dt)
        B_int_V[:i,i-1]=np.maximum(B_int_V[:i,i-1],B) # np.max는 배열내 최댓값을 구하는 것
    
    
    # # 5. 주식전환가치
    
    # In[58]:
    
    
    Conversion_Start=wb.sheets[sheet_name].range('C93').value
    
    
    Conversion_array=(np.arange(0,n+1)>=Conversion_Start)*1
    S_int_V=Conversion_array*Diluted_S_Tree
    
    
    # # 6. 전환사채 내재가치(Intrinsic Value)
    
    # In[59]:
    
    
    RCPS_int_V=(S_int_V > B_int_V)*S_int_V+(S_int_V<B_int_V)*B_int_V
    
    
    # # 전환사채 가치 계산
    
    # **수의상환 Array 생성**
    
    # In[60]:
    
    
    Period_of_call, Reverse_C_Price=wb.sheets[sheet_name].range('C107').expand('down').value
    
    Period_of_call=n+1
    
    Reverse_Call_array=(np.arange(0,n+1)>=Period_of_call)*1*Reverse_C_Price
    
    
    # **각 테이블 생성**
    
    # In[61]:
    
    
    RCPS_V=np.zeros_like(S_Tree)
    D_Table     =np.zeros_like(S_Tree)
    RCPS_H_V=np.zeros_like(S_Tree)
    S_H_V   =np.zeros_like(S_Tree)
    S_H_V_1 =np.zeros_like(S_Tree)
    B_H_V   =np.zeros_like(S_Tree)
    B_H_V_1 =np.zeros_like(S_Tree)
    
    
    # **n 번째 마지막노드 값 채우기**
    
    # In[62]:
    
    
    RCPS_V[:,-1]     = RCPS_int_V[:,-1]
    D_Table[:,-1]    = (RCPS_V[:,-1]==S_int_V[:,-1])*10+(RCPS_V[:,-1]==B_int_V[:,-1])*20+((RCPS_V[:,-1]!=S_int_V[:,-1])*(RCPS_V[:,-1]!=B_int_V[:,-1]))*30    # 10 : 전환 , 20 : 상환 , 30 : 보유
    
    S_H_V  [:,-1]    = (D_Table[:,-1]==10)*S_int_V[:,-1]
    S_H_V_1[:,-1]    = (D_Table[:,-1]==10)*S_int_V[:,-1]
    B_H_V  [:,-1]    = (D_Table[:,-1]==20)*B_int_V[:,-1]
    B_H_V_1[:,-1]    = (D_Table[:,-1]==20)*B_int_V[:,-1]
    
    RCPS_H_V[:,-1] = S_H_V[:,-1]+B_H_V[:,-1]
    
    
    # In[63]:
    
    
    # 마지막노드 전값부터 처음값까지 채우기
    
    Rf_discount=np.exp(-Rf_F/100*dt) #한구간 할인이니 dt만큼만 할인하면 됨
    Kd_discount=np.exp(-ADJ_Kd_F/100*dt)*(ADJ_Kd_F>0)
    Reverse_Call_table=np.triu(np.tile(Reverse_Call_array,(n+1,1)))
    
    for i in np.arange(n,0,-1):
        S_H_V[:i,i-1]    = (S_H_V_1[:i,i]*p[i]+S_H_V_1[1:i+1,i]*q[i])*Rf_discount[i]
        B_H_V[:i,i-1]    = (B_H_V_1[:i,i]*0.5+B_H_V_1[1:i+1,i]*0.5)*Kd_discount[:i,i]
    
        RCPS_H_V[:i,i-1] = S_H_V[:i,i-1]+B_H_V[:i,i-1]
        
        if i>=int(Period_of_call):
            RCPS_V[:i,i-1]     = np.maximum(RCPS_int_V[:i,i-1],np.minimum(RCPS_H_V[:i,i-1],Reverse_Call_table[:i,i-1]))
        else:
            RCPS_V[:i,i-1]     = np.maximum(RCPS_int_V[:i,i-1],RCPS_H_V[:i,i-1])
        
        D_Table[:i,i-1]  = (RCPS_V[:i,i-1]==S_int_V[:i,i-1])*10+\
                           (RCPS_V[:i,i-1]==B_int_V[:i,i-1])*20+\
                           (RCPS_V[:i,i-1]==Reverse_Call_array[i-1])*20+\
                           ((RCPS_V[:i,i-1]!=S_int_V[:i,i-1])*(RCPS_V[:i,i-1]!=B_int_V[:i,i-1])*(RCPS_V[:i,i-1]!=Reverse_Call_array[i-1]))*30    # 10 : 전환 , 20 : 상환 , 30 : 보유
    
        S_H_V_1[:i,i-1]  = (D_Table[:i,i-1]==10)*S_int_V[:i,i-1]+((D_Table[:i,i-1]!=10)*(D_Table[:i,i-1]!=20))*(S_H_V_1[:i,i]*p[i]+S_H_V_1[1:i+1,i]*q[i])*Rf_discount[i]
        B_H_V_1[:i,i-1]  = (D_Table[:i,i-1]==20)*B_int_V[:i,i-1]+((D_Table[:i,i-1]!=10)*(D_Table[:i,i-1]!=20))*(B_H_V_1[:i,i]*p[i]+B_H_V_1[1:i+1,i]*q[i])*Kd_discount[:i,i]
    
    
    
        
    wb.sheets[sheet_name].range('E130').value =RCPS_V[0,0]
    wb.sheets[sheet_name].range('G130').value=B_int_V[0,0]  ## 수의상환권, 상환권이 모두 없는 상태에서 채권내재가치가 일반사채가치와 동일
    return wb