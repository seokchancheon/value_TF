#!/usr/bin/env python
# coding: utf-8

# 

# # 부스트래핑 함수를 만들어 사용하기 

# ### 함수 만들어 사용하기 [refer](https://wikidocs.net/24)

# In[1]:


def FYS_Rate(Matuality,YTM, Num_node,freq):
    from scipy import interpolate
    import pandas as pd
    import numpy as np
    np.set_printoptions(precision=14)
    linear_func = interpolate.interp1d(Matuality, YTM, kind='linear')### 기본은 연단위로 그린거고
    new_Matuality=[]
    for i in range(0,Num_node+1): #새로운 node list를 만듬   freq를 고려하여 새로은 단위로
        new_Matuality.append(i/freq)
    YTM = linear_func(new_Matuality)
    # Bootstrapping   ## YTM단위를 %로 100을 나눠서 맞춰줌
    Spot_rate = np.zeros_like(YTM)  # SPOT이자율
    Forward_Rate = np.zeros_like(YTM)  # 선도 이자율
    principal_and_interest = np.zeros_like(YTM)  # 만기원리금
    PV_interest = np.zeros_like(YTM)  # 이자현가
    PV_principal_and_interest = np.zeros_like(YTM)  # 만기원리금 현가
    PV_factor = np.zeros_like(YTM)  # 만기원리금 현가
    sum = 0.
    ##시작 값 0번 라인 설정
    Spot_rate[0] = YTM[0]
    Forward_Rate[0] = YTM[0]
    for i in range(1, Num_node + 1):
        principal_and_interest[i] = 1 + ((YTM[i] / freq) / 100)  # 주기로 나눠줘야함
        PV_interest[i] = 1 * ((YTM[i] / freq) / 100) * sum  # 주기로 나눠줘야함
        PV_principal_and_interest[i] = 1 - PV_interest[i]
        Spot_rate[i] = ((principal_and_interest[i] / PV_principal_and_interest[i]) ** (1/ i) - 1) * 100*freq  
        # freq로 YTM을 나눈후의 이자율이니 freq승을 다시해 줘야 1년이자율이 됨
        if YTM[i]==YTM[i-1]:
            Spot_rate[i]=Spot_rate[i-1]
        
        PV_factor[i] = 1 / (1 + (Spot_rate[i] / 100)) ** ((i) / freq)
        sum = sum + PV_factor[i]
    for i in range(1, Num_node + 1):
        Forward_Rate[i] = (((1 + (Spot_rate[i] / 100)) ** (i / freq) / ((1 + Spot_rate[i - 1] / 100) ** ((i - 1) / freq))) - 1) * 100 * (freq)
        if YTM[i]==YTM[i-1]:
            Forward_Rate[i]=Forward_Rate[i-1]    
    
    return Forward_Rate, YTM, Spot_rate


# In[ ]:





# In[1]:


def FYS_Rate_Kd(Matuality,YTM, Num_node,freq):
    from scipy import interpolate
    import pandas as pd
    import numpy as np
    np.set_printoptions(precision=14)
    linear_func = interpolate.interp1d(Matuality, YTM, kind='linear')### 기본은 연단위로 그린거고
    new_Matuality=[]
    for i in range(0,Num_node+1): #새로운 node list를 만듬   freq를 고려하여 새로은 단위로
        new_Matuality.append(i/freq)
    YTM = linear_func(new_Matuality)
    # Bootstrapping   ## YTM단위를 %로 100을 나눠서 맞춰줌
    Spot_rate = np.zeros_like(YTM)  # SPOT이자율
    Forward_Rate = np.zeros_like(YTM)  # 선도 이자율
    principal_and_interest = np.zeros_like(YTM)  # 만기원리금
    PV_interest = np.zeros_like(YTM)  # 이자현가
    PV_principal_and_interest = np.zeros_like(YTM)  # 만기원리금 현가
    PV_factor = np.zeros_like(YTM)  # 만기원리금 현가
    sum = 0.
    ##시작 값 0번 라인 설정
    Spot_rate[0] = YTM[0]
    Forward_Rate[0] = YTM[0]
    for i in range(1, Num_node + 1):
        principal_and_interest[i] = 1 + ((YTM[i] / freq) / 100)  # 주기로 나눠줘야함
        PV_interest[i] = 1 * ((YTM[i] / freq) / 100) * sum  # 주기로 나눠줘야함
        PV_principal_and_interest[i] = 1 - PV_interest[i]
        Spot_rate[i] = ((principal_and_interest[i] / PV_principal_and_interest[i]) ** (1/ i) - 1) * 100*freq  
        # freq로 YTM을 나눈후의 이자율이니 freq승을 다시해 줘야 1년이자율이 됨
        if YTM[i]==YTM[i-1]:
            Spot_rate[i]=Spot_rate[i-1]
        
        PV_factor[i] = 1 / (1 + (Spot_rate[i] / 100)) ** ((i) / freq)
        sum = sum + PV_factor[i]
    for i in range(1, Num_node + 1):
        Forward_Rate[i] = (((1 + (Spot_rate[i] / 100)) ** (i / freq) / ((1 + Spot_rate[i - 1] / 100) ** ((i - 1) / freq))) - 1) * 100 * (freq)
        if YTM[i]==YTM[i-1]:
            Forward_Rate[i]=Forward_Rate[i-1]    
    
    return Forward_Rate, YTM, Spot_rate

