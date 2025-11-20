# -*- coding: utf-8 -*-
"""
Created on Tue Aug 16 18:01:58 2022

@author: 1659765
"""

import pandas as pd
import numpy as np

#%%As the sender is different every quarter the format will change too
#==============================================================================
# data_breach=pd.read_excel("CCIB FM Q3 2021 - Q2 2022_Contacts.xlsx")
# data_breach_name=data_breach[["Incident ID","Involved Staff ID","Involved Staff Name"]]
# data_breach_name=data_breach_name.dropna()
# data_breach_comment=pd.read_excel("CCIB FM Q3 2021 - Q2 2022_Reporter's summary.xlsx",skiprows=2)
# data_breach_comment=data_breach_comment.drop(['Incident Status'],axis=1)
# data_breach=data_breach.drop(["Involved Staff ID","Involved Staff Name"],axis=1)
# data_breach=data_breach.fillna(method='ffill')
# data_breach=data_breach.join(data_breach_name.set_index('Incident ID'),on='Incident ID')
# data_breach=data_breach[data_breach['Involved Staff ID'].notnull()]
# data_breach=data_breach.join(data_breach_comment.set_index('Incident ID'),on='Incident ID')
# data_breach["Type of Breach"]="Data Breach"
# data_breach["Accountability"]="NA"
# data_breach=data_breach[["Involved Staff ID","Type of Breach","Business Risk Rating by BRM","Accountability","Month","Reporter's Summary"]]
# data_breach.set_axis(axis=1,labels=["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment"])
# 
# data_breach["Severity"]=data_breach["Severity"].str.split('(',expand=True)[1].str.split(')',expand=True)[0].str.strip()
# data_breach["Employee PSID"]=data_breach["Employee PSID"].astype(int).astype(str)
# 
# 
#==============================================================================
data_breach=pd.read_excel("Data Breach.xlsx")
data_breach["Severity"]=data_breach["Severity"].str.split('(',expand=True)[1].str.split(')',expand=True)[0].str.strip()
data_breach["Employee PSID"]=data_breach["Employee PSID"].astype(str)
data_breach.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Data Breach\Data_Breach_Output.xlsx",index=False)
data_breach.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Final\Data_Breach_Output.xlsx",index=False)