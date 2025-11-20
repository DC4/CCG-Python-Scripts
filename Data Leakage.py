# -*- coding: utf-8 -*-
"""
Created on Tue Aug 16 18:01:58 2022

@author: 1659765
"""

import pandas as pd
import numpy as np
import datetime
import os
from datetime import date

def previous_quarter(ref):
    if ref.month < 4:
        return datetime.date(ref.year - 2, 12, 31)
    elif ref.month < 7:
        return datetime.date(ref.year-1, 3, 31)
    elif ref.month < 10:
        return datetime.date(ref.year-1, 6, 30)
    return datetime.date(ref.year-1, 9, 30)

cut_offdate=previous_quarter(date.today())
cut_offdate = pd.to_datetime(cut_offdate, format='%Y-%m-%d')

data_leakage=pd.read_excel("MI Reports by Business Function L2 - New HR Format.xlsx")
data_leakage["Type of Breach"]="Data Leakage"
data_leakage["Accountability"]="NA"
data_leakage=data_leakage[["Case ID","UserID","Type of Breach","Grade","Accountability","Closed Date","Recommendation"]]
data_leakage.drop_duplicates(inplace=True)
data_leakage=data_leakage.drop(['Case ID'],axis=1)
data_leakage = data_leakage.set_axis(axis=1,labels=["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment"])
data_leakage["Date"]=pd.to_datetime(data_leakage["Date"],format=r'%m/%d/%Y %I:%M:%S %p')

data_leakage=data_leakage[(data_leakage["Date"]>cut_offdate)]
data_leakage=data_leakage[data_leakage['Date'].notnull()]
data_leakage=data_leakage[np.invert(data_leakage["Severity"]=='Non-Event')]
data_leakage["Employee PSID"]=data_leakage["Employee PSID"].astype(int).astype(str)

data_leakage.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Data Leakage\Data_Leakage_Output.xlsx",index=False)
data_leakage.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Final\Data_Leakage_Output.xlsx",index=False)