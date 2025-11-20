# -*- coding: utf-8 -*-
"""
Created on Thu Aug 25 16:56:25 2022

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
# os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct_Py_Scripts//BJ")
BJ_rawdata=pd.ExcelFile("BJ Exceptions - FM and TM.xlsx")
BJ=BJ_rawdata.parse('BJ Data')


BJ["Type of Breach"]="Market Abuse (BlueJeans Recording)"
BJ["Severity"]=""
BJ["Accountability"]="NA"
BJ["Comments"]=""
BJ["Q_Y"]=BJ["Quarter"].astype(str)+" "+BJ["Year"].astype(str)
BJ=BJ[["Bank ID","Type of Breach","Severity","Accountability","Q_Y","Comments","Defect Count","Product Desk"]]
BJ=BJ[np.invert(BJ["Product Desk"].astype(str).str.contains("Treasury"))]

BJ=BJ[["Bank ID","Type of Breach","Severity","Accountability","Q_Y","Comments","Defect Count"]]
BJ=BJ.groupby(["Bank ID","Type of Breach","Severity","Accountability","Comments","Q_Y"]).sum().reset_index()
BJ["Comments"]=BJ["Defect Count"].astype(str)+" breach(es) in a quarter"
BJ=BJ[["Bank ID","Type of Breach","Severity","Accountability","Q_Y","Comments"]]
BJ = BJ.set_axis(axis=1,labels=["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment"])

BJ["Date"]=np.where(BJ["Date"].str[:2]=="Q1","31 March "+BJ["Date"].str[-4:],BJ["Date"])
BJ["Date"]=np.where(BJ["Date"].str[:2]=="Q2","30 June "+BJ["Date"].str[-4:],BJ["Date"])
BJ["Date"]=np.where(BJ["Date"].str[:2]=="Q3","30 September "+BJ["Date"].str[-4:],BJ["Date"])
BJ["Date"]=np.where(BJ["Date"].str[:2]=="Q4","31 December "+BJ["Date"].str[-4:],BJ["Date"])


BJ=BJ[["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment"]]
BJ.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\BJ\BJ_Output.xlsx",index=False)
BJ.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Final\BJ_Output.xlsx",index=False)

