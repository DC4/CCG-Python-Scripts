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

# Cuttoffdate logic
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


EOD_rawdata=pd.ExcelFile("EOD Exceptions - Consolidated East and West Results.xlsx")
EOD_East=EOD_rawdata.parse('EAST EOD Recap')
EOD_East=EOD_East.loc[EOD_East.index.repeat(EOD_East["Defect count"])]
EOD_East["Type of Breach"]="Market Abuse (EOD Update)"
EOD_East["Severity"]=""
EOD_East["Accountability"]="NA"
EOD_East=EOD_East[["Bank ID","Type of Breach","Severity","Accountability","Period 2","EOD email summary","Product Desk"]]
EOD_East = EOD_East.set_axis(axis=1,labels=["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment","Product Desk"])


EOD_West=EOD_rawdata.parse('WEST EOD Recap')
EOD_West["Q_Y"]=EOD_West["Quarter"].astype(str)+EOD_West["Year"].astype(str)
EOD_West=EOD_West.loc[EOD_West.index.repeat(EOD_West["Defect Count"])]
EOD_West["Type of Breach"]="Market Abuse (EOD Update)"
EOD_West["Severity"]=""
EOD_West["Accountability"]="NA"
EOD_West=EOD_West[["Bank ID","Type of Breach","Severity","Accountability","Period 2","Comments","Product Desk"]]
EOD_West = EOD_West.set_axis(axis=1,labels=["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment","Product Desk"])

EOD=pd.concat([EOD_West,EOD_East])

EOD=EOD[(EOD["Date"]>cut_offdate)]

EOD=EOD.sort_values(by=["Employee PSID","Date"]).reset_index(drop=True)
EOD["Cumulative Count"]=EOD.groupby("Employee PSID").cumcount()+1

EOD_reverse=EOD.sort_values(by=["Employee PSID","Date"],ascending=[True,False]).reset_index(drop=True)
# EOD["Time Diff"]=EOD["Date"].dt.to_period('M')-EOD_reverse["Date"].dt.to_period('M')
EOD["Time Diff"]=EOD["Date"].dt.to_period('M').astype(int)-EOD_reverse["Date"].dt.to_period('M').astype(int)
EOD["Severity"]=np.where(np.logical_and(EOD["Time Diff"]<=6,EOD["Cumulative Count"]>=5),"5 or more period 6 mths","")
EOD["Severity"]=np.where(EOD["Cumulative Count"]>=7,"7 or more period 12 mths",EOD["Severity"])

EOD=EOD[np.invert(EOD["Product Desk"].str.contains("Treasury"))]
EOD=EOD[np.invert(EOD["Product Desk"].str.contains("TM"))]

EOD=EOD[["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment"]]

EOD.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\EOD\EOD_Output.xlsx",index=False)
EOD.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Final\EOD_Output.xlsx",index=False)