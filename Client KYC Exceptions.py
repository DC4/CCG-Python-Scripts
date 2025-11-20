# -*- coding: utf-8 -*-
"""
Created on Mon Jun 24 16:31:17 2024

@author: 1510806
"""

# import pandas
import pandas as pd
import numpy as np
import os
import time
import openpyxl
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Font
import pandas as pd
import win32com.client as win32
import time
import sys
import subprocess
import os
from datetime import date
import numpy as np
import win32com.client


os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct//Client KYC Exceptions")
# Getting the Base Datasets
# Raw_FMSWData
Raw_FMSWData = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct//Client KYC Exceptions//FMSW Client KYC Exceptions.xlsx')
Raw_FMSWData = pd.read_excel(Raw_FMSWData, 'FMSW Client KYC Exceptions')
print("Shape of Raw_FMSWData:", Raw_FMSWData.shape)

# Mapping
Raw_Mapping = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct//Mapping.xlsx')
Mapping = pd.read_excel(Raw_Mapping, 'Sheet1')
print("Shape of Mapping:", Mapping.shape)

# Stafflist
Raw_Stafflist = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct//Stafflist.xlsx')
Stafflist = pd.read_excel(Raw_Stafflist, 'Sheet1')
print("Shape of Stafflist:", Stafflist.shape)

# Building output Dataset
output = pd.DataFrame()
# Fetching all data from Raw_FMSWData
output['Quarter'] = pd.PeriodIndex(Raw_FMSWData['Workflow Initiated (UTC)'], freq='Q')
output['Date'] = Raw_FMSWData['Workflow Initiated (UTC)']
output[['Employee PSID', 'Name']] = Raw_FMSWData['Employee PSID - Name'].str.split("-", expand = True)
output[['LM PSID', 'LM Name']] = Raw_FMSWData['Primary FO LM'].str.split("-", expand = True)
output['Type of Breach'] = 'Client KYC Exceptions'
output['Sub categories'] = 'Compliance Risk'
output['Severity'] = "Rag - " + Raw_FMSWData['MI RAG Status']
output['Accountability'] = Raw_FMSWData['Disciplinary Action Taken']
output['Materiality'] = output['Type of Breach'] + output['Severity']
output['Comment'] = Raw_FMSWData['FO Staff Justification Comments']
output['FMSW/non-FMSW'] = 'FMSW'

# output['Material?']
def Categorize_Severity(Severity):
    if Severity in ("RED"):
        return 'Material'
    elif Severity in ("AMBER") or Severity in ("GREEN"):
        return 'Non-Material'
    else:
        return 'Uncategorized'
    
output['Material?'] = Raw_FMSWData['MI RAG Status'].apply(Categorize_Severity)
# output.to_excel("Output_Missed_Trade_Temporary.xlsx")


# Fetching all data from Stafflist
# To build the columns 'LM Location', 'LM Region' from Stafflist
Stafflist_Temp = Stafflist[['Bank Id', 'Staff Country', 'Staff Region', 'Supervisor Id', 'Supervisor Name', 'Business Function Level 6 Desc', 'Role']]
Stafflist_Temp.set_axis(['Employee PSID', 'Location', 'Region', 'LM PSID', 'LM Name', 'Business (Lvl 6)', 'Role'], axis='columns', inplace=True) 
Stafflist_Temp['Employee PSID'] = Stafflist_Temp['Employee PSID'].astype(int)
Stafflist_Temp['LM PSID'] = Stafflist_Temp['LM PSID'].astype("Int64")
output['Employee PSID'] = output['Employee PSID'].astype(int)
Stafflist_Temp.drop_duplicates(inplace=True)
output.drop_duplicates(inplace=True)
output = pd.merge(output, Stafflist_Temp, on="Employee PSID", how="left")
# cols = ['Region_y', 'LM PSID_y', 'LM Name_y']
cols = ['LM PSID_y', 'LM Name_y']
output = output.drop(cols, axis=1)
output.set_axis(['Quarter', 'Date', 'Employee PSID', 'Name', 'LM PSID', 'LM Name',
       'Type of Breach', 'Sub categories', 'Severity', 'Accountability',
       'Materiality', 'Comment', 'FMSW/non-FMSW', 'Material?', 'Location',
       'Region', 'Business (Lvl 6)', 'Role'], axis='columns', inplace=True)
Stafflist_LM = Stafflist[['Bank Id', 'Staff Country', 'Staff Region']]
Stafflist_LM.set_axis(['LM PSID', 'LM Location', 'LM Region'], axis='columns', inplace=True)
output['LM PSID'] = output['LM PSID'].astype(int)
output = pd.merge(output, Stafflist_LM, on="LM PSID", how="left")
# cols = ['Location_y']
# output = output.drop(cols, axis=1)
output.set_axis(['Quarter', 'Date', 'Employee PSID', 'Name', 'LM PSID', 'LM Name',
       'Type of Breach', 'Sub categories', 'Severity', 'Accountability',
       'Materiality', 'Comment', 'FMSW/non-FMSW', 'Material?', 'Location',
       'Region', 'Business (Lvl 6)', 'Role', 'LM Location', 'LM Region'], axis='columns', inplace=True)


# Rearrange column names
output = output[['Quarter',	'Date',	'Employee PSID',	'Name',	'Location',	'Region',	'LM PSID',	'LM Name',	
                 'LM Location',	'LM Region',	'Business (Lvl 6)',	'Type of Breach',	'Sub categories',	
                 'Severity',	'Accountability',	'Materiality',	'Material?',	'Comment',	'FMSW/non-FMSW',	'Role']]

# Creating new dataset for reference finally
Client_KYC_Exceptions_Output = output
Client_KYC_Exceptions_Output.to_excel("Client_KYC_Exceptions_Output.xlsx")















