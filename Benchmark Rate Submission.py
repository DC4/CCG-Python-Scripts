# -*- coding: utf-8 -*-
"""
Created on Mon Jun 24 12:13:10 2024

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


os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct//Benchmark Rate Submission")
# Getting the Base Datasets
# Raw_FMSWData
Raw_FMSWData = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct//Benchmark Rate Submission//FMSW Benchmark Rate Submission.xlsx')
Raw_FMSWData = pd.read_excel(Raw_FMSWData, 'FMSW Benchmark Rate Submission')
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
output[['Employee PSID', 'Name']] = Raw_FMSWData['Responsible Employee PSID - Name '].str.split("-", expand = True)
output[['LM PSID', 'LM Name']] = Raw_FMSWData['Line Manager'].str.split("-", expand = True)
output['Severity'] = Raw_FMSWData['Issue Category']
output['Accountability'] = Raw_FMSWData['Fair Accountability Outcome']
output['Comment'] = Raw_FMSWData['Supervisor Remarks']
output['Type of Breach'] = 'Benchmark Rate Submission'
output['Sub categories'] = 'Compliance Risk'
output['FMSW/non-FMSW'] = 'FMSW'

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
       'Severity', 'Accountability', 'Comment', 'Type of Breach',
       'Sub categories', 'FMSW/non-FMSW', 'Location', 'Region',
       'Business (Lvl 6)', 'Role'], axis='columns', inplace=True) 
Stafflist_LM = Stafflist[['Bank Id', 'Staff Country', 'Staff Region']]
Stafflist_LM.set_axis(['LM PSID', 'LM Location', 'LM Region'], axis='columns', inplace=True)
output['LM PSID'] = output['LM PSID'].astype(int)
output = pd.merge(output, Stafflist_LM, on="LM PSID", how="left")
output['Location'] = output['Location'].apply(lambda x: x.strip())
output['Region'] = output['Region'].apply(lambda x: x.strip())
output['LM Location'] = output['LM Location'].apply(lambda x: x.strip())
output['LM Region'] = output['LM Region'].apply(lambda x: x.strip())
output['Business (Lvl 6)'] = output['Business (Lvl 6)'].apply(lambda x: x.strip())
output['Materiality'] = output['Type of Breach'] + output['Severity']
output['Issue Category'] = Raw_FMSWData['Issue Category']

# output['Material?']
def Categorize_Issue(Issue_Category):
    if Issue_Category in ("1A - Material Impact on Submission"):
        return 'Material'
    else:
        return 'Non-Material'
    
output['Material?'] = output['Issue Category'].apply(Categorize_Issue)
# output.to_excel("Output_Missed_Trade_Temporary.xlsx")
output['Role'] = output['Role'].apply(lambda x: x.strip())

# Rearrange column names
output = output[['Quarter',	'Date',	'Employee PSID',	'Name',	'Location',	'Region',	'LM PSID',	
                 'LM Name',	'LM Location',	'LM Region',	'Business (Lvl 6)',	'Type of Breach',	
                 'Sub categories',	'Severity',	'Accountability',	'Materiality',	'Material?',	
                 'Comment',	'FMSW/non-FMSW',	'Role']]

# Creating new dataset for reference finally
Benchmark_Rate_Submission_Output = output
Benchmark_Rate_Submission_Output.to_excel("Benchmark_Rate_Submission_Output.xlsx")