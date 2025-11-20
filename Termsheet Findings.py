# -*- coding: utf-8 -*-
"""
Created on Tue Jun 25 12:15:52 2024

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


os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct//Termsheet Findings")
# Getting the Base Datasets
# Raw_FMSWData
Raw_FMSWData = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct//Termsheet Findings//FMSW Termsheet Findings.xlsx')
Raw_FMSWData = pd.read_excel(Raw_FMSWData, 'FMSW Termsheet Findings')
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
output['Quarter'] = pd.PeriodIndex(Raw_FMSWData['Review Month'], freq='Q')
output['Date'] = Raw_FMSWData['Review Month']
output[['Employee PSID', 'Name']] = Raw_FMSWData['Marketer Name'].str.split("-", expand = True)
output[['LM PSID', 'LM Name']] = Raw_FMSWData['LM Name'].str.split("-", expand = True)
output['Type of Breach'] = 'Termsheet Findings'
output['Sub categories'] = 'Compliance Risk'
output['Severity'] = Raw_FMSWData['Breach Category']
output['Accountability'] = 'NA'
output['Materiality'] = output['Type of Breach'] + output['Severity']
output['Material?'] = Raw_FMSWData['Breach Impact Category after Sales Explanations']
output['Comment'] = Raw_FMSWData['Remarks (Sales\' explanations and whether repeat offender)']
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
       'Type of Breach', 'Sub categories', 'Severity', 'Accountability',
       'Materiality', 'Material?', 'Comment', 'FMSW/non-FMSW', 'Location',
       'Region', 'Business (Lvl 6)', 'Role'], axis='columns', inplace=True)


Stafflist_LM = Stafflist[['Bank Id', 'Staff Country', 'Staff Region']]
Stafflist_LM.set_axis(['LM PSID', 'LM Location', 'LM Region'], axis='columns', inplace=True)
output['LM PSID'] = output['LM PSID'].astype(int)
output = pd.merge(output, Stafflist_LM, on="LM PSID", how="left")
# cols = ['Location_y']
# output = output.drop(cols, axis=1)
output.set_axis(['Quarter', 'Date', 'Employee PSID', 'Name', 'LM PSID', 'LM Name',
       'Type of Breach', 'Sub categories', 'Severity', 'Accountability',
       'Materiality', 'Material?', 'Comment', 'FMSW/non-FMSW', 'Location',
       'Region', 'Business (Lvl 6)', 'Role', 'LM Location', 'LM Region'], axis='columns', inplace=True)

# Rearrange column names
output = output[['Quarter',	'Date',	'Employee PSID',	'Name',	'Location',	'Region',	'LM PSID',	'LM Name',	
                 'LM Location',	'LM Region',	'Business (Lvl 6)',	'Type of Breach',	'Sub categories',	
                 'Severity',	'Accountability',	'Materiality',	'Material?',	'Comment',	'FMSW/non-FMSW',	'Role']]

# Creating new dataset for reference finally
Termsheet_Findings_Output = output
Termsheet_Findings_Output.to_excel("Termsheet_Findings_Output.xlsx")



