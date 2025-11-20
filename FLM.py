# -*- coding: utf-8 -*-
"""
Created on Thu Dec 19 13:46:10 2024

@author: 1510806
"""
# -*- coding: utf-8 -*-
"""
Created on Thu Mar 28 14:58:41 2024

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

os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct//Front Line Monitoring - Non-FMSW")

# Getting the Base Datasets
# FLM
FLM = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct//FLM.xlsx')
FLM = pd.read_excel(FLM, 'Sheet1')
print("Shape of FLM:", FLM.shape)
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
# Fetching all data from FLM
output['Quarter'] = pd.PeriodIndex(FLM['Date'], freq='Q')
output['Date'] = FLM['Date']
output['Employee PSID'] = FLM['Employee PSID']
output['Type of Breach'] = 'Front Line Monitoring'
output['Sub categories'] = 'Compliance Risk'
output['Severity'] = 'Medium - High'
output['Accountability'] ='NA'
output['Materiality'] = 'Front Line Monitoring Medium - High'
output['Material?'] = 'Material'
output['Comment'] = FLM['Comment']
output['FMSW/non-FMSW'] = 'non-FMSW'


# output[['Employee PSID', 'Name', 'Dummy']] = FLM['PSID - Name'].str.split("-", expand = True)
# output['Location'] = FLM['Responsible User\'s Country']
# output['Business (Lvl 6)']
# # Retain rows with only RED RAG Status
# output = output[output['Severity'].str.contains("RED", na=False, case=False)]

# Fetching all data from Stafflist
# To build the columns 'LM Location', 'LM Region' from Stafflist

Stafflist_Temp = Stafflist[['Bank Id', 'Staff Name', 'Staff Country', 'Staff Region', 'Supervisor Id', 'Supervisor Name', 'Business Function Level 6 Desc', 'Role']]
Stafflist_Temp.set_axis(['Employee PSID', 'Name', 'Location', 'Region', 'LM PSID', 'LM Name', 'Business (Lvl 6)', 'Role'], axis='columns', inplace=True) 
Stafflist_Temp['Employee PSID'] = Stafflist_Temp['Employee PSID'].astype(int)
Stafflist_Temp['LM PSID'] = Stafflist_Temp['LM PSID'].astype("Int64")
output['Employee PSID'] = output['Employee PSID'].astype(int)
output = pd.merge(output, Stafflist_Temp, on="Employee PSID", how="left")
Stafflist_LM = Stafflist[['Bank Id', 'Staff Country', 'Staff Region']]
Stafflist_LM.set_axis(['LM PSID', 'LM Location', 'LM Region'], axis='columns', inplace=True) 
output = pd.merge(output, Stafflist_LM, on="LM PSID", how="left")

# # Trim the spaces in all column values
# print("Add in all the column names to trim")
# # output['Quarter'] = output['Quarter'].apply(lambda x: x.strip())
# output['Date'] = output['Date'].apply(lambda x: x.strip())
# # output['Employee PSID'] = output['Employee PSID'].apply(lambda x: x.strip())
# output['Name'] = output['Name'].apply(lambda x: x.strip())
# output['Location'] = output['Location'].apply(lambda x: x.strip())
# # output['LM PSID'] = output['LM PSID'].apply(lambda x: x.strip())
# output['Type of Breach'] = output['Type of Breach'].apply(lambda x: x.strip())
# output['Sub categories'] = output['Sub categories'].apply(lambda x: x.strip())
# output['Severity'] = output['Severity'].apply(lambda x: x.strip())
# output['Accountability'] = output['Accountability'].apply(lambda x: x.strip())
# output['Materiality'] = output['Materiality'].apply(lambda x: x.strip())
# output['Material?'] = output['Material?'].apply(lambda x: x.strip())
# output['Comment'] = output['Comment'].apply(lambda x: x.strip())
# output['FMSW/non-FMSW'] = output['FMSW/non-FMSW'].apply(lambda x: x.strip())
# output['Region'] = output['Region'].apply(lambda x: x.strip())
# # output['LM PSID'] = output['LM PSID'].apply(lambda x: x.strip())
# output['LM Name'] = output['LM Name'].apply(lambda x: x.strip())
# output['LM Location'] = output['LM Location'].apply(lambda x: x.strip())
# output['LM Region'] = output['LM Region'].apply(lambda x: x.strip())
# output['Business (Lvl 6)'] = output['Business (Lvl 6)'].apply(lambda x: x.strip())
# output['Role'] = output['Role'].apply(lambda x: x.strip())

# Rearrange column names
output = output[['Quarter',	'Date',	'Employee PSID',	'Name',	'Location',	'Region',	'LM PSID',	'LM Name',	
                 'LM Location',	'LM Region',	'Business (Lvl 6)',	'Type of Breach',	'Sub categories',	'Severity',	
                 'Accountability',	'Materiality',	'Material?',	'Comment',	'FMSW/non-FMSW',	'Role']]

# Creating new dataset for reference finally
FLM_Out = output
FLM_Out.to_excel("FLM_Output.xlsx")

########################################################################################################################################################

