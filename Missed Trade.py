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

os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct//Missed_Trade")
# Getting the Base Datasets
# FMSWData_12
Raw_FMSWData_12 = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct//FMSWData_12.xlsx')
FMSWData_12 = pd.read_excel(Raw_FMSWData_12, 'Missed Trade')
print("Shape of FMSWData_12:", FMSWData_12.shape)
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
# Fetching all data from FMSWData_12
output['Quarter'] = pd.PeriodIndex(FMSWData_12['Workflow Initiated (UTC)'], freq='Q')
output['Date'] = FMSWData_12['Workflow Initiated (UTC)']
output['Employee PSID'] = FMSWData_12['Responsible Employee PSID - Name ']
output[['Employee PSID', 'Name']] = FMSWData_12['Responsible Employee PSID - Name '].str.split("-", expand = True)
output['Location'] = FMSWData_12['Responsible Staff Location']
output['Region'] = FMSWData_12['Responsible Staff Region']
output['LM PSID'] = FMSWData_12['Responsible Staff LM PSId']
output['LM Name'] = FMSWData_12['Responsible Staff LM Name']
output['Severity'] = FMSWData_12['Severity']
output['Accountability'] = FMSWData_12['Fair Accountability Outcome']
output['Comment'] = FMSWData_12['Staff Responsible Comments']
output['Type of Breach'] = 'Missed trade'
output['Sub categories'] = 'Operational Risk'
output['FMSW/non-FMSW'] = 'FMSW'

# output['Material?']
def Categorize_Severity(Severity):
    if Severity in ("High", "high") or Severity in ("Medium", "medium"):
        return 'Material'
    elif Severity in ("Low", "low"):
        return 'Non-Material'
    else:
        return 'Uncategorized'
    
output['Material?'] = output['Severity'].apply(Categorize_Severity)
# output.to_excel("Output_Missed_Trade_Temporary.xlsx")


# Fetching all data from Stafflist
# To build the columns 'LM Location', 'LM Region' from Stafflist
Stafflist_Temp = Stafflist[['Bank Id', 'Staff Region', 'Supervisor Id', 'Supervisor Name', 'Business Function Level 6 Desc', 'Role']]
Stafflist_Temp.set_axis(['Employee PSID', 'Region', 'LM PSID', 'LM Name', 'Business (Lvl 6)', 'Role'], axis='columns', inplace=True) 
Stafflist_Temp['Employee PSID'] = Stafflist_Temp['Employee PSID'].astype(int)
Stafflist_Temp['LM PSID'] = Stafflist_Temp['LM PSID'].astype("Int64")
output['Employee PSID'] = output['Employee PSID'].astype(int)
Stafflist_Temp.drop_duplicates(inplace=True)
output.drop_duplicates(inplace=True)
output = pd.merge(output, Stafflist_Temp, on="Employee PSID", how="left")
cols = ['Region_y', 'LM PSID_y', 'LM Name_y']
output = output.drop(cols, axis=1)
output.set_axis(['Quarter', 'Date', 'Employee PSID', 'Name', 'Location', 'Region',
       'LM PSID', 'LM Name', 'Severity', 'Accountability', 'Comment',
       'Type of Breach', 'Sub categories', 'FMSW/non-FMSW', 'Material?',
        'Business (Lvl 6)', 'Role'], axis='columns', inplace=True) 
Stafflist_LM = Stafflist[['Bank Id', 'Staff Country', 'Staff Region']]
Stafflist_LM.set_axis(['LM PSID', 'LM Location', 'LM Region'], axis='columns', inplace=True) 
output = pd.merge(output, Stafflist_LM, on="LM PSID", how="left")
# output.to_excel("output.xlsx")

# Trim the spaces in all column values
print("Add in all the column names to trim")
# output['Employee PSID'] = output['Employee PSID'].apply(lambda x: x.strip())
# output['Quarter'] = output['Quarter'].apply(lambda x: x.strip())
output['Date'] = output['Date'].apply(lambda x: x.strip())
output['Name'] = output['Name'].apply(lambda x: x.strip())
output['Location'] = output['Location'].apply(lambda x: x.strip())
output['Region'] = output['Region'].apply(lambda x: x.strip())
# output['LM PSID'] = output['LM PSID'].apply(lambda x: x.strip())
output['LM Name'] = output['LM Name'].apply(lambda x: x.strip())
output['Severity'] = output['Severity'].apply(lambda x: x.strip())
output['Accountability'] = output['Accountability'].apply(lambda x: x.strip())
output['Comment'] = output['Comment'].apply(lambda x: x.strip())
output['Type of Breach'] = output['Type of Breach'].apply(lambda x: x.strip())
output['Sub categories'] = output['Sub categories'].apply(lambda x: x.strip())
output['FMSW/non-FMSW'] = output['FMSW/non-FMSW'].apply(lambda x: x.strip())
output['Material?'] = output['Material?'].apply(lambda x: x.strip())
output['LM Location'] = output['LM Location'].apply(lambda x: x.strip())
output['LM Region'] = output['LM Region'].apply(lambda x: x.strip())
output['Business (Lvl 6)'] = output['Business (Lvl 6)'].apply(lambda x: x.strip())
output['Role'] = output['Role'].apply(lambda x: x.strip())

# Rearrange column names
output = output[['Quarter',	'Date',	'Employee PSID',	'Name',	'Location',	'Region',	'LM PSID',	
                 'LM Name',	'LM Location',	'LM Region',	'Business (Lvl 6)',	'Type of Breach',	
                 'Sub categories',	'Severity',	'Accountability',	'Material?',	
                 'Comment',	'FMSW/non-FMSW',	'Role']]

# Creating new dataset for reference finally
Missed_Trade_Output = output
Missed_Trade_Output.to_excel("Missed_Trade_Output.xlsx")


# Trim the column names if they have extra spaces
# output.rename(columns={'Responsible Employee PSID - Name  ': 'Responsible Employee PSID - Name'}, inplace=True)
# output = pd.merge(output, Stafflist_Subset, on="col1", how="left", indicator=True)
# output['LM Location'] = 
# output['LM Region'] =
# output['Business (Lvl 6)'] = 
# output['Role'] = 



