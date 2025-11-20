# -*- coding: utf-8 -*-
"""
Created on Mon Apr  8 18:46:45 2024

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
import datetime

os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct//Ops Loss")

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

# Getting the Base Datasets
# FMSWData_12
Ops = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct//Ops Loss//Ops.xlsx')
Ops = pd.read_excel(Ops, 'Sheet1')
print("Shape of Ops:", Ops.shape)

# Stafflist
Raw_Stafflist = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct//Stafflist.xlsx')
Stafflist = pd.read_excel(Raw_Stafflist, 'Sheet1')
print("Shape of Stafflist:", Stafflist.shape)

# Building output Dataset
output = pd.DataFrame()
output['Quarter'] = pd.PeriodIndex(Ops['Date'], freq='Q')
output['Date'] = Ops['Date']
output['Date'] = Ops['Date']
output['Employee PSID'] = Ops['Employee PSID']
output['Type of Breach'] = 'Operational Errors (FO) >$50k'
output['Sub categories'] = 'Operational Risk'
output['Severity'] = 'High'
output['Accountability'] = Ops['Accountability'].replace("", np.NaN).fillna('NA')
output['Materiality'] = 'Operational Errors (FO) >$50k High'
output['Material?'] = 'Material'
output['Comment'] = Ops['Comment']
output['FMSW/non-FMSW'] = 'non-FMSW'

# Rolling 12 months logic
output = output[output['Date'] > cut_offdate]

# Fetching all data from Stafflist
# To build the columns 'LM Location', 'LM Region' from Stafflist
Stafflist_Temp = Stafflist[['Bank Id', 'Staff Name', 'Staff Country', 'Staff Region', 'Supervisor Id', 'Supervisor Name', 'Business Function Level 6 Desc', 'Role']]
Stafflist_Temp.set_axis(['Employee PSID', 'Name', 'Location', 'Region', 'LM PSID', 'LM Name', 'Business (Lvl 6)', 'Role'], axis='columns', inplace=True) 
Stafflist_Temp['Employee PSID'] = Stafflist_Temp['Employee PSID'].astype(int)
Stafflist_Temp['LM PSID'] = Stafflist_Temp['LM PSID'].astype("Int64")
output['Employee PSID'] = output['Employee PSID'].astype(int)
Stafflist_Temp.drop_duplicates(inplace=True)
output.drop_duplicates(inplace=True)
output = pd.merge(output, Stafflist_Temp, on="Employee PSID", how="left")
# cols = ['Region_y', 'LM PSID_y', 'LM Name_y']
# cols = ['Region_y', 'Dummy', 'Workflow Status']
# output = output.drop(cols, axis=1)
# print(output.columns)
# output.to_excel("output.xlsx")
# output.set_axis(['Employee PSID', 'Region', 'LM PSID', 'LM Name', 'Business (Lvl 6)', 'Role'], axis='columns', inplace=True) 
Stafflist_LM = Stafflist[['Bank Id', 'Staff Country', 'Staff Region']]
Stafflist_LM.set_axis(['LM PSID', 'LM Location', 'LM Region'], axis='columns', inplace=True)
Stafflist_LM['LM PSID'] = Stafflist_LM['LM PSID'].astype("Int64")
output['LM PSID'] = output['LM PSID'].astype("Int64")
output = pd.merge(output, Stafflist_LM, on="LM PSID", how="left")


# Trim the spaces in all column values
print("Add in all the column names to trim")
# output['Quarter'] = output['Quarter'].apply(lambda x: x.strip())
# output['Date'] = output['Date'].apply(lambda x: x.strip())
# output['Employee PSID'] = output['Employee PSID'].apply(lambda x: x.strip())
# output['Name'] = output['Name'].apply(lambda x: x.strip())
# output['Location'].apply(lambda x: x.strip())
# output['Region'] = output['Region'].apply(lambda x: x.strip())
# output['LM PSID'] = output['LM PSID'].apply(lambda x: x.strip())
# output['LM PSID'] = output['LM PSID'].apply(lambda x: x.strip())
# output['LM Name'] = output['LM Name'].apply(lambda x: x.strip())
# output['LM Location'] = output['LM Location'].apply(lambda x: x.strip())
# output['LM Region'] = output['LM Region'].apply(lambda x: x.strip())
# output['Business (Lvl 6)'] = output['Business (Lvl 6)'].apply(lambda x: x.strip())
output['Type of Breach'] = output['Type of Breach'].apply(lambda x: x.strip())
output['Sub categories'] = output['Sub categories'].apply(lambda x: x.strip())
output['Severity'] = output['Severity'].apply(lambda x: x.strip())
output['Accountability'] = output['Accountability'].apply(lambda x: x.strip())
output['Materiality'] = output['Materiality'].apply(lambda x: x.strip())
output['Material?'] = output['Material?'].apply(lambda x: x.strip())
output['Comment'] = output['Comment'].apply(lambda x: x.strip())
output['FMSW/non-FMSW'] = output['FMSW/non-FMSW'].apply(lambda x: x.strip())
# output['Role'] = output['Role'].apply(lambda x: x.strip())

# Rearrange column names
output = output[['Quarter',	'Date',	'Employee PSID',	'Name',	'Location',	'Region',	'LM PSID',	'LM Name',	
                 'LM Location',	'LM Region',	'Business (Lvl 6)',	'Type of Breach',	'Sub categories',	'Severity',	
                 'Accountability',	'Materiality',	'Material?',	'Comment',	'FMSW/non-FMSW',	'Role']]

# Creating new dataset for reference finally
Ops_Loss_Output = output
Ops_Loss_Output.to_excel("Ops_Loss_Output.xlsx")