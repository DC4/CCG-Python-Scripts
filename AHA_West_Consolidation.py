# -*- coding: utf-8 -*-
"""
Created on Sun Sep  4 20:25:35 2022

@author: 1510806
"""

import openpyxl
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Font
import pandas as pd
import win32com.client as win32
import time
import subprocess
import os
import pandas as pd
import numpy as np
import xlsxwriter
from datetime import date
import numpy as np
from pandas.tseries.offsets import *

print("Change Date_filter to needed month start date")
print("Ensure consolidated file does not have any rows for the month we run Ex: Rows with 'Date' as Dec should not be there when running for Dec - this is to prevent duplicates")
Date_filter_1 = '2022-12-01'
# Date_filter_2 = '2022-06-30'

## os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//CCG//AHA//OLD_Last_3Months//2022.Sep Pack - Aug Data//East_West_Files")
os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Monthly_East_West_Files")

#########################################################################################################################################
#########################################################################################################################################
                                                    ########### EOD WEST ###########
#########################################################################################################################################
#########################################################################################################################################
                                
EOD_West = pd.read_excel('Covid 19 results West-2022.xlsx', sheet_name = "Eod Recap")

print("Length of EOD_West:", len(EOD_West))
# pandas convert column with integers to date time
print("Type of EOD_West[\"Date\"]:", type(EOD_West["Date"]))
EOD_West["Date"] = pd.to_datetime(EOD_West["Date"], errors='coerce')
# EOD_West = EOD_West[EOD_West["Date"].dt.strftime('%Y-%m-%d') > "2022-06-01"]
EOD_West = EOD_West[EOD_West["Date"].dt.strftime('%Y-%m-%d') >= Date_filter_1]
# Drop rows with NaN
EOD_West = EOD_West[~EOD_West["Date"].isnull()]
EOD_West.tail()
# Rearrange columns for consolidation
# Remove Existing incomplete or blank columns
# EOD_West = EOD_West.drop(['Region', 'Test'], axis=1)
EOD_West = EOD_West.drop(['Test'], axis=1)
# EOD_West = EOD_West.drop(['Test'], axis=1)
EOD_West.columns
EOD_West.tail()
# Adding needed columns
# EOD_West['Period 2'] = ''
# Building Period 2 Column
# from pandas.tseries.offsets import *
EOD_West["Period 2"] = EOD_West['Date']
EOD_West["Period 2"] = pd.to_datetime(EOD_West["Period 2"], errors='coerce')
# EOD_West["Period 2"] = EOD_West['Date'] + Week(weekday=4)
#EOD_West['Period'] = EOD_West['Period'].str.strip()
#EOD_West['Period 2'] = EOD_West['Period'].str[-9:]
#print (EOD_West['Period 2'].dtypes)
#EOD_West['Period 2'] = EOD_West['Period 2'].str.replace('th','')
#EOD_West['Period 2'] = EOD_West['Period 2'].str.replace('nd','')
#EOD_West['Period 2'] = EOD_West['Period 2'].str.replace('st','')
#EOD_West['Period 2'] = EOD_West['Period 2'].str[-8:]
#EOD_West['Period 2'] = EOD_West['Period 2'].str.replace('-','')
#EOD_West['Period 2'] = EOD_West['Period 2'].str[-5:]
EOD_West.tail()

EOD_West['Region'] = "WEST"
EOD_West['Test'] = "FALSE"
EOD_West['WEST Query'] = ""
EOD_West['Remediation'] = ""
EOD_West['Cleaning'] = ""
EOD_West['WeekNum'] = ""
EOD_West['Remediation'] = ""

# Cleaning attribute logic
EOD_West['Cleaning1'] = EOD_West['Bank ID'].astype(int)
EOD_West['Cleaning1'] = EOD_West['Bank ID'].astype(str)
EOD_West['Cleaning2'] = EOD_West['Period 2'].apply(lambda x: x.strftime('%d%m%y'))
# EOD_West['Cleaning2'] = int(EOD_West['Period 2'].strftime("%Y%m%d%H%M%S"))
EOD_West['Cleaning2'] = EOD_West['Cleaning2'].astype(str)
EOD_West['Cleaning'] = EOD_West['Cleaning1'].str[0:7] + EOD_West['Cleaning2']
EOD_West['Cleaning'] = EOD_West['Cleaning'].str.strip()

# WeekNum Logic
EOD_West['WeekNum'] = EOD_West['Date'].dt.week.astype(str) + EOD_West['Country']
EOD_West['WeekNum'] = EOD_West['WeekNum'].str.strip()

EOD_West = EOD_West[['Bank ID', 'Name', 'Country', 'Role', 'Product Desk', 'Supervisor', 'Supervisor Country', 'Period 2', 'Population', 'Sample', 
                     'Defect Count', 'Comments', 'WEST Query', 'Remediation', 'Cleaning', 'Region', 'WeekNum']]

# Fetching Consolidation sheet info
Conso_EOD_West = pd.read_excel('Consolidated COVID 19 Results - WEST.xlsx', sheet_name = "EOD Email Recap")

# Append to Consolidated Sheet
Conso_EOD_West.rename(columns = {'Defect count':'Defect Count'}, inplace = True)
Conso_EOD_West = Conso_EOD_West[['Bank ID', 'Name', 'Country', 'Role', 'Product Desk', 'Supervisor', 'Supervisor Country', 'Period 2', 'Population', 'Sample', 
                     'Defect Count', 'Comments', 'WEST Query', 'Remediaiton', 'Cleaning', 'Region', 'WeekNum']]
Conso_EOD_West.columns
EOD_West.columns
Conso_EOD_West = Conso_EOD_West.append(EOD_West)
# To Fix miscellaneous duplicates
# Conso_EOD_West = Conso_EOD_West.drop(['Region', 'Test'], axis=1)
Conso_EOD_West = Conso_EOD_West.drop(['Region'], axis=1)
# Conso_EOD_West = Conso_EOD_West.drop(['Region'], axis=1)
Conso_EOD_West['Region'] = "WEST"
Conso_EOD_West['Test'] = "FALSE"

# Cleaning attribute logic after concatenation of current data
EOD_West['Cleaning1'] = EOD_West['Bank ID'].astype(int)
EOD_West['Cleaning1'] = EOD_West['Bank ID'].astype(str)
EOD_West['Cleaning2'] = EOD_West['Period 2'].apply(lambda x: x.strftime('%d%m%y'))
# EOD_West['Cleaning2'] = int(EOD_West['Period 2'].strftime("%Y%m%d%H%M%S"))
EOD_West['Cleaning2'] = EOD_West['Cleaning2'].astype(str)
EOD_West['Cleaning'] = EOD_West['Cleaning1'].str[0:7] + EOD_West['Cleaning2']
EOD_West['Cleaning'] = EOD_West['Cleaning'].str.strip()

# WeekNum Logic after concatenation of current data
EOD_West['WeekNum'] = EOD_West['Period 2'].dt.week.astype(str) + EOD_West['Country']
EOD_West['WeekNum'] = EOD_West['WeekNum'].str.strip()


# Conso_EOD_West = Conso_EOD_West.drop_duplicates()
Conso_EOD_West.columns
# Takes only date part from Timestamp
Conso_EOD_West["Period 2"] = Conso_EOD_West["Period 2"].dt.date
# Conso_EOD_West["Date"] = Conso_EOD_West["Date"].dt.date
Conso_EOD_West["Bank ID"] = Conso_EOD_West["Bank ID"].fillna(0)
Conso_EOD_West["Bank ID"] = Conso_EOD_West["Bank ID"].astype(int)
Conso_EOD_West["Population"] = Conso_EOD_West["Population"].fillna(0)
Conso_EOD_West["Population"] = Conso_EOD_West["Population"].astype(int)
Conso_EOD_West["Sample"] = Conso_EOD_West["Sample"].fillna(0)
Conso_EOD_West["Sample"] = Conso_EOD_West["Sample"].astype(int)
Conso_EOD_West["Defect Count"] = Conso_EOD_West["Defect Count"].fillna(0)
Conso_EOD_West["Defect Count"] = Conso_EOD_West["Defect Count"].astype(int)

# Remove Dups
Conso_EOD_West = Conso_EOD_West.drop_duplicates()

# Fetch only rows with dates
Conso_EOD_West = Conso_EOD_West[~Conso_EOD_West["Period 2"].isnull()]
Conso_EOD_West.tail()

# Rearrange columns
Conso_EOD_West["Period 2"] = pd.to_datetime(Conso_EOD_West["Period 2"], errors='coerce')
Conso_EOD_West.rename(columns = {'Remediation':'Remediaiton'}, inplace = True)
Conso_EOD_West = Conso_EOD_West[['Bank ID', 'Name', 'Country', 'Role', 'Product Desk', 'Supervisor', 'Supervisor Country', 'Period 2', 'Population', 'Sample', 
                     'Defect Count', 'Comments', 'WEST Query', 'Remediaiton', 'Cleaning', 'Region', 'WeekNum']]

######################## Check the defect counts in df2 ########################
df7 = Conso_EOD_West[Conso_EOD_West["Period 2"].dt.strftime('%Y-%m-%d') >= Date_filter_1]
df7
len(df7)
df8 = df7[df7["Defect Count"] > 0]
df8
len(df8)

#########################################################################################################################################
#########################################################################################################################################
                                          ########### ACC Soft Phone WEST ###########
#########################################################################################################################################   
######################################################################################################################################### 
                            
Acc_Soft_Phone_West = pd.read_excel('Covid 19 results West-2022.xlsx', sheet_name = "Accurate softphone line")

print("Length of Acc_Soft_Phone_West:", len(Acc_Soft_Phone_West))

# pandas convert column with integers to date time
print("Type of Acc_Soft_Phone_West[\"Date\"]:", type(Acc_Soft_Phone_West["Date"]))
Acc_Soft_Phone_West["Date"] = pd.to_datetime(Acc_Soft_Phone_West["Date"], errors='coerce')
print (Acc_Soft_Phone_West["Date"].dtypes)
# Filtering greater than the 1st of the needed month
# Acc_Soft_Phone_West = Acc_Soft_Phone_West[Acc_Soft_Phone_West['Date'] > Date_filter]
# Acc_Soft_Phone_West = Acc_Soft_Phone_West[(Acc_Soft_Phone_West['Date'] >= Date_filter_1) & (Acc_Soft_Phone_West['Date'] <= Date_filter_2)]
Acc_Soft_Phone_West = Acc_Soft_Phone_West[Acc_Soft_Phone_West["Date"].dt.strftime('%Y-%m-%d') >= Date_filter_1]
# Drop rows with NaN
Acc_Soft_Phone_West = Acc_Soft_Phone_West[~Acc_Soft_Phone_West["Date"].isnull()]
Acc_Soft_Phone_West.tail()
# Rearrange columns for consolidation
# Remove Existing incomplete or blank columns
Acc_Soft_Phone_West.columns
# Acc_Soft_Phone_West = Acc_Soft_Phone_West.drop(['Region', 'Test'], axis=1)
Acc_Soft_Phone_West.tail()
# Adding needed columns
# Acc_Soft_Phone_West['Period 2'] = ''
# Building Period 2 Column
Acc_Soft_Phone_West["Period 2"] = Acc_Soft_Phone_West['Date']
Acc_Soft_Phone_West["Period 2"] = pd.to_datetime(Acc_Soft_Phone_West["Period 2"], errors='coerce')
Acc_Soft_Phone_West.tail()

Acc_Soft_Phone_West['Region'] = "WEST"
Acc_Soft_Phone_West['Test'] = "FALSE"
Acc_Soft_Phone_West['WEST Query'] = ""
Acc_Soft_Phone_West['Remediation'] = ""
Acc_Soft_Phone_West['Cleaning'] = ""
Acc_Soft_Phone_West['WeekNum'] = ""
Acc_Soft_Phone_West['Remediation'] = ""
Acc_Soft_Phone_West['Supervisor'] = ""
Acc_Soft_Phone_West['Supervisor Country'] = ""

# Cleaning attribute logic
Acc_Soft_Phone_West['Cleaning1'] = Acc_Soft_Phone_West['Bank ID'].astype(int)
Acc_Soft_Phone_West['Cleaning1'] = Acc_Soft_Phone_West['Bank ID'].astype(str)
Acc_Soft_Phone_West['Cleaning2'] = Acc_Soft_Phone_West['Period 2'].apply(lambda x: x.strftime('%d%m%y'))
# Acc_Soft_Phone_West['Cleaning2'] = int(Acc_Soft_Phone_West['Period 2'].strftime("%Y%m%d%H%M%S"))
Acc_Soft_Phone_West['Cleaning2'] = Acc_Soft_Phone_West['Cleaning2'].astype(str)
Acc_Soft_Phone_West['Cleaning'] = Acc_Soft_Phone_West['Cleaning1'].str[0:7] + Acc_Soft_Phone_West['Cleaning2']
Acc_Soft_Phone_West['Cleaning'] = Acc_Soft_Phone_West['Cleaning'].str.strip()

# WeekNum Logic
Acc_Soft_Phone_West['WeekNum'] = Acc_Soft_Phone_West['Date'].dt.week.astype(str) + Acc_Soft_Phone_West['Country']
Acc_Soft_Phone_West['WeekNum'] = Acc_Soft_Phone_West['WeekNum'].str.strip()

Acc_Soft_Phone_West = Acc_Soft_Phone_West[['Bank ID', 'Name', 'Country', 'Role', 'Product Desk', 'Supervisor', 'Supervisor Country', 
                                           'Period 2', 'Population', 'Sample', 'Defect Count', 'Comments', 'WEST Query', 'Remediation', 
                                           'Cleaning', 'Region', 'WeekNum']]


# Fetching Consolidation sheet info
Conso_Acc_Soft_Phone_West = pd.read_excel('Consolidated COVID 19 Results - WEST.xlsx', sheet_name = "Accurate Softphone Line")

# Append to Consolidated Sheet
Conso_Acc_Soft_Phone_West.rename(columns = {'Defect count':'Defect Count'}, inplace = True)
Conso_Acc_Soft_Phone_West.rename(columns = {'West Query':'WEST Query'}, inplace = True)
Conso_Acc_Soft_Phone_West = Conso_Acc_Soft_Phone_West[['Bank ID', 'Name', 'Country', 'Role', 'Product Desk', 'Supervisor', 'Supervisor Country', 
                                           'Period 2', 'Population', 'Sample', 'Defect Count', 'Comments', 'WEST Query', 'Remediation', 
                                           'Cleaning', 'Region', 'WeekNum']]

Conso_Acc_Soft_Phone_West.columns
Acc_Soft_Phone_West.columns
Conso_Acc_Soft_Phone_West = Conso_Acc_Soft_Phone_West.append(Acc_Soft_Phone_West)
# To Fix miscellaneous duplicates
# Conso_Acc_Soft_Phone_West = Conso_Acc_Soft_Phone_West.drop(['Region', 'Test'], axis=1)
Conso_Acc_Soft_Phone_West = Conso_Acc_Soft_Phone_West.drop(['Region'], axis=1)
Conso_Acc_Soft_Phone_West['Region'] = "WEST"
Conso_Acc_Soft_Phone_West['Test'] = "FALSE"
# Conso_Acc_Soft_Phone_West = Conso_Acc_Soft_Phone_West.drop_duplicates()
Conso_Acc_Soft_Phone_West.columns

# Cleaning attribute logic after concatenation of current data
Acc_Soft_Phone_West['Cleaning1'] = Acc_Soft_Phone_West['Bank ID'].astype(int)
Acc_Soft_Phone_West['Cleaning1'] = Acc_Soft_Phone_West['Bank ID'].astype(str)
Acc_Soft_Phone_West['Cleaning2'] = Acc_Soft_Phone_West['Period 2'].apply(lambda x: x.strftime('%d%m%y'))
# Acc_Soft_Phone_West['Cleaning2'] = int(Acc_Soft_Phone_West['Period 2'].strftime("%Y%m%d%H%M%S"))
Acc_Soft_Phone_West['Cleaning2'] = Acc_Soft_Phone_West['Cleaning2'].astype(str)
Acc_Soft_Phone_West['Cleaning'] = Acc_Soft_Phone_West['Cleaning1'].str[0:7] + Acc_Soft_Phone_West['Cleaning2']
Acc_Soft_Phone_West['Cleaning'] = Acc_Soft_Phone_West['Cleaning'].str.strip()

# WeekNum Logic after concatenation of current data
Acc_Soft_Phone_West['WeekNum'] = Acc_Soft_Phone_West['Period 2'].dt.week.astype(str) + Acc_Soft_Phone_West['Country']
Acc_Soft_Phone_West['WeekNum'] = Acc_Soft_Phone_West['WeekNum'].str.strip()

# Dummy as per EOD to be included

EOD_ACC_West = EOD_West
EOD_ACC_West['Remediation (if applicable)'] = ''
EOD_ACC_West['Population'] = 0
EOD_ACC_West['Sample'] = 0
EOD_ACC_West['Defect Count'] = 0
EOD_ACC_West['Comments'] = "NA - no new softphone extension assigned for the covered week"
EOD_ACC_West = EOD_ACC_West[['Bank ID', 'Name', 'Country', 'Role', 'Product Desk', 'Supervisor', 'Supervisor Country', 
                                           'Period 2', 'Population', 'Sample', 'Defect Count', 'Comments', 'WEST Query', 'Remediation', 
                                           'Cleaning', 'Region', 'WeekNum']]

# Dummy concatenation
Conso_Acc_Soft_Phone_West.columns
Acc_Soft_Phone_West.columns
EOD_ACC_West.columns
EOD_ACC_West.rename(columns = {'Line manager':'Supervisor'}, inplace = True)
EOD_ACC_West.rename(columns = {'Line manager Country':'Supervisor Country'}, inplace = True)
Conso_Acc_Soft_Phone_West = Conso_Acc_Soft_Phone_West.append(EOD_ACC_West)
Conso_Acc_Soft_Phone_West.columns

# Takes only date part from Timestamp
Conso_Acc_Soft_Phone_West["Period 2"] = pd.to_datetime(Conso_Acc_Soft_Phone_West["Period 2"], errors='coerce')
Conso_Acc_Soft_Phone_West["Period 2"] = Conso_Acc_Soft_Phone_West["Period 2"].dt.date
# Conso_Acc_Soft_Phone_West["Date"] = Conso_Acc_Soft_Phone_West["Date"].dt.date
Conso_Acc_Soft_Phone_West["Bank ID"] = Conso_Acc_Soft_Phone_West["Bank ID"].fillna(0)
Conso_Acc_Soft_Phone_West["Bank ID"] = Conso_Acc_Soft_Phone_West["Bank ID"].astype(int)
Conso_Acc_Soft_Phone_West["Population"] = Conso_Acc_Soft_Phone_West["Population"].fillna(0)
Conso_Acc_Soft_Phone_West["Population"] = Conso_Acc_Soft_Phone_West["Population"].astype(int)
Conso_Acc_Soft_Phone_West["Sample"] = Conso_Acc_Soft_Phone_West["Sample"].fillna(0)
Conso_Acc_Soft_Phone_West["Sample"] = Conso_Acc_Soft_Phone_West["Sample"].astype(int)
Conso_Acc_Soft_Phone_West["Defect Count"] = Conso_Acc_Soft_Phone_West["Defect Count"].fillna(0)
Conso_Acc_Soft_Phone_West["Defect Count"] = Conso_Acc_Soft_Phone_West["Defect Count"].astype(int)

# Remove Dups
Conso_Acc_Soft_Phone_West = Conso_Acc_Soft_Phone_West.drop_duplicates()

# Fetch only rows with dates
Conso_Acc_Soft_Phone_West = Conso_Acc_Soft_Phone_West[~Conso_Acc_Soft_Phone_West["Period 2"].isnull()]
Conso_Acc_Soft_Phone_West.tail()

# Rearrange columns
Conso_Acc_Soft_Phone_West["Period 2"] = pd.to_datetime(Conso_Acc_Soft_Phone_West["Period 2"], errors='coerce')
Conso_Acc_Soft_Phone_West.rename(columns = {'WEST Query':'West Query'}, inplace = True)
Conso_Acc_Soft_Phone_West = Conso_Acc_Soft_Phone_West[['Bank ID', 'Name', 'Country', 'Role', 'Product Desk', 'Supervisor', 'Supervisor Country', 'Period 2', 'Population', 'Sample', 
                     'Defect Count', 'Comments', 'WEST Query', 'Remediation', 'Cleaning', 'Region', 'WeekNum']]

######################## Check the defect counts in df2 ########################
df9 = Conso_Acc_Soft_Phone_West[Conso_Acc_Soft_Phone_West["Period 2"].dt.strftime('%Y-%m-%d') >= Date_filter_1]
df9
len(df9)
df10 = df9[df9["Defect Count"] > 0]
df10
len(df10)

######################################################################################################################################### 
#########################################################################################################################################
                                          ########### Recorded_Retrievable WEST ###########
#########################################################################################################################################
#########################################################################################################################################
                                
RR_West = pd.read_excel('Covid 19 results West-2022.xlsx', sheet_name = "Recorded_Retrievable")

print("Length of RR_West:", len(RR_West))

# pandas convert column with integers to date time
print("Type of RR_West[\"Date\"]:", type(RR_West["Date"]))
RR_West["Date"] = pd.to_datetime(RR_West["Date"], errors='coerce')
print (RR_West["Date"].dtypes)
# Filtering greater than the 1st of the needed month
# RR_West = RR_West[RR_West['Date'] > Date_filter]
# RR_West = RR_West[(RR_West['Date'] >= Date_filter_1) & (RR_West['Date'] <= Date_filter_2)]
RR_West = RR_West[RR_West["Date"].dt.strftime('%Y-%m-%d') >= Date_filter_1]

# Drop rows with NaN
RR_West = RR_West[~RR_West["Date"].isnull()]
RR_West.tail()
RR_West.columns
# Rearrange columns for consolidation
# Remove Existing incomplete or blank columns
# RR_West = RR_West.drop(['Region', 'Test'], axis=1)
RR_West.tail()
# Building Period 2 Column
RR_West["Period 2"] = RR_West['Date']
RR_West.tail()

RR_West['Region'] = "WEST"
RR_West['Test'] = "FALSE"
RR_West['Supervisor'] = ""
RR_West['WEST Query'] = ""
RR_West['Remediation'] = ""
RR_West['Cleaning'] = ""
RR_West['WeekNum'] = ""

# Cleaning attribute logic
RR_West['Cleaning1'] = RR_West['Bank ID'].astype(int)
RR_West['Cleaning1'] = RR_West['Bank ID'].astype(str)
RR_West['Cleaning2'] = RR_West['Period 2'].apply(lambda x: x.strftime('%d%m%y'))
# RR_West['Cleaning2'] = int(RR_West['Period 2'].strftime("%Y%m%d%H%M%S"))
RR_West['Cleaning2'] = RR_West['Cleaning2'].astype(str)
RR_West['Cleaning'] = RR_West['Cleaning1'].str[0:7] + RR_West['Cleaning2']
RR_West['Cleaning'] = RR_West['Cleaning'].str.strip()

# WeekNum Logic
RR_West['WeekNum'] = RR_West['Date'].dt.week.astype(str) + RR_West['Country']
RR_West['WeekNum'] = RR_West['WeekNum'].str.strip()

RR_West = RR_West[['Bank ID', 'Name', 'Country', 'Role', 'Product Desk', 'Supervisor', 'Supervisor Country', 'Period 2', 'Population', 'Sample', 
                     'Defect Count', 'Comments', 'WEST Query', 'Remediation', 'Cleaning', 'Region', 'WeekNum']]

# Fetching Consolidation sheet info
Conso_RR_West = pd.read_excel('Consolidated COVID 19 Results - WEST.xlsx', sheet_name = "Recorded_Retrievable")

# Append to Consolidated Sheet
# Conso_RR_West.rename(columns = {'Defect count':'Defect Count'}, inplace = True)
Conso_RR_West = Conso_RR_West[['Bank ID', 'Name', 'Country', 'Role', 'Product Desk', 'Supervisor', 'Supervisor Country', 'Period 2', 'Population', 'Sample', 
                     'Defect Count', 'Comments', 'WEST Query', 'Remediation', 'Cleaning', 'Region', 'WeekNum']]
Conso_RR_West.columns
RR_West.columns
Conso_RR_West = Conso_RR_West.append(RR_West)
# To Fix miscellaneous duplicates
# Conso_RR_West = Conso_RR_West.drop(['Region', 'Test'], axis=1)
Conso_RR_West = Conso_RR_West.drop(['Region'], axis=1)
Conso_RR_West['Region'] = "WEST"
Conso_RR_West['Test'] = "FALSE"

# Cleaning attribute logic
RR_West['Cleaning1'] = RR_West['Bank ID'].astype(int)
RR_West['Cleaning1'] = RR_West['Bank ID'].astype(str)
RR_West['Cleaning2'] = RR_West['Period 2'].apply(lambda x: x.strftime('%d%m%y'))
# RR_West['Cleaning2'] = int(RR_West['Period 2'].strftime("%Y%m%d%H%M%S"))
RR_West['Cleaning2'] = RR_West['Cleaning2'].astype(str)
RR_West['Cleaning'] = RR_West['Cleaning1'].str[0:7] + RR_West['Cleaning2']
RR_West['Cleaning'] = RR_West['Cleaning'].str.strip()

# WeekNum Logic
RR_West['WeekNum'] = RR_West['Period 2'].dt.week.astype(str) + RR_West['Country']
RR_West['WeekNum'] = RR_West['WeekNum'].str.strip()

# Conso_RR_West = Conso_RR_West.drop_duplicates()
Conso_RR_West.columns
Conso_RR_West["Bank ID"] = Conso_RR_West["Bank ID"].fillna(0)
Conso_RR_West["Bank ID"] = Conso_RR_West["Bank ID"].astype(int)
Conso_RR_West["Population"] = Conso_RR_West["Population"].fillna(0)
Conso_RR_West["Population"] = Conso_RR_West["Population"].astype(int)
Conso_RR_West["Sample"] = Conso_RR_West["Sample"].fillna(0)
Conso_RR_West["Sample"] = Conso_RR_West["Sample"].astype(int)
Conso_RR_West["Defect Count"] = Conso_RR_West["Defect Count"].fillna(0)
Conso_RR_West["Defect Count"] = Conso_RR_West["Defect Count"].astype(int)

# Remove Dups
Conso_RR_West = Conso_RR_West.drop_duplicates()

# Fetch only rows with dates
Conso_RR_West = Conso_RR_West[~Conso_RR_West["Period 2"].isnull()]
Conso_RR_West.tail()

# Rearrange columns
Conso_RR_West = Conso_RR_West[['Bank ID', 'Name', 'Country', 'Role', 'Product Desk', 'Supervisor', 'Supervisor Country', 'Period 2', 'Population', 'Sample', 
                     'Defect Count', 'Comments', 'WEST Query', 'Remediation', 'Cleaning', 'Region', 'WeekNum']]


######################## Check the defect counts in df2 ########################
df11 = Conso_RR_West[Conso_RR_West["Period 2"].dt.strftime('%Y-%m-%d') >= Date_filter_1]
df11
len(df11)
df12 = df11[df11["Defect Count"] > 0]
df12
len(df12)


################ WRITE TO EXCEL ##################

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Consolidated COVID 19 Results - WEST NEW.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.
# Fetching Consolidation sheet info
Conso_Logged_in_West = pd.read_excel('Consolidated COVID 19 Results - WEST.xlsx', sheet_name = "Logged in")
Conso_EOD_West.to_excel(writer, sheet_name='EOD Email Recap', index=False)
Conso_Acc_Soft_Phone_West.to_excel(writer, sheet_name='Accurate Softphone Line', index=False)
Conso_RR_West.to_excel(writer, sheet_name='Recorded_Retrievable', index=False)
Conso_Logged_in_West.to_excel(writer, sheet_name='Logged in', index=False)


# Close the Pandas Excel writer and output the Excel file.
writer.save()


print("Change Date_filter to needed month start date")
