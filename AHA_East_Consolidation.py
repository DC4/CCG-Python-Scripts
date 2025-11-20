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
                                                    ########### EOD EAST ###########
#########################################################################################################################################
#########################################################################################################################################
                                
EOD_East = pd.read_excel('Covid 19 East Result-2022.xlsx', sheet_name = "Eod Recap")

print("Length of EOD_East:", len(EOD_East))
# pandas convert column with integers to date time
print("Type of EOD_East[\"Date\"]:", type(EOD_East["Date"]))
EOD_East["Date"] = pd.to_datetime(EOD_East["Date"], errors='coerce')
# EOD_East = EOD_East[EOD_East["Date"].dt.strftime('%Y-%m-%d') > "2022-06-01"]
EOD_East = EOD_East[EOD_East["Date"].dt.strftime('%Y-%m-%d') >= Date_filter_1]
# Drop rows with NaN
EOD_East = EOD_East[~EOD_East["Date"].isnull()]
EOD_East.tail()
# Rearrange columns for consolidation
# Remove Existing incomplete or blank columns
EOD_East = EOD_East.drop(['Region', 'Test'], axis=1)
EOD_East.tail()
# Adding needed columns
# EOD_East['Period 2'] = ''
# Building Period 2 Column
from pandas.tseries.offsets import *
EOD_East["Period 2"] = EOD_East['Date'] + Week(weekday=4)
#EOD_East['Period'] = EOD_East['Period'].str.strip()
#EOD_East['Period 2'] = EOD_East['Period'].str[-9:]
#print (EOD_East['Period 2'].dtypes)
#EOD_East['Period 2'] = EOD_East['Period 2'].str.replace('th','')
#EOD_East['Period 2'] = EOD_East['Period 2'].str.replace('nd','')
#EOD_East['Period 2'] = EOD_East['Period 2'].str.replace('st','')
#EOD_East['Period 2'] = EOD_East['Period 2'].str[-8:]
#EOD_East['Period 2'] = EOD_East['Period 2'].str.replace('-','')
#EOD_East['Period 2'] = EOD_East['Period 2'].str[-5:]
EOD_East.tail()

EOD_East.rename(columns = {'Population ':'Population'}, inplace = True)
EOD_East['Region'] = "EAST"
EOD_East['Test'] = "FALSE"
EOD_East = EOD_East[['Period', 'Period 2', 'Bank ID', 'Name', 'Country', 'Role', 'Product Desk', 
'Line manager ID', 'Line manager', 'Line manager Country', 'Day', 'Date', 
'Population', 'Sample', 'Defect Count', 'EOD email summary', 'Region', 'Test']]

# Fetching Consolidation sheet info
Conso_EOD_East = pd.read_excel('Consolidated COVID 19 Results - EAST.xlsx', sheet_name = "EOD Email Recap")

# Append to Consolidated Sheet
Conso_EOD_East.rename(columns = {'Defect count':'Defect Count'}, inplace = True)
Conso_EOD_East = Conso_EOD_East[['Period', 'Period 2', 'Bank ID', 'Name', 'Country', 'Role', 'Product Desk', 
'Line manager ID', 'Line manager', 'Line manager Country', 'Day', 'Date', 
'Population', 'Sample', 'Defect Count', 'EOD email summary', 'Region', 'Test']]
Conso_EOD_East.columns
EOD_East.columns
Conso_EOD_East = Conso_EOD_East.append(EOD_East)
# To Fix miscellaneous duplicates
Conso_EOD_East = Conso_EOD_East.drop(['Region', 'Test'], axis=1)
Conso_EOD_East['Region'] = "EAST"
Conso_EOD_East['Test'] = "FALSE"
# Conso_EOD_East = Conso_EOD_East.drop_duplicates()
Conso_EOD_East.columns
# Takes only date part from Timestamp
Conso_EOD_East["Period 2"] = Conso_EOD_East["Period 2"].dt.date
Conso_EOD_East["Date"] = Conso_EOD_East["Date"].dt.date
Conso_EOD_East["Bank ID"] = Conso_EOD_East["Bank ID"].astype(int)
Conso_EOD_East["Population"] = Conso_EOD_East["Population"].astype(int)
Conso_EOD_East["Sample"] = Conso_EOD_East["Sample"].astype(int)
Conso_EOD_East["Defect Count"] = Conso_EOD_East["Defect Count"].fillna(0)
Conso_EOD_East["Defect Count"] = Conso_EOD_East["Defect Count"].astype(int)

# Fetch only rows with dates
Conso_EOD_East = Conso_EOD_East[~Conso_EOD_East["Date"].isnull()]
Conso_EOD_East.tail()

# Rearrange columns
Conso_EOD_East = Conso_EOD_East[['Period', 'Period 2', 'Bank ID', 'Name', 'Country', 'Role', 'Product Desk', 
'Line manager ID', 'Line manager', 'Line manager Country', 'Day', 'Date', 
'Population', 'Sample', 'Defect Count', 'EOD email summary', 'Region', 'Test']]

# Remove Dups
Conso_EOD_East = Conso_EOD_East.drop_duplicates()

######################## Check the defect counts in df2 ########################
Conso_EOD_East["Date"] = pd.to_datetime(Conso_EOD_East["Date"], errors='coerce')
df1 = Conso_EOD_East[Conso_EOD_East["Date"].dt.strftime('%Y-%m-%d') >= Date_filter_1]
df1
len(df1)
df2 = df1[df1["Defect Count"] > 0]
df2
len(df2)

#########################################################################################################################################
#########################################################################################################################################
                                          ########### ACC Soft Phone EAST ###########
#########################################################################################################################################   
#########################################################################################################################################                    

Acc_Soft_Phone_East = pd.read_excel('Covid 19 East Result-2022.xlsx', sheet_name = "Accurate Softphone line")

print("Length of Acc_Soft_Phone_East:", len(Acc_Soft_Phone_East))

# pandas convert column with integers to date time
print("Type of Acc_Soft_Phone_East[\"Date\"]:", type(Acc_Soft_Phone_East["Date"]))
Acc_Soft_Phone_East["Date"] = pd.to_datetime(Acc_Soft_Phone_East["Date"], errors='coerce')
print (Acc_Soft_Phone_East["Date"].dtypes)
# Filtering greater than the 1st of the needed month
# Acc_Soft_Phone_East = Acc_Soft_Phone_East[Acc_Soft_Phone_East['Date'] > Date_filter]
# Acc_Soft_Phone_East = Acc_Soft_Phone_East[(Acc_Soft_Phone_East['Date'] >= Date_filter_1) & (Acc_Soft_Phone_East['Date'] <= Date_filter_2)]
Acc_Soft_Phone_East = Acc_Soft_Phone_East[Acc_Soft_Phone_East["Date"].dt.strftime('%Y-%m-%d') >= Date_filter_1]
# Drop rows with NaN
Acc_Soft_Phone_East = Acc_Soft_Phone_East[~Acc_Soft_Phone_East["Date"].isnull()]
Acc_Soft_Phone_East.tail()
# Rearrange columns for consolidation
# Remove Existing incomplete or blank columns
Acc_Soft_Phone_East = Acc_Soft_Phone_East.drop(['Region', 'Test'], axis=1)
Acc_Soft_Phone_East.tail()
# Adding needed columns
# EOD_East['Period 2'] = ''
# Building Period 2 Column
# from pandas.tseries.offsets import *
Acc_Soft_Phone_East["Period 2"] = Acc_Soft_Phone_East['Date'] + Week(weekday=4)
Acc_Soft_Phone_East.tail()

Acc_Soft_Phone_East['Region'] = "EAST"
Acc_Soft_Phone_East['Test'] = "FALSE"
# Adding needed columns
Acc_Soft_Phone_East['Supervisor'] = ""
Acc_Soft_Phone_East['Supervisor Country'] = ""
Acc_Soft_Phone_East['Defect Count'] = ""
Acc_Soft_Phone_East['Comments'] = ""
Acc_Soft_Phone_East['Remediation (if applicable)'] = ""
Acc_Soft_Phone_East = Acc_Soft_Phone_East[['Period', 'Period 2', 'Bank ID', 'Name', 'Country', 'Role', 
'Product Desk', 'Supervisor', 'Supervisor Country', 'Date', 'Population ', 'Sample', 'Defect Count', 'Comments', 
'Remediation (if applicable)', 'Region', 'Test']]

# Fetching Consolidation sheet info
Conso_Acc_Soft_Phone_East = pd.read_excel('Consolidated COVID 19 Results - EAST.xlsx', sheet_name = "Accurate Softphone Line")

# Append to Consolidated Sheet
Conso_Acc_Soft_Phone_East.rename(columns = {'Defect count':'Defect Count'}, inplace = True)
Conso_Acc_Soft_Phone_East = Conso_Acc_Soft_Phone_East[['Period', 'Period 2', 'Bank ID', 'Name', 'Country', 'Role', 
'Product Desk', 'Supervisor', 'Supervisor Country', 'Date', 'Population', 'Sample', 'Defect Count', 'Comments', 
'Remediation (if applicable)', 'Region', 'Test']]
Conso_Acc_Soft_Phone_East.columns
Acc_Soft_Phone_East.columns
Conso_Acc_Soft_Phone_East = Conso_Acc_Soft_Phone_East.append(Acc_Soft_Phone_East)
# To Fix miscellaneous duplicates
Conso_Acc_Soft_Phone_East = Conso_Acc_Soft_Phone_East.drop(['Region', 'Test'], axis=1)
Conso_Acc_Soft_Phone_East['Region'] = "EAST"
Conso_Acc_Soft_Phone_East['Test'] = "FALSE"

# Dummy as per EOD to be included
EOD_ACC_East = EOD_East
EOD_ACC_East['Remediation (if applicable)'] = ''
EOD_ACC_East['Population'] = 0
EOD_ACC_East['Sample'] = 0
EOD_ACC_East['Defect Count'] = 0
EOD_ACC_East['Comments'] = "NA - no new softphone extension assigned for the covered week"
EOD_ACC_East = EOD_ACC_East[['Period', 'Period 2', 'Bank ID', 'Name', 'Country', 'Role', 'Product Desk',
'Line manager', 'Line manager Country', 'Date', 'Population', 'Sample', 'Defect Count', 'Comments', 
'Remediation (if applicable)', 'Region', 'Test']]

# Dummy concatenation
Conso_Acc_Soft_Phone_East.columns
Acc_Soft_Phone_East.columns
EOD_ACC_East.columns
EOD_ACC_East.rename(columns = {'Line manager':'Supervisor'}, inplace = True)
EOD_ACC_East.rename(columns = {'Line manager Country':'Supervisor Country'}, inplace = True)
Conso_Acc_Soft_Phone_East = Conso_Acc_Soft_Phone_East.append(EOD_ACC_East)
Conso_Acc_Soft_Phone_East.columns

# Conso_EOD_East = Conso_EOD_East.drop_duplicates()
Conso_Acc_Soft_Phone_East.columns
Conso_Acc_Soft_Phone_East["Bank ID"] = Conso_Acc_Soft_Phone_East["Bank ID"].fillna(0)
Conso_Acc_Soft_Phone_East["Bank ID"] = Conso_Acc_Soft_Phone_East["Bank ID"].astype(int)
Conso_Acc_Soft_Phone_East["Population"] = Conso_Acc_Soft_Phone_East["Population"].fillna(0)
Conso_Acc_Soft_Phone_East["Population"] = Conso_Acc_Soft_Phone_East["Population"].astype(int)
Conso_Acc_Soft_Phone_East["Sample"] = Conso_Acc_Soft_Phone_East["Sample"].fillna(0)
Conso_Acc_Soft_Phone_East["Sample"] = Conso_Acc_Soft_Phone_East["Sample"].astype(int)
#import re
#dfs['patterns'] = Conso_Acc_Soft_Phone_East["Defect Count"].str.findall(r'\d+\/\d+\/\d+|\d+')
#Conso_Acc_Soft_Phone_East["Defect Count"] = Conso_Acc_Soft_Phone_East["Defect Count"].str.findall(r'\d+\.\d+')
Conso_Acc_Soft_Phone_East["Defect Count"] = Conso_Acc_Soft_Phone_East["Defect Count"].fillna(0)
#Conso_Acc_Soft_Phone_East = Conso_Acc_Soft_Phone_East.dropna()
Conso_Acc_Soft_Phone_East["Defect Count"] = pd.to_numeric(Conso_Acc_Soft_Phone_East["Defect Count"], errors='coerce')
Conso_Acc_Soft_Phone_East["Defect Count"] = Conso_Acc_Soft_Phone_East["Defect Count"].astype(float)
Conso_Acc_Soft_Phone_East["Defect Count"] = Conso_Acc_Soft_Phone_East["Defect Count"].fillna(0)
Conso_Acc_Soft_Phone_East["Defect Count"] = Conso_Acc_Soft_Phone_East["Defect Count"].astype(int)
Conso_Acc_Soft_Phone_East["Defect Count"].unique()

# Fetch only rows with dates
Conso_Acc_Soft_Phone_East = Conso_Acc_Soft_Phone_East[~Conso_Acc_Soft_Phone_East["Date"].isnull()]
Conso_Acc_Soft_Phone_East.tail()

# Rearrange columns
Conso_Acc_Soft_Phone_East = Conso_Acc_Soft_Phone_East[['Period', 'Period 2', 'Bank ID', 'Name', 'Country', 'Role', 
'Product Desk', 'Supervisor', 'Supervisor Country', 'Date', 'Population ', 'Sample', 'Defect Count', 'Comments', 
'Remediation (if applicable)', 'Region', 'Test']]

# Remove Dups
Conso_Acc_Soft_Phone_East = Conso_Acc_Soft_Phone_East.drop_duplicates()

######################## Check the defect counts in df2 ########################
Conso_Acc_Soft_Phone_East["Period 2"] = pd.to_datetime(Conso_Acc_Soft_Phone_East["Period 2"], errors = 'coerce')
Conso_Acc_Soft_Phone_East["Period 2"] = Conso_Acc_Soft_Phone_East["Period 2"].dt.date
Conso_Acc_Soft_Phone_East["Date"] = Conso_Acc_Soft_Phone_East["Date"].fillna(pd.to_datetime('1989-01-01'))
Conso_Acc_Soft_Phone_East["Date"] = pd.to_datetime(Conso_Acc_Soft_Phone_East["Date"], errors = 'coerce')
Conso_Acc_Soft_Phone_East["Date"] = Conso_Acc_Soft_Phone_East["Date"].dt.date
Conso_Acc_Soft_Phone_East["Date"] =  pd.to_datetime(Conso_Acc_Soft_Phone_East["Date"], format='%Y-%m-%d %H:%M:%S')
df3 = Conso_Acc_Soft_Phone_East[Conso_Acc_Soft_Phone_East["Date"].dt.strftime('%Y-%m-%d') >= Date_filter_1]
df3
len(df3)
df4 = df3[df3["Defect Count"] > 0]
df3["Defect Count"].unique()
df4
len(df4)


######################################################################################################################################### 
#########################################################################################################################################
                                          ########### Recorded_Retrievable EAST ###########
#########################################################################################################################################
#########################################################################################################################################

                                
RR_East = pd.read_excel('Covid 19 East Result-2022.xlsx', sheet_name = "Recorded_Retrievable")

print("Length of RR_East:", len(RR_East))

# pandas convert column with integers to date time
print("Type of RR_East[\"Date\"]:", type(RR_East["Date"]))
RR_East["Date"] = pd.to_datetime(RR_East["Date"], errors='coerce')
print (RR_East["Date"].dtypes)
# Filtering greater than the 1st of the needed month
# RR_East = RR_East[RR_East['Date'] > Date_filter]
# RR_East = RR_East[(RR_East['Date'] >= Date_filter_1) & (RR_East['Date'] <= Date_filter_2)]
RR_East = RR_East[RR_East["Date"].dt.strftime('%Y-%m-%d') >= Date_filter_1]

# Drop rows with NaN
RR_East = RR_East[~RR_East["Date"].isnull()]
RR_East.tail()
# Rearrange columns for consolidation
# Remove Existing incomplete or blank columns
RR_East = RR_East.drop(['Region', 'Test'], axis=1)
RR_East.tail()
# Building Period 2 Column
# from pandas.tseries.offsets import *
RR_East["Period 2"] = RR_East['Date'] + Week(weekday=4)
RR_East.tail()

RR_East['Region'] = "EAST"
RR_East['Test'] = "FALSE"
RR_East = RR_East[['Period', 'Period 2', 'Bank ID', 'Name', 'Country', 'Role', 'Product Desk', 
'Line manager ID', 'Line manager', 'Line manager Country', 'Day', 'Date', 
'Population', 'Sample', 'Defect', 'Remarks', 'Region', 'Test']]

# Fetching Consolidation sheet info
Conso_RR_East = pd.read_excel('Consolidated COVID 19 Results - EAST.xlsx', sheet_name = "Recorded_Retrievable")

# Append to Consolidated Sheet
# Conso_RR_East.rename(columns = {'Defect count':'Defect Count'}, inplace = True)
Conso_RR_East = Conso_RR_East[['Period', 'Period 2', 'Bank ID', 'Name', 'Country', 'Role', 
'Product Desk', 'Supervisor', 'Supervisor Country', 'Date', 'Population', 'Sample', 'Defect Count', 'Comments', 
'Remediation (if applicable)', 'Region', 'Test']]
RR_East.rename(columns = {'Line manager':'Supervisor'}, inplace = True)
RR_East.rename(columns = {'Line manager Country':'Supervisor Country'}, inplace = True)
RR_East.rename(columns = {'Defect':'Defect Count'}, inplace = True)
RR_East.rename(columns = {'Remarks':'Comments'}, inplace = True)
RR_East.rename(columns = {'Remarks':'Comments'}, inplace = True)
RR_East['Remediation (if applicable)'] = ''
RR_East = RR_East.drop(['Line manager ID', 'Day'], axis=1)
RR_East = RR_East[['Period', 'Period 2', 'Bank ID', 'Name', 'Country', 'Role', 
'Product Desk', 'Supervisor', 'Supervisor Country', 'Date', 'Population', 'Sample', 'Defect Count', 'Comments', 
'Remediation (if applicable)', 'Region', 'Test']]
Conso_RR_East.columns
RR_East.columns
Conso_RR_East = Conso_RR_East.append(RR_East)
# To Fix miscellaneous duplicates
Conso_RR_East = Conso_RR_East.drop(['Region', 'Test'], axis=1)
Conso_RR_East['Region'] = "EAST"
Conso_RR_East['Test'] = "FALSE"
# Conso_EOD_East = Conso_EOD_East.drop_duplicates()
Conso_RR_East.columns
Conso_RR_East["Bank ID"] = Conso_RR_East["Bank ID"].fillna(0)
Conso_RR_East["Bank ID"] = Conso_RR_East["Bank ID"].astype(int)
Conso_RR_East["Population"] = Conso_RR_East["Population"].fillna(0)
Conso_RR_East["Population"] = Conso_RR_East["Population"].astype(int)
Conso_RR_East["Sample"] = Conso_RR_East["Sample"].fillna(0)
Conso_RR_East["Sample"] = Conso_RR_East["Sample"].astype(int)
Conso_RR_East["Defect Count"] = Conso_RR_East["Defect Count"].fillna(0)
Conso_RR_East["Defect Count"] = Conso_RR_East["Defect Count"].astype(int)

# Fetch only rows with dates
Conso_RR_East = Conso_RR_East[~Conso_RR_East["Date"].isnull()]
Conso_RR_East.tail()

# Rearrange columns
Conso_RR_East = Conso_RR_East[['Period', 'Period 2', 'Bank ID', 'Name', 'Country', 'Role', 
'Product Desk', 'Supervisor', 'Supervisor Country', 'Date', 'Population', 'Sample', 'Defect Count', 'Comments', 
'Remediation (if applicable)', 'Region', 'Test']]

# Remove Dups
Conso_RR_East = Conso_RR_East.drop_duplicates()

######################## Check the defect counts in df2 ########################
df5 = Conso_RR_East[Conso_RR_East["Date"].dt.strftime('%Y-%m-%d') >= Date_filter_1]
df5
len(df5)
df5["Defect Count"].unique()
df6 = df5[df5["Defect Count"] > 0]
df6
len(df6)


################ WRITE TO EXCEL ##################

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('Consolidated COVID 19 Results - EAST NEW.xlsx', engine='xlsxwriter')

# Write each dataframe to a different worksheet.
# Fetching Consolidation sheet info
Conso_Logged_in_East = pd.read_excel('Consolidated COVID 19 Results - EAST.xlsx', sheet_name = "Logged in")
Conso_EOD_East.to_excel(writer, sheet_name='EOD Email Recap', index=False)
Conso_Acc_Soft_Phone_East.to_excel(writer, sheet_name='Accurate Softphone Line', index=False)
Conso_RR_East.to_excel(writer, sheet_name='Recorded_Retrievable', index=False)
Conso_Logged_in_East.to_excel(writer, sheet_name='Logged in', index=False)


# Close the Pandas Excel writer and output the Excel file.
writer.save()


print("Change Date_filter to needed month start date")
