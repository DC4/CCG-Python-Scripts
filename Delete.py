# -*- coding: utf-8 -*-
"""
Created on Tue Aug 13 17:42:53 2024

@author: 1510806
"""
# -*- coding: utf-8 -*-
"""
Created on Tue Mar 19 16:35:59 2024

@author: 1510806
"""
# -*- coding: utf-8 -*-
"""
Created on Fri Feb 23 11:35:58 2024

@author: 1510806
"""
# import pandas
import pandas as pd
import numpy as np
import os
# Import time module
import time

 
# record start time
start_time = time.time()

os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop")
#Issues = pd.read_csv('1022_Issues and Actions Report (1).csv', encoding='cp1252',  low_memory=False, error_bad_lines=False)

print("Define the Start_Date and End_Date Dates for filtering each dataset")
print("All 4 Excels - 2 Riz & 2 Sahira - Remove top few lines in the dataset and retain data only about columns")
print("All 4 Excels - 2 Riz & 2 Sahira - Convert the Sensitivity of the Excel from Confidential to Internal")
print("Change the month in - FMO_Non_Hubbed_Raw - mentioned like 'Feb 2024'")
print("For India ")

# yyyy-mm-dd
Start_Date = pd.to_datetime('2024-07-01', format='%Y-%m-%d')
End_Date = pd.to_datetime('2024-08-08', format='%Y-%m-%d')

##############################################################################################################################
#################################################    INDIA - HUB & NON_HUB    ################################################
##############################################################################################################################

###############################################################
####################### Riz KCI IN Data #######################
###############################################################

# Reading the Raw Riz KCI File with all data
Riz_KCI_IN = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Base Raw Data//Riz Data//Metric Responses KCI -Completed.xlsx')
# Reading specific data alone from the sheet
Riz_KCI_IN_Raw = pd.read_excel(Riz_KCI_IN, 'Metric Responses KCI -Completed')
print("Shape of Riz_KCI_IN_Raw: ", Riz_KCI_IN_Raw.shape)

# Filtering based on Due Date
Riz_KCI_IN_Raw['Due Date'] = pd.to_datetime(Riz_KCI_IN_Raw['Due Date'])
Riz_KCI_IN_Raw = Riz_KCI_IN_Raw[(Riz_KCI_IN_Raw['Due Date'].ge(Start_Date)) & (Riz_KCI_IN_Raw['Due Date'].le(End_Date))]
print("Shape of Riz_KCI_IN_Raw after Due Date filter: ", Riz_KCI_IN_Raw.shape)

# Filtering for Org Nodes
# GBS INDIA KCI Non - Hub Org nodes (Riz):
Org1 = Riz_KCI_IN_Raw[Riz_KCI_IN_Raw['Data Owner Organization'].str.contains("L3-71000325 OPS FM Ops ǀ NA ǀ GBS India", na=False, case=False)]
Org2 = Riz_KCI_IN_Raw[Riz_KCI_IN_Raw['Data Owner Organization'].str.contains("L3-71000325 OPS FM Ops ǀ NA ǀ GBS Malaysia", na=False, case=False)]
Org1_1 = Riz_KCI_IN_Raw[Riz_KCI_IN_Raw['Data Owner Organization'].str.contains("L1-90000025 CCIB Financial Markets ǀ NA ǀ AUSTRALIA", na=False, case=False)]
Org1_2 = Riz_KCI_IN_Raw[Riz_KCI_IN_Raw['Data Owner Organization'].str.contains("L1-90000025 CCIB Financial Markets ǀ NA ǀ GROUP", na=False, case=False)]
Org1_3 = Riz_KCI_IN_Raw[Riz_KCI_IN_Raw['Data Owner Organization'].str.contains("L1-90000025 CCIB Financial Markets ǀ NA ǀ GBS India", na=False, case=False)]
# Org3 = Riz_KCI_IN_Raw[Riz_KCI_IN_Raw['Data Owner Organization'].str.contains("L3-70000101 CCIB FM Head Financial Mkt ǀ NA ǀ GBS India", na=False, case=False)]
KCI_Non_Hub_IN = Org1.append([Org2,Org1_1,Org1_2,Org1_3])
print("Shape of KCI_Non_Hub_IN with Org node filters: ", KCI_Non_Hub_IN.shape)

# Reading the Raw FMO Non-Hubbed data
print("Choose the Month sheet here manually to the needed month like - Feb 2024")
FMO_Non_Hubbed = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Base Raw Data//Riz Data//CCG Control Monitoring BAU Tracker FMO Non-Hubbed_2024.xlsx')
# Reading specific data alone from the sheet
FMO_Non_Hubbed_Raw = pd.read_excel(FMO_Non_Hubbed, 'July 2024')
print("Shape of FMO_Non_Hubbed_Raw: ", FMO_Non_Hubbed_Raw.shape)
# Filter Indicator type
KCI_FMO_Non_Hubbed_Raw = FMO_Non_Hubbed_Raw[FMO_Non_Hubbed_Raw['Indicator Type'].str.contains("KCI", na=False, case=False)]
print("Shape of KCI_FMO_Non_Hubbed_Raw: ", KCI_FMO_Non_Hubbed_Raw.shape)

# vlookup with the Test Ids from KCI_FMO_Non_Hubbed_Raw
KCI_Non_Hub_IN = pd.merge(KCI_Non_Hub_IN,  KCI_FMO_Non_Hubbed_Raw['Test ID'],  left_on ='Metric Response ID', right_on = 'Test ID', how ='inner')
KCI_Non_Hub_IN['Data From'] = 'KCI_Non_Hub_IN'
print("Shape of KCI_Non_Hub_IN after vlookup: ", KCI_Non_Hub_IN.shape)
# KCI_Non_Hub_IN.columns

###############################################################
####################### Riz CST IN Data #######################
###############################################################

# Reading the Raw Riz CST File with all data
Riz_CST_IN = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Base Raw Data//Riz Data//Control Sample Test Report .xlsx')
# Reading specific data alone from the sheet
Riz_CST_IN_Raw = pd.read_excel(Riz_CST_IN, 'Control Sample Test Report ')
print("Shape of Riz_CST_IN_Raw: ", Riz_CST_IN_Raw.shape)

# Filtering based on Due Date
Riz_CST_IN_Raw['Due Date'] = pd.to_datetime(Riz_CST_IN_Raw['Due Date'])
Riz_CST_IN_Raw = Riz_CST_IN_Raw[(Riz_CST_IN_Raw['Due Date'].ge(Start_Date)) & (Riz_CST_IN_Raw['Due Date'].le(End_Date))]
print("Shape of Riz_CST_IN_Raw after Due Date filter: ", Riz_CST_IN_Raw.shape)

# Filtering for Org Nodes
# GBS INDIA CST Non - Hub Org nodes (Riz):
Org4 = Riz_CST_IN_Raw[Riz_CST_IN_Raw['Tester Organization'].str.contains("L3-71000325 OPS FM Ops ǀ NA ǀ GBS India", na=False, case=False)]
Org5 = Riz_CST_IN_Raw[Riz_CST_IN_Raw['Tester Organization'].str.contains("L3-71000325 OPS FM Ops ǀ NA ǀ GBS Malaysia", na=False, case=False)]
Org6 = Riz_CST_IN_Raw[Riz_CST_IN_Raw['Tester Organization'].str.contains("L1-90000025 CCIB Financial Markets ǀ NA ǀ AUSTRALIA", na=False, case=False)]
Org7 = Riz_CST_IN_Raw[Riz_CST_IN_Raw['Tester Organization'].str.contains("L1-90000025 CCIB Financial Markets ǀ NA ǀ GROUP", na=False, case=False)]
Org7_1 = Riz_CST_IN_Raw[Riz_CST_IN_Raw['Tester Organization'].str.contains("L1-90000025 CCIB Financial Markets ǀ NA ǀ GBS India", na=False, case=False)]
CST_Non_Hub_IN = Org4.append([Org5, Org6, Org7,Org7_1])
print("Shape of CST_Non_Hub_IN after Tester Organization filter: ", CST_Non_Hub_IN.shape)

# Reading the Raw FMO Non-Hubbed data
print("Choose the Month sheet here manually to the needed month like - Feb 2024")
FMO_Non_Hubbed = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Base Raw Data//Riz Data//CCG Control Monitoring BAU Tracker FMO Non-Hubbed_2024.xlsx')
# Reading specific data alone from the sheet
FMO_Non_Hubbed_Raw = pd.read_excel(FMO_Non_Hubbed, 'July 2024')
print("Shape of FMO_Non_Hubbed_Raw: ", FMO_Non_Hubbed_Raw.shape)
# Filter Indicator type
CST_FMO_Non_Hubbed_Raw = FMO_Non_Hubbed_Raw[FMO_Non_Hubbed_Raw['Indicator Type'].str.contains("CST", na=False, case=False)]
print("Shape of CST_FMO_Non_Hubbed_Raw: ", CST_FMO_Non_Hubbed_Raw.shape)

# vlookup with the Test Ids from CST_FMO_Non_Hubbed_Raw
CST_Non_Hub_IN = pd.merge(CST_Non_Hub_IN,  CST_FMO_Non_Hubbed_Raw['Test ID'],  left_on ='Execution ID', right_on = 'Test ID', how ='inner')
CST_Non_Hub_IN['Data From'] = 'CST_Non_Hub_IN'
print("Shape of CST_Non_Hub_IN after vlookup: ", CST_Non_Hub_IN.shape)
# CST_Non_Hub_IN.columns

###############################################################
####################### Sahira KCI IN Data ####################
###############################################################

# Reading the Raw Riz KCI File with all data
Sahira_KCI_IN = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Base Raw Data//Sahira Data//Metric Response - Jul 24.xlsx')
# Reading specific data alone from the sheet
Sahira_KCI_IN_Raw = pd.read_excel(Sahira_KCI_IN, 'Metric Response - Jul 24')
print("Shape of Sahira_KCI_IN_Raw: ", Sahira_KCI_IN_Raw.shape)

# Filtering based on Due Date
Sahira_KCI_IN_Raw['Responded Date'] = pd.to_datetime(Sahira_KCI_IN_Raw['Responded Date'])
Sahira_KCI_IN_Raw = Sahira_KCI_IN_Raw[(Sahira_KCI_IN_Raw['Responded Date'].ge(Start_Date)) & (Sahira_KCI_IN_Raw['Responded Date'].le(End_Date))]
print("Shape of Sahira_KCI_IN_Raw after Responded Date filter: ", Sahira_KCI_IN_Raw.shape)

# Filtering for Org Nodes
# GBS INDIA KCI Hub Org nodes (Sahira’s):
Org8 = Sahira_KCI_IN_Raw[Sahira_KCI_IN_Raw['Data Owner Organization'].str.contains("L3-71000325 OPS FM Ops ǀ NA ǀ GBS India", na=False, case=False)]
Org9 = Sahira_KCI_IN_Raw[Sahira_KCI_IN_Raw['Data Owner Organization'].str.contains("L3-71000325 OPS FM Ops ǀ NA ǀ GBS Malaysia", na=False, case=False)]
Org10 = Sahira_KCI_IN_Raw[Sahira_KCI_IN_Raw['Data Owner Organization'].str.contains("L1-90000025 CCIB Financial Markets ǀ NA ǀ AUSTRALIA", na=False, case=False)]
Org10_1 = Sahira_KCI_IN_Raw[Sahira_KCI_IN_Raw['Data Owner Organization'].str.contains("L1-90000025 CCIB Financial Markets ǀ NA ǀ GROUP", na=False, case=False)]
Org10_2 = Sahira_KCI_IN_Raw[Sahira_KCI_IN_Raw['Data Owner Organization'].str.contains("L1-90000025 CCIB Financial Markets ǀ NA ǀ GBS India", na=False, case=False)]

KCI_Hub_IN = Org8.append([Org9, Org10,Org10_1,Org10_2])
print("Shape of Sahira_KCI_IN_Raw after Data Owner Organization filter: ", KCI_Hub_IN.shape)

# Add 'Data From' column
KCI_Hub_IN['Data From'] = 'KCI_Hub_IN'

###############################################################
####################### Sahira CST IN Data ####################
###############################################################

# Reading the Raw Riz KCI File with all data
Sahira_CST_IN = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Base Raw Data//Sahira Data//CST - Jul 24.xlsx')
# Reading specific data alone from the sheet
Sahira_CST_IN_Raw = pd.read_excel(Sahira_CST_IN, 'CST - Jul 24')
print("Shape of Sahira_CST_IN_Raw: ", Sahira_CST_IN_Raw.shape)

# Filtering based on Due Date
Sahira_CST_IN_Raw['Due Date'] = pd.to_datetime(Sahira_CST_IN_Raw['Due Date'])
Sahira_CST_IN_Raw = Sahira_CST_IN_Raw[(Sahira_CST_IN_Raw['Due Date'].ge(Start_Date)) & (Sahira_CST_IN_Raw['Due Date'].le(End_Date))]
print("Shape of Sahira_CST_IN_Raw after Due Date filter: ", Sahira_CST_IN_Raw.shape)

# Filtering for Org Nodes
# GBS INDIA CST Hub Org nodes (Sahira’s):
Org11 = Sahira_CST_IN_Raw[Sahira_CST_IN_Raw['Tester Organization'].str.contains("L3-71000325 OPS FM Ops ǀ NA ǀ GBS India", na=False, case=False)]
Org12 = Sahira_CST_IN_Raw[Sahira_CST_IN_Raw['Tester Organization'].str.contains("L3-71000325 OPS FM Ops ǀ NA ǀ GBS Malaysia", na=False, case=False)]
Org13 = Sahira_CST_IN_Raw[Sahira_CST_IN_Raw['Tester Organization'].str.contains("L1-90000025 CCIB Financial Markets ǀ NA ǀ AUSTRALIA", na=False, case=False)]
Org14 = Sahira_CST_IN_Raw[Sahira_CST_IN_Raw['Tester Organization'].str.contains("L1-90000025 CCIB Financial Markets ǀ NA ǀ GROUP", na=False, case=False)]
Org14_1 = Sahira_CST_IN_Raw[Sahira_CST_IN_Raw['Tester Organization'].str.contains("L1-90000025 CCIB Financial Markets ǀ NA ǀ GBS India", na=False, case=False)]
CST_Hub_IN = Org11.append([Org12, Org13, Org14,Org14_1])
print("Shape of CST_Hub_IN after Tester Organization Filter: ", CST_Hub_IN.shape)

# Add 'Data From' column
CST_Hub_IN['Data From'] = 'CST_Hub_IN'

##############################################################################################################################
###################################################    MY - HUB & NON_HUB    #################################################
##############################################################################################################################

###############################################################
####################### Riz KCI MY Data #######################
###############################################################

# Reading the Raw Riz KCI File with all data
Riz_KCI_MY = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Base Raw Data//Riz Data//Metric Responses KCI -Completed.xlsx')
# Reading specific data alone from the sheet
Riz_KCI_MY_Raw = pd.read_excel(Riz_KCI_MY, 'Metric Responses KCI -Completed')
print("Shape of Riz_KCI_MY_Raw: ", Riz_KCI_MY_Raw.shape)

# Filtering based on Due Date
Riz_KCI_MY_Raw['Due Date'] = pd.to_datetime(Riz_KCI_MY_Raw['Due Date'])
Riz_KCI_MY_Raw = Riz_KCI_MY_Raw[(Riz_KCI_MY_Raw['Due Date'].ge(Start_Date)) & (Riz_KCI_MY_Raw['Due Date'].le(End_Date))]
print("Shape of Riz_KCI_MY_Raw after Due Date Filter: ", Riz_KCI_MY_Raw.shape)

# Filtering for Org Nodes
# Tester Organization (For Non-Hub Malaysia)
Org15 = Riz_KCI_MY_Raw[Riz_KCI_MY_Raw['Data Owner Organization'].str.contains("L3-70000101 CCIB FM Head Financial Mkt ǀ NA ǀ GBS India", na=False, case=False)]
Org16 = Riz_KCI_MY_Raw[Riz_KCI_MY_Raw['Data Owner Organization'].str.contains("99999981 Central Control Group ǀ NA ǀ GROUP", na=False, case=False)]
Org17 = Riz_KCI_MY_Raw[Riz_KCI_MY_Raw['Data Owner Organization'].str.contains("99999982 Central Control Group ǀ NA ǀ GROUP", na=False, case=False)]
Org18 = Riz_KCI_MY_Raw[Riz_KCI_MY_Raw['Data Owner Organization'].str.contains("L3-71000325 OPS FM Operations ǀ NA ǀ VIETNAM", na=False, case=False)]
Org19 = Riz_KCI_MY_Raw[Riz_KCI_MY_Raw['Data Owner Organization'].str.contains("L3-71000325 OPS FM Operations ǀ NA ǀ GBS Malaysia", na=False, case=False)]
KCI_Non_Hub_MY = Org15.append([Org16, Org17, Org18, Org19])
print("Shape of KCI_Non_Hub_MY after Data Owner Organization Filter: ", KCI_Non_Hub_MY.shape)

# Reading the Raw FMO Non-Hubbed data
print("Choose the Month sheet here manually to the needed month like - Feb 2024")
FMO_Non_Hubbed = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Base Raw Data//Riz Data//CCG Control Monitoring BAU Tracker FMO Non-Hubbed_2024.xlsx')
# Reading specific data alone from the sheet
FMO_Non_Hubbed_Raw = pd.read_excel(FMO_Non_Hubbed, 'July 2024')
print("Shape of FMO_Non_Hubbed_Raw: ", FMO_Non_Hubbed_Raw.shape)
# Filter Indicator type
KCI_FMO_Non_Hubbed_Raw = FMO_Non_Hubbed_Raw[FMO_Non_Hubbed_Raw['Indicator Type'].str.contains("KCI", na=False, case=False)]
print("Shape of KCI_FMO_Non_Hubbed_Raw after vlookup: ", KCI_FMO_Non_Hubbed_Raw.shape)

# vlookup with the Test Ids from KCI_FMO_Non_Hubbed_Raw
KCI_Non_Hub_MY = pd.merge(KCI_Non_Hub_MY,  KCI_FMO_Non_Hubbed_Raw['Test ID'],  left_on ='Metric Response ID', right_on = 'Test ID', how ='inner')
KCI_Non_Hub_MY['Data From'] = 'KCI_Non_Hub_MY'
# KCI_Non_Hub_MY.columns


###############################################################
####################### Riz CST MY Data #######################
###############################################################

# Reading the Raw Riz CST File with all data
Riz_CST_MY = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Base Raw Data//Riz Data//Control Sample Test Report .xlsx')
# Reading specific data alone from the sheet
Riz_CST_MY_Raw = pd.read_excel(Riz_CST_MY, 'Control Sample Test Report ')
print("Shape of Riz_CST_MY_Raw: ", Riz_CST_MY_Raw.shape)

# Filtering based on Due Date
Riz_CST_MY_Raw['Due Date'] = pd.to_datetime(Riz_CST_MY_Raw['Due Date'])
Riz_CST_MY_Raw = Riz_CST_MY_Raw[(Riz_CST_MY_Raw['Due Date'].ge(Start_Date)) & (Riz_CST_MY_Raw['Due Date'].le(End_Date))]
print("Shape of Riz_CST_MY_Raw after Due Date filter: ", Riz_CST_MY_Raw.shape)

# Filtering for Org Nodes
# Tester Organization (For Non-Hub Malaysia)
Org20 = Riz_CST_MY_Raw[Riz_CST_MY_Raw['Tester Organization'].str.contains("L3-70000101 CCIB FM Head Financial Mkt ǀ NA ǀ GBS India", na=False, case=False)]
Org21 = Riz_CST_MY_Raw[Riz_CST_MY_Raw['Tester Organization'].str.contains("99999981 Central Control Group ǀ NA ǀ GROUP", na=False, case=False)]
Org22 = Riz_CST_MY_Raw[Riz_CST_MY_Raw['Tester Organization'].str.contains("99999982 Central Control Group ǀ NA ǀ GROUP", na=False, case=False)]
Org23 = Riz_CST_MY_Raw[Riz_CST_MY_Raw['Tester Organization'].str.contains("L3-71000325 OPS FM Operations ǀ NA ǀ VIETNAM", na=False, case=False)]
Org24 = Riz_CST_MY_Raw[Riz_CST_MY_Raw['Tester Organization'].str.contains("L3-71000325 OPS FM Operations ǀ NA ǀ GBS Malaysia", na=False, case=False)]
CST_Non_Hub_MY = Org20.append([Org21, Org22, Org23, Org24])
print("Shape of CST_Non_Hub_MY after Tester Organization filter: ", CST_Non_Hub_MY.shape)

# Reading the Raw FMO Non-Hubbed data
print("Choose the Month sheet here manually to the needed month like - Feb 2024")
FMO_Non_Hubbed = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Base Raw Data//Riz Data//CCG Control Monitoring BAU Tracker FMO Non-Hubbed_2024.xlsx')
# Reading specific data alone from the sheet
FMO_Non_Hubbed_Raw = pd.read_excel(FMO_Non_Hubbed, 'July 2024')
print("Shape of FMO_Non_Hubbed_Raw: ", FMO_Non_Hubbed_Raw.shape)
# Filter Indicator type
CST_FMO_Non_Hubbed_Raw = FMO_Non_Hubbed_Raw[FMO_Non_Hubbed_Raw['Indicator Type'].str.contains("CST", na=False, case=False)]
print("Shape of CST_FMO_Non_Hubbed_Raw: ", CST_FMO_Non_Hubbed_Raw.shape)

# vlookup with the Test Ids from CST_FMO_Non_Hubbed_Raw
CST_Non_Hub_MY = pd.merge(CST_Non_Hub_MY,  CST_FMO_Non_Hubbed_Raw['Test ID'],  left_on ='Execution ID', right_on = 'Test ID', how ='inner')
CST_Non_Hub_MY['Data From'] = 'CST_Non_Hub_MY'
# CST_Non_Hub_MY.columns


###############################################################
####################### Sahira KCI MY Data ####################
###############################################################

# Reading the Raw Riz KCI File with all data
Sahira_KCI_MY = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Base Raw Data//Sahira Data//Metric Response - Jul 24.xlsx')
# Reading specific data alone from the sheet
Sahira_KCI_MY_Raw = pd.read_excel(Sahira_KCI_MY, 'Metric Response - Jul 24')
print("Shape of Sahira_KCI_MY_Raw: ", Sahira_KCI_MY_Raw.shape)

# Filtering based on Due Date
Sahira_KCI_MY_Raw['Responded Date'] = pd.to_datetime(Sahira_KCI_MY_Raw['Responded Date'])
Sahira_KCI_MY_Raw = Sahira_KCI_MY_Raw[(Sahira_KCI_MY_Raw['Responded Date'].ge(Start_Date)) & (Sahira_KCI_MY_Raw['Responded Date'].le(End_Date))]
print("Shape of Sahira_KCI_MY_Raw after Responded Date filter: ", Sahira_KCI_MY_Raw.shape)

# Filtering for Org Nodes
# GBS INDIA KCI Hub Org nodes (Sahira’s):
Org25 = Sahira_KCI_MY_Raw[Sahira_KCI_MY_Raw['Data Owner Organization'].str.contains("L3-71000325 OPS FM Ops ǀ NA ǀ GBS Malaysia", na=False, case=False)]
Org26 = Sahira_KCI_MY_Raw[Sahira_KCI_MY_Raw['Data Owner Organization'].str.contains("L3-71000325 OPS FM Operations ǀ NA ǀ GBS Malaysia", na=False, case=False)]
KCI_Hub_MY = Org25.append([Org26])
print("Shape of KCI_Hub_MY after Data Owner Organization filter: ", KCI_Hub_MY.shape)

# Add 'Data From' column
KCI_Hub_MY['Data From'] = 'KCI_Hub_MY'

###############################################################
####################### Sahira CST MY Data ####################
###############################################################

# Reading the Raw Riz KCI File with all data
Sahira_CST_MY = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Base Raw Data//Sahira Data//CST - Jul 24.xlsx')
# Reading specific data alone from the sheet
Sahira_CST_MY_Raw = pd.read_excel(Sahira_CST_MY, 'CST - Jul 24')
print("Shape of Sahira_CST_MY_Raw: ", Sahira_CST_MY_Raw.shape)

# Filtering based on Due Date
Sahira_CST_MY_Raw['Due Date'] = pd.to_datetime(Sahira_CST_MY_Raw['Due Date'])
Sahira_CST_MY_Raw = Sahira_CST_MY_Raw[(Sahira_CST_MY_Raw['Due Date'].ge(Start_Date)) & (Sahira_CST_MY_Raw['Due Date'].le(End_Date))]
print("Shape of Sahira_CST_MY_Raw after Due Date filter: ", Sahira_CST_MY_Raw.shape)

# Filtering for Org Nodes
# GBS INDIA CST Hub Org nodes (Sahira’s):
Org27 = Sahira_CST_MY_Raw[Sahira_CST_MY_Raw['Tester Organization'].str.contains("L3-71000325 OPS FM Ops ǀ NA ǀ GBS Malaysia", na=False, case=False)]
Org28 = Sahira_CST_MY_Raw[Sahira_CST_MY_Raw['Tester Organization'].str.contains("L3-71000325 OPS FM Operations ǀ NA ǀ GBS Malaysia", na=False, case=False)]
CST_Hub_MY = Org27.append([Org28])
print("Shape of CST_Hub_MY after Tester Organization filter: ", CST_Hub_MY.shape)

# Add 'Data From' column
CST_Hub_MY['Data From'] = 'CST_Hub_MY'

###############################################################
################## Consolidating KCI & CST Data ###############
###############################################################

KCI_All_Data = KCI_Non_Hub_IN.append([KCI_Hub_IN, KCI_Non_Hub_MY, KCI_Hub_MY])
CST_All_Data = CST_Non_Hub_IN.append([CST_Hub_IN, CST_Non_Hub_MY,CST_Hub_MY])

os.chdir('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Python Workings')

KCI_All_Data.to_excel("Output - KCI_All_Data_Temp.xlsx")
CST_All_Data.to_excel("Output - CST_All_Data_Temp.xlsx")

print("Excluding below task names – This is only for MY and not IN")

# Task Name / Tracked Item Name (EXCLUDE the below)
# Excluding below task names – This is only for MY and not IN
print("""
Manually to exclude the task names mentioned below from both the data set KCI_All_Data & CST_All_Data using vlookup:
      FM FO	Task Name / Tracked Item Name (EXCLUDE the below)
1	Telephone Monitoring Supervisory Check_CST
2	FM \| TRADING REGULATIONS – SALES \| ISDA STAYS/MARGIN RULES REVIEWS \| MARGIN REFORM RULE REVIEW (GENUINE BREACHES) CST
3	FM \\| Trading Regulations - Sales\\| Dodd Frank Disclosure \\| Pre Trade Mid Market Disclosure PTMM CST
4	FM \\| Trading Regulations – Sales \\| ISDA Stays/Margin Rules Reviews \\| Margin Reform Rule Review (Recalled FMSW Workflows) CST
5	FM \\| FOFC - SALES \\| MiFID Product Governance Controls I Ensure relevant transactions submitted for Approval CST
6	FM \\| FOFC - SALES \\| MiFID Product Governance Controls I Distributors feedback CST
7	FM \\| FOFC - Sales \\| Appropriateness Post Trade Monitoring \\| Appropriateness Post Trade Monitoring - CST
8	FM \| MARKET ABUSE - TRADING \| AHA COVID CHECKS - Recorded Retrievable \| CST3
9	FM \| Market Abuse - Trading \| AHA COVID checks \| BlueJeans Meeting Recording CST
10	FM \| MARKET ABUSE - TRADING \| AHA COVID CHECKS - Softphone Lines \| CST2
11	FM \\| FOFC - SALES \\| MiFID Product Governance Controls \\| Target Market Assessment Sent CST
12	FM \| Market Abuse - Trading \| AHA COVID checks \| EOD Update CST
13	FM \\| FOFC - SALES \\| MiFID Product Governance Controls I NPAC noting & DTAG sharepoint completeness CST
14	FM \| Mis-selling \| Pre-Deal \| Failure to Comply with Market Commentary or Colour Requirements \| Adherence to Market Colour Market Commentary Investment Recommendation and marketing CST
15	Supervision of Marketing Materials CST
""")

# Reading the KCI_All_Data
KCI_All_Data = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Python Workings//Output - KCI_All_Data_Temp.xlsx')
# Reading specific data alone from the sheet
KCI_All_Data = pd.read_excel(KCI_All_Data, 'Sheet1')
print("Shape of KCI_All_Data: ", KCI_All_Data.shape)

# Reading the CST_All_Data
CST_All_Data = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Python Workings//Output - CST_All_Data_Temp.xlsx')
# Reading specific data alone from the sheet
CST_All_Data = pd.read_excel(CST_All_Data, 'Sheet1')
print("Shape of CST_All_Data: ", CST_All_Data.shape)

###############################################################
####################### KCI Pivot #############################
###############################################################

print ("""Frame the 1st column "Data From" by consolidating all KCI data
       Ensure to filter the ‘Due Date” & ‘CCG’ for the Raw base data before groupby. 
       Fill all the empty rows in 'Additional Comments' also all other columns with 
       EMPTY rows in GROUP BY with NIL - so counts dont differ in groupby
       """)

KCI_All_Data['Additional Comments'] = KCI_All_Data['Additional Comments'].replace("", np.NaN).fillna('NA')
CST_All_Data['Test Result'] = CST_All_Data['Test Result'].replace("", np.NaN).fillna('NA')
print("Output the final consolidated data if pivots have to be tried out manually")
KCI_All_Data.to_excel("Output - KCI_All_Data.xlsx")
CST_All_Data.to_excel("Output - CST_All_Data.xlsx")


# Raw_Data = pd.read_excel('MY&IN_New_Stats.xlsx')
# Reading the Raw File with all data
# xls = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//KCI_All_Data - Mar.xlsx')
# Reading specific data alone from the sheet
# KCI_Raw_Filtered_Non_Hub_MY = pd.read_excel(xls, 'KCI_All_Data')


KCI_Raw_Filtered_Non_Hub_MY = KCI_All_Data

KCI_Pivot = KCI_Raw_Filtered_Non_Hub_MY.groupby(['Monitored for Organizations','Data From','Metric ID','Metric', 'Additional Comments'])['Metric ID','Response Value'].agg(['count','sum']).reset_index()
KCI_Pivot.columns

# Countries Sampled - to identify data for specific Countries
KCI_Specific_Samples = KCI_Pivot[KCI_Pivot['Additional Comments'].str.contains('Country Sampled', na=False, case=False)]
# Excluding Nil Samples
KCI_Specific_Samples = KCI_Specific_Samples[~KCI_Specific_Samples['Additional Comments'].str.contains('Nil Sample', na=False, case=False)]
KCI_Specific_Samples = KCI_Specific_Samples[~KCI_Specific_Samples['Additional Comments'].str.contains('Country Sampled (Applicable for HUB): Nil', na=False, case=False)]
KCI_Specific_Samples = KCI_Specific_Samples[~KCI_Specific_Samples['Additional Comments'].str.contains('Country Sampled (Applicable for HUB): Nil Sample', na=False, case=False)]
KCI_Specific_Samples = KCI_Specific_Samples[~KCI_Specific_Samples['Additional Comments'].str.contains('2.Country Sampled (Applicable for HUB): Nil', na=False, case=False)]
KCI_Specific_Samples = KCI_Specific_Samples[~KCI_Specific_Samples['Additional Comments'].str.contains('2.Country Sampled (Applicable for HUB): Nil Sample', na=False, case=False)]
KCI_Specific_Samples = KCI_Specific_Samples[~KCI_Specific_Samples['Additional Comments'].str.contains('2.Country Sampled (Applicable for HUB): Nil Samples', na=False, case=False)]

KCI_Pivot.to_excel('Output - KCI_Pivot.xlsx', merge_cells=False)
KCI_Specific_Samples.to_excel('Output - KCI_Specific_Samples.xlsx', merge_cells=False)

print (""" In the Output - Refer Column F - 'Metric ID.count' for Metric ID Count and Column I - 'Response Value.sum' 
       for the Total Exceptions. DELETE Columns G & H (Metric ID.sum, Response Value.count as it is not needed 
     In the excel "KCI_Specific_Samples.xlsx" just copy and paste "Additional Comments" in next sheet and do "Data" - "Text to columns" 
     choose "Other" and automatically it is delimited based on 2. Country Sampled. So we can easily extract "Country Sampled" alone.
      Now paste this newly created column in the same sheet in KCI_Specific_Samples.xlsx in the last. We can now apply filter on this last column and 
      filter Ex: IN, India etc... then we can find the Country sampled countrywise
      """)


# Pivot and obtaining data for each country:

# Monitored_Org = KCI_Pivot['Monitored for Organizations'].unique()
Country_List = ['ANGOLA','AUSTRALIA','BAHRAIN','BANGLADESH','BOTSWANA','BRAZIL', 'CAMEROON','CHINA','COTE','DUBAI','EGYPT','FRANCE','GAMBIA','GBS India',
                'GBS Malaysia','GERMANY','GHANA','GROUP','HONG KONG','INDIA','INDONESIA','IRAQ','JAPAN', 'JORDAN', 'KENYA','KOREA', 'MACAU', 'MALAYSIA',
                'MAURITIUS','NEPAL','NIGERIA','OMAN','PAKISTAN','PHILIPPINES','QATAR','SAUDI ARABIA','SIERRA LEONE','SINGAPORE',
                'SOUTH AFRICA','SRI LANKA','TANZANIA','TAIWAN','TANZANIA','THAILAND','UGANDA','UNITED ARAB EMIRATES','UNITED KINGDOM',
                'UNITED STATES OF AMERICA','VIETNAM','ZAMBIA','ZIMBABWE']

# Output to excel sheet as per needed countries
counter = 0
for each in Country_List:
    print(each)
    subset_KCI = KCI_Pivot[KCI_Pivot['Monitored for Organizations'].str.contains(each)]
    print("***********************************************************************")
    print("KCI subset of:", each)
    print("***********************************************************************")
    # https://stackoverflow.com/questions/67519829/writing-multiple-data-frame-in-same-excel-sheet-one-below-other-in-python
    print("Total check for the Feb 24:\n", subset_KCI["Data From"].value_counts())
    df1 = pd.DataFrame(subset_KCI["Data From"].value_counts())
    # print("Total check for the Feb 24:\n", subset["Data From"].value_counts().tolist())
    print("Total exceptions for the Feb 24:\n", subset_KCI[(             'Response Value',   'sum')].value_counts())
    df2 = pd.DataFrame(subset_KCI[(             'Response Value',   'sum')].value_counts())
    print(subset_KCI)
    df3 = pd.DataFrame(subset_KCI)
    counter += 1
    print ("Iteration count: ", counter)
    excel_name = 'C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Python Workings//Stat_Subset//KCI_Subset//' + 'KCI - ' + each + '.xlsx'
    dfs_kci = [df1, df2, df3]
    startrow = 0
    with pd.ExcelWriter(excel_name) as writer:
        for df in dfs_kci:
            df.to_excel(writer, engine="xlsxwriter", startrow=startrow)
            startrow += (df.shape[0] + 10)


###############################################################
####################### CST Pivot #############################
###############################################################

print ("""Frame the 1st column "Data From" by consolidating all CST data
       Ensure to filter the ‘Due Date” for the Raw base data before groupby. 
       Fill all the empty rows in 'Test Result' with NIL also all other columns with 
       EMPTY rows in GROUP BY with NIL - so counts dont differ in groupby
       """)

# Raw_Data = pd.read_excel('MY&IN_New_Stats.xlsx')
# Reading the Raw File with all data
# xls = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//CST_All_Data - Mar.xlsx')
# Reading specific data alone from the sheet
# CST_Raw_Filtered_Non_Hub_MY = pd.read_excel(xls, 'CST_All_Data')

CST_Raw_Filtered_Non_Hub_MY = CST_All_Data

CST_Pivot = CST_Raw_Filtered_Non_Hub_MY.groupby(['Tested Organization','Data From','Plan ID','Plan Name', 'Test Result'])['Plan ID','Exceptions'].agg(['count','sum']).reset_index()
CST_Pivot.columns

# Countries Sampled - to identify data for specific Countries
CST_Specific_Samples = CST_Pivot[CST_Pivot['Test Result'].str.contains('Country Sampled', na=False, case=False)]
# Excluding Nil Samples
CST_Specific_Samples = CST_Specific_Samples[~CST_Specific_Samples['Test Result'].str.contains('Country Sampled: Nil Sample', na=False, case=False)]
#  Remove Country Sampled: Nil
CST_Specific_Samples = CST_Specific_Samples[~CST_Specific_Samples['Test Result'].str.contains('Country Sampled: Nil', na=False, case=False)]
CST_Specific_Samples = CST_Specific_Samples[~CST_Specific_Samples['Test Result'].str.contains('Country Sampled (Applicable for HUB): Nil', na=False, case=False)]
CST_Specific_Samples = CST_Specific_Samples[~CST_Specific_Samples['Test Result'].str.contains('Country Sampled (Applicable for HUB): Nil Sample', na=False, case=False)]
CST_Specific_Samples = CST_Specific_Samples[~CST_Specific_Samples['Test Result'].str.contains('2.Country Sampled (Applicable for HUB): Nil', na=False, case=False)]
CST_Specific_Samples = CST_Specific_Samples[~CST_Specific_Samples['Test Result'].str.contains('2.Country Sampled (Applicable for HUB): Nil Sample', na=False, case=False)]
CST_Specific_Samples = CST_Specific_Samples[~CST_Specific_Samples['Test Result'].str.contains('2.Country Sampled (Applicable for HUB): Nil Samples', na=False, case=False)]


CST_Pivot.to_excel('Output - CST_Pivot.xlsx', merge_cells=False)
CST_Specific_Samples.to_excel('Output - CST_Specific_Samples.xlsx', merge_cells=False)

print (""" In the Output - Refer Column F -  'Plan ID.count' for Metric ID Count and Column I - 'Exceptions.sum' 
       for the Total Exceptions. DELETE Columns G & H (Plan ID.sum, Exceptions.count as it is not needed 
       In the excel "CST_Specific_Sampeles.xlsx" just copy and paste "Test Result" in next sheet and do "Data" - "Text to columns" 
      choose "Other" and automatically it is delimited based on 2. Country Sampled. So we can easily extract "Country Sampled" alone.
      Now paste this newly created column in the same sheet in CST_Specific_Samples.xlsx in the last. We can now apply filter on this alst column and 
      filter Ex: IN, India etc... then we can find the Country sampled countrywise
      """)
      


# Pivot and obtaining data for each country:

# Monitored_Org = KCI_Pivot['Monitored for Organizations'].unique()
Country_List = ['ANGOLA','AUSTRALIA','BAHRAIN','BANGLADESH','BOTSWANA','BRAZIL', 'CAMEROON','CHINA','COTE','DUBAI','EGYPT','FRANCE','GAMBIA','GBS India',
                'GBS Malaysia','GERMANY','GHANA','GROUP','HONG KONG','INDIA','INDONESIA','IRAQ','JAPAN', 'JORDAN', 'KENYA','KOREA', 'MACAU', 'MALAYSIA',
                'MAURITIUS','NEPAL','NIGERIA','OMAN','PAKISTAN','PHILIPPINES','QATAR','SAUDI ARABIA','SIERRA LEONE','SINGAPORE',
                'SOUTH AFRICA','SRI LANKA','TANZANIA','TAIWAN','TANZANIA','THAILAND','UGANDA','UNITED ARAB EMIRATES','UNITED KINGDOM',
                'UNITED STATES OF AMERICA','VIETNAM','ZAMBIA','ZIMBABWE']

# Output to excel sheet as per needed countries
counter = 0
for each in Country_List:
    print(each)
    subset_CST = CST_Pivot[CST_Pivot['Tested Organization'].str.contains(each)]
    print("***********************************************************************")
    print("CST Subset of:", each)
    print("***********************************************************************")
    # https://stackoverflow.com/questions/67519829/writing-multiple-data-frame-in-same-excel-sheet-one-below-other-in-python
    print("Total check for the Feb 24:\n", subset_CST["Data From"].value_counts())
    df4 = pd.DataFrame(subset_CST["Data From"].value_counts())
    # print("Total check for the Feb 24:\n", subset["Data From"].value_counts().tolist())
    print("Total exceptions for the Feb 24:\n", subset_CST[(             'Exceptions',   'sum')].value_counts())
    df5 = pd.DataFrame(subset_CST[(             'Exceptions',   'sum')].value_counts())
    print(subset_CST)
    df6 = pd.DataFrame(subset_CST)
    counter += 1
    print ("Iteration count: ", counter)
    excel_name = 'C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Python Workings//Stat_Subset//CST_Subset//' + 'CST - ' +  each + '.xlsx'
    dfs_cst = [df4, df5, df6]
    startrow = 0
    with pd.ExcelWriter(excel_name) as writer:
        for df in dfs_cst:
            df.to_excel(writer, engine="xlsxwriter", startrow=startrow)
            startrow += (df.shape[0] + 10)



###################################################################
########################## CST West Africa ########################
###################################################################

# Logic for Africa West Countries alone to extract GROUP data details
print("Logic for Africa West Countries alone to extract GROUP data details")

os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Python Workings")

# Reading the Raw Riz KCI File with all data
CST_Specific = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Python Workings//Output - CST_Specific_Samples.xlsx')
# Reading specific data alone from the sheet
CST_Specific = pd.read_excel(CST_Specific, 'Sheet1')
print("Shape of CST_Specific: ", CST_Specific.shape)

# Filter only GROUP data

CST_Specific = CST_Specific[CST_Specific['Tested Organization.'].str.contains("GROUP", na=False, case=False)]
CST_Specific["CONCATENATE"] = CST_Specific["Plan ID."] + "    " + CST_Specific["Plan Name."]

# Filtering for Global & West African countries
searchfor = ['Global', 'CM', 'CDI', 'GM', 'GH', 'NG', 'SL', 'Cote', 'Cameroon', 'Ghana', 'Gambia', 'Nigeria', 'Sierra']
CST_Specific = CST_Specific[CST_Specific['Test Result.'].str.contains('|'.join(searchfor))]


CST_Cameroon_WA = CST_Specific[CST_Specific['Test Result.'].str.contains("Global|Cameroon|CM", na=False, case=True)]
CST_Cameroon_WA["Country_Count"] = CST_Cameroon_WA['Test Result.'].str.count("Global|Cameroon|CM")
# Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled
print("Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled, this logic works for Global as well")
CST_Cameroon_WA = CST_Cameroon_WA[CST_Cameroon_WA["Country_Count"] > 1]
# CST_Cameroon_WA["CONCATENATE"] = CST_Cameroon_WA["Plan ID."] + "    " + CST_Cameroon_WA["Plan Name."]
print("Shape of Cameroon_WA: ", CST_Cameroon_WA.shape)

CST_CDI_WA = CST_Specific[CST_Specific['Test Result.'].str.contains("Global|Cote|CDI", na=False, case=True)]
CST_CDI_WA["Country_Count"] = CST_CDI_WA['Test Result.'].str.count("Global|Cote|CDI")
# Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled
print("Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled, this logic works for Global as well")
CST_CDI_WA = CST_CDI_WA[CST_CDI_WA["Country_Count"] > 1]
# CST_CDI_WA["CONCATENATE"] = CST_CDI_WA["Plan ID."] + "    " + CST_CDI_WA["Plan Name."]
print("Shape of CDI_WA: ", CST_CDI_WA.shape)

CST_Gambia_WA = CST_Specific[CST_Specific['Test Result.'].str.contains("Global|Gambia|GM", na=False, case=True)]
CST_Gambia_WA["Country_Count"] = CST_Gambia_WA['Test Result.'].str.count("Global|Gambia|GM")
# Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled
print("Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled, this logic works for Global as well")
CST_Gambia_WA = CST_Gambia_WA[CST_Gambia_WA["Country_Count"] > 1]
# CST_Gambia_WA["CONCATENATE"] = CST_Gambia_WA["Plan ID."] + "    " + CST_Gambia_WA["Plan Name."]
print("Shape of Gambia_WA: ", CST_Gambia_WA.shape)

CST_Ghana_WA = CST_Specific[CST_Specific['Test Result.'].str.contains("Global|Ghana|GH", na=False, case=True)]
CST_Ghana_WA["Country_Count"] = CST_Ghana_WA['Test Result.'].str.count("Global|Ghana|GH")
# Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled
print("Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled, this logic works for Global as well")
CST_Ghana_WA = CST_Ghana_WA[CST_Ghana_WA["Country_Count"] > 1]
# CST_Ghana_WA["CONCATENATE"] = CST_Ghana_WA["Plan ID."] + "    " + CST_Ghana_WA["Plan Name."]
print("Shape of Ghana_WA: ", CST_Ghana_WA.shape)

CST_Nigeria_WA = CST_Specific[CST_Specific['Test Result.'].str.contains("Global|Nigeria|NG", na=False, case=True)]
CST_Nigeria_WA["Country_Count"] = CST_Nigeria_WA['Test Result.'].str.count("Global|Nigeria|NG")
# Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled
print("Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled, this logic works for Global as well")
CST_Nigeria_WA = CST_Nigeria_WA[CST_Nigeria_WA["Country_Count"] > 1]
# CST_Nigeria_WA["CONCATENATE"] = CST_Nigeria_WA["Plan ID."] + "    " + CST_Nigeria_WA["Plan Name."]
print("Shape of Nigeria_WA: ", CST_Nigeria_WA.shape)

CST_SierraLeone_WA = CST_Specific[CST_Specific['Test Result.'].str.contains("Global|Sierra|SL", na=False, case=True)]
CST_SierraLeone_WA["Country_Count"] = CST_SierraLeone_WA['Test Result.'].str.count("Global|Sierra|SL")
# Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled
print("Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled, this logic works for Global as well")
CST_SierraLeone_WA = CST_SierraLeone_WA[CST_SierraLeone_WA["Country_Count"] > 1]
# CST_SierraLeone_WA["CONCATENATE"] = CST_SierraLeone_WA["Plan ID."] + "    " + CST_SierraLeone_WA["Plan Name."]
print("Shape of SierraLeone_WA: ", CST_SierraLeone_WA.shape)

# After Filtering output
os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Python Workings//Specific - West Africa")
CST_Specific.to_excel('Output - CST - CST_Specific_West_Africa.xlsx', merge_cells=False)
print("Shape of Consolidated CST_Specific for all combinations after all filters: ", CST_Specific.shape)

CST_Cameroon_WA.to_excel('Output - CST - Cameroon_WA.xlsx', merge_cells=False)
CST_CDI_WA.to_excel('Output - CST - CDI_WA.xlsx', merge_cells=False)
CST_Gambia_WA.to_excel('Output - CST -Gambia_WA.xlsx', merge_cells=False)
CST_Ghana_WA.to_excel('Output - CST -Ghana_WA.xlsx', merge_cells=False)
CST_Nigeria_WA.to_excel('Output - CST - Nigeria_WA.xlsx', merge_cells=False)
CST_SierraLeone_WA.to_excel('Output - CST - SierraLeone_WA.xlsx', merge_cells=False)


###################################################################
########################## KCI West Africa ########################
###################################################################

# Logic for Africa West Countries alone to extract GROUP data details
print("Logic for Africa West Countries alone to extract GROUP data details")

os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Python Workings")

# Reading the Raw Riz KCI File with all data
KCI_Specific = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Python Workings//Output - KCI_Specific_Samples.xlsx')
# Reading specific data alone from the sheet
KCI_Specific = pd.read_excel(KCI_Specific, 'Sheet1')
print("Shape of KCI_Specific: ", KCI_Specific.shape)

# Filter only GROUP data

KCI_Specific = KCI_Specific[KCI_Specific['Monitored for Organizations.'].str.contains("GROUP", na=False, case=False)]
KCI_Specific["CONCATENATE"] = KCI_Specific["Metric ID."] + "    " + KCI_Specific["Metric."]
    
# Filtering for Global & West African countries
searchfor = ['Global', 'CM', 'CDI', 'GM', 'GH', 'NG', 'SL', 'Cote', 'Cameroon', 'Ghana', 'Gambia', 'Nigeria', 'Sierra']
KCI_Specific = KCI_Specific[KCI_Specific['Additional Comments.'].str.contains('|'.join(searchfor))]

KCI_Cameroon_WA = KCI_Specific[KCI_Specific['Additional Comments.'].str.contains("Global|Cameroon|CM", na=False, case=True)]
KCI_Cameroon_WA["Country_Count"] = KCI_Cameroon_WA['Additional Comments.'].str.count("Global|Cameroon|CM")
# Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled
print("Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled, this logic works for Global as well")
KCI_Cameroon_WA = KCI_Cameroon_WA[KCI_Cameroon_WA["Country_Count"] > 1]
# KCI_Cameroon_WA["CONCATENATE"] = KCI_Cameroon_WA["Metric ID."] + "    " + KCI_Cameroon_WA["Metric."]
print("Shape of Cameroon_WA: ", KCI_Cameroon_WA.shape)

KCI_CDI_WA = KCI_Specific[KCI_Specific['Additional Comments.'].str.contains("Global|Cote|CDI", na=False, case=True)]
KCI_CDI_WA["Country_Count"] = KCI_CDI_WA['Additional Comments.'].str.count("Global|Cote|CDI")
# Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled
print("Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled, this logic works for Global as well")
KCI_CDI_WA = KCI_CDI_WA[KCI_CDI_WA["Country_Count"] > 1]
# KCI_CDI_WA["CONCATENATE"] = KCI_CDI_WA["Metric ID."] + "    " + KCI_CDI_WA["Metric."]
print("Shape of CDI_WA: ", KCI_CDI_WA.shape)

KCI_Gambia_WA = KCI_Specific[KCI_Specific['Additional Comments.'].str.contains("Global|Gambia|GM", na=False, case=True)]
KCI_Gambia_WA["Country_Count"] = KCI_Gambia_WA['Additional Comments.'].str.count("Global|Gambia|GM")
# Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled
print("Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled, this logic works for Global as well")
KCI_Gambia_WA = KCI_Gambia_WA[KCI_Gambia_WA["Country_Count"] > 1]
# KCI_Gambia_WA["CONCATENATE"] = KCI_Gambia_WA["Metric ID."] + "    " + KCI_Gambia_WA["Metric."]
print("Shape of Gambia_WA: ", KCI_Gambia_WA.shape)

KCI_Ghana_WA = KCI_Specific[KCI_Specific['Additional Comments.'].str.contains("Global|Ghana|GH", na=False, case=True)]
KCI_Ghana_WA["Country_Count"] = KCI_Ghana_WA['Additional Comments.'].str.count("Global|Ghana|GH")
# Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled
print("Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled, this logic works for Global as well")
KCI_Ghana_WA = KCI_Ghana_WA[KCI_Ghana_WA["Country_Count"] > 1]
# KCI_Ghana_WA["CONCATENATE"] = KCI_Ghana_WA["Metric ID."] + "    " + KCI_Ghana_WA["Metric."]
print("Shape of Ghana_WA: ", KCI_Ghana_WA.shape)

KCI_Nigeria_WA = KCI_Specific[KCI_Specific['Additional Comments.'].str.contains("Global|Nigeria|NG", na=False, case=True)]
KCI_Nigeria_WA["Country_Count"] = KCI_Nigeria_WA['Additional Comments.'].str.count("Global|Nigeria|NG")
# Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled
print("Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled, this logic works for Global as well")
KCI_Nigeria_WA = KCI_Nigeria_WA[KCI_Nigeria_WA["Country_Count"] > 1]
# KCI_Nigeria_WA["CONCATENATE"] = KCI_Nigeria_WA["Metric ID."] + "    " + KCI_Nigeria_WA["Metric."]
print("Shape of Nigeria_WA: ", KCI_Nigeria_WA.shape)

KCI_SierraLeone_WA = KCI_Specific[KCI_Specific['Additional Comments.'].str.contains("Global|Sierra|SL", na=False, case=True)]
KCI_SierraLeone_WA["Country_Count"] = KCI_SierraLeone_WA['Additional Comments.'].str.count("Global|Sierra|SL")
# Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled
print("Picking the count of countries more that 1 so the country is present in both Country Supporterd and Country sampled, this logic works for Global as well")
KCI_SierraLeone_WA = KCI_SierraLeone_WA[KCI_SierraLeone_WA["Country_Count"] > 1]
# KCI_SierraLeone_WA["CONCATENATE"] = KCI_SierraLeone_WA["Metric ID."] + "    " + KCI_SierraLeone_WA["Metric."]
print("Shape of SierraLeone_WA: ", KCI_SierraLeone_WA.shape)

# After Filtering output
os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SRM_New//New Stats//Python Workings//Specific - West Africa")
KCI_Specific.to_excel('Output - CST - KCI_Specific_West_Africa.xlsx', merge_cells=False)
print("Shape of Consolidated KCI_Specific for all combinations after all filters: ", KCI_Specific.shape)

KCI_Cameroon_WA.to_excel('Output - KCI - Cameroon_WA.xlsx', merge_cells=False)
KCI_CDI_WA.to_excel('Output - KCI - CDI_WA.xlsx', merge_cells=False)
KCI_Gambia_WA.to_excel('Output - KCI -Gambia_WA.xlsx', merge_cells=False)
KCI_Ghana_WA.to_excel('Output - KCI -Ghana_WA.xlsx', merge_cells=False)
KCI_Nigeria_WA.to_excel('Output - KCI - Nigeria_WA.xlsx', merge_cells=False)
KCI_SierraLeone_WA.to_excel('Output - KCI - SierraLeone_WA.xlsx', merge_cells=False)


# Record End Time and Calculate Execution Time
end_time = time.time()
execution_time = start_time - end_time
print("Execution time:",execution_time)