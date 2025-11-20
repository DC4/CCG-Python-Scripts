# -*- coding: utf-8 -*-
"""
Created on Wed Dec 21 20:04:19 2022

@author: 1510806
"""

# import pandas
import pandas as pd
import numpy as np
import os
# Import time module
import time
 
# record start time
start = time.time()

os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Py_Data_RRA_Automation")
#Issues = pd.read_csv('1022_Issues and Actions Report (1).csv', encoding='cp1252',  low_memory=False, error_bad_lines=False)

####################### Pivot - Issues #######################

Issues = pd.read_excel('1022_Issues and Actions Report (1).xlsx')
Issues = pd.DataFrame(Issues)
print(Issues.shape)
Issues_1 = Issues[['Risk ID', 'Risk Name', 'Issue ID', 'Owner', 'Rationale For Rating', 'Rating (GRAM)', 'Rating (CRAM)', 'Issue Title', 'Issue Description', 'Due Date', 'Action Due Date', 'Overdue Issues', 'Overdue Actions', 'Status', 'Impacted Organization - Country / Group', 'Owner Organization - Country / Group', 'Owner Organization']]
print(Issues_1.shape)
Issues_1.columns
Issues_1['Owner Organization - Country / Group'] = Issues_1['Owner Organization - Country / Group'].str.upper()
Issues_1['Owner Organization - Country / Group'] = Issues_1['Owner Organization - Country / Group'].str.strip()
Issues_1['Status'] = Issues_1['Status'].str.upper()
Issues_1['Status'] = Issues_1['Status'].str.strip()
Issues_1['Impacted Organization - Country / Group'] = Issues_1['Impacted Organization - Country / Group'].str.upper()
Issues_1['Impacted Organization - Country / Group'] = Issues_1['Impacted Organization - Country / Group'].str.strip()

# Filter
Status_exclude = ['CANCELLED', 'CLOSED']
Issues_1 = Issues_1[~Issues_1['Status'].isin(Status_exclude)]
Issues_1 = Issues_1[Issues_1['Impacted Organization - Country / Group'] == 'HONG KONG']
#Issues_HK = Issues_1[Issues_1['Status'] == 'HONG KONG']
# Issues_HK = Issues_1[Issues_1['Owner Organization - Country / Group'] == 'HONG KONG']
print(Issues_1.shape)

Pivot_Issues_1 = pd.pivot_table(
    data=Issues_1,
    index = ['Risk ID', 'Risk Name', 'Issue ID', 'Owner', 'Rationale For Rating', 'Rating (GRAM)', 'Rating (CRAM)', 'Issue Title', 'Issue Description', 'Due Date', 'Action Due Date', 'Overdue Issues', 'Overdue Actions', 'Status', 'Impacted Organization - Country / Group', 'Owner Organization - Country / Group'],
    values='Owner Organization',
    aggfunc='count'
).reset_index().replace('dummy',np.nan)

# Pivot_Issues_1.to_excel("Pivot_Issues_1.xlsx")

####################### Pivot - Events #######################

Events = pd.read_excel('1022_Internal Risk Events (1).xlsx')
Events = pd.DataFrame(Events)
print(Events.shape)
# Events_1 = Events[['Related Risk ID', 'Related Risk Name', 'Event ID', 'Event Title', 'Created By', 'Owner', 'Description', 'MRE', 'Created On', 'Date of Discovery', 'Closed Date', 'Country','Days to Log']]
Events_1 = Events[['Related Risk ID', 'Related Risk Name', 'Event ID', 'Event Title', 'Created By', 'Owner', 'Description', 'MRE', 'Created On', 'Date of Discovery', 'Closed Date', 'Country']]
print(Events_1.shape)
Events_1.columns
Events_1['Country'] = Events_1['Country'].str.upper()
Events_1['Country'] = Events_1['Country'].str.strip()
Events_1['Related Risk Name'] = Events_1['Related Risk Name'].str.upper()
Events_1['Related Risk Name'] = Events_1['Related Risk Name'].str.strip()

#Filter
Events_1 = Events_1[Events_1['Country'] == 'HONG KONG']
Events_1 = Events_1[Events_1['Related Risk Name'] == 'TREASURY MARKETS \| TRADE INITIATION, EXECUTION & CAPTURE \| ERRONEOUS TM TRADE-PROCESSING FAILURE']
# Events_HK = Events_1[Events_1['Country'] == 'HONG KONG']
print(Events_1.shape)

Pivot_Events_1 = pd.pivot_table(
    data=Events_1,
    index = ['Related Risk ID', 'Related Risk Name', 'Event ID', 'Event Title', 'Created By', 'Owner', 'Description', 'MRE', 'Created On', 'Date of Discovery', 'Closed Date', 'Country'],
    # values='Days to Log',
    values = 'Event ID',
    aggfunc='count'
).reset_index().replace('dummy',np.nan)

# Pivot_Events_1.to_excel("Pivot_Events_1.xlsx")

####################### Pivot - Overdue Metrics #######################

Metrics = pd.read_excel('1022_Risk and Control Monitor Report (1).xlsx')
Metrics = pd.DataFrame(Metrics)
print(Metrics.shape)
Metrics_1 = Metrics[['Risk ID', 'Risk Name', 'Indicator Type', 'Control Owner Bank ID', 'Control Owner', 'Metric ID/Plan ID', 'Name of Metric/Plan', 'Response Value/Exception No', 'Threshold Status/Result', 'Due Date', 'Date Submitted', 'Status', 'Monitored for/Tested Organization Country']]
print(Metrics_1.shape)
Metrics_1.columns
Metrics_1['Monitored for/Tested Organization Country'] = Metrics_1['Monitored for/Tested Organization Country'].str.upper()
Metrics_1['Monitored for/Tested Organization Country'] = Metrics_1['Monitored for/Tested Organization Country'].str.strip()
Metrics_1['Threshold Status/Result'] = Metrics_1['Threshold Status/Result'].str.upper()
Metrics_1['Threshold Status/Result'] = Metrics_1['Threshold Status/Result'].str.strip()
Metrics_1['Status'] = Metrics_1['Status'].str.upper()
Metrics_1['Status'] = Metrics_1['Status'].str.strip()
# Metrics_1_HK = Metrics_1[Metrics_1['Monitored for/Tested Organization Country'] == 'HONG KONG']

#Filter
Country_Details = ['GROUP', 'HONG KONG']
Metrics_1 = Metrics_1[Metrics_1['Monitored for/Tested Organization Country'].isin(Country_Details)]
Metrics_1 = Metrics_1[Metrics_1['Response Value/Exception No'] > 0.0]
Threshold_Details = ['PASSED', 'WITHIN THRESHOLD']
Metrics_1 = Metrics_1[~Metrics_1['Monitored for/Tested Organization Country'].isin(Threshold_Details)]
Metrics_1 = Metrics_1[Metrics_1['Due Date'] < '2022-06-01']
Status_exclude_met = ['CANCELLED', 'COMPLETED', 'BLANK', 'NAN']
Metrics_1 = Metrics_1[~Metrics_1['Status'].isin(Status_exclude_met)]
print(Metrics_1.shape)


Pivot_Metrics_1 = Metrics_1

#Pivot_Metrics_1 = pd.pivot_table(
#    data=Metrics_1,
#    index = ['Risk ID', 'Risk Name', 'Indicator Type', 'Control Owner Bank ID', 'Control Owner', 'Metric ID/Plan ID', 'Name of Metric/Plan', 'Response Value/Exception No', 'Threshold Status/Result', 'Due Date', 'Date Submitted', 'Status', 'Monitored for/Tested Organization Country'],
#    # values='Days to Log',
#    values = 'Risk ID',
#    aggfunc='count'
#).reset_index().replace('dummy',np.nan)

# Pivot_Metrics_1.to_excel("Pivot_Metrics_1.xlsx")

####################### Pivot - Metric Exceptions #######################

Metrics_Exceptions = pd.read_excel('1022_Risk and Control Monitor Report (1).xlsx')
Metrics_Exceptions = pd.DataFrame(Metrics_Exceptions)
print(Metrics_Exceptions.shape)
Metrics_2 = Metrics_Exceptions[['Risk ID', 'Risk Name', 'Indicator Type', 'Metric ID/Plan ID', 'Name of Metric/Plan', 'Test Results-CST', 'Exception Details-CST', 'Additional Comments-KCI/KRI', 'Response Value/Exception No', 'Threshold Status/Result', 'Monitored for/Tested Organization Country', 'Monitored for/Tested Organization', 'Date Submitted']]
print(Metrics_2.shape)
Metrics_2.columns
Metrics_2['Monitored for/Tested Organization Country'] = Metrics_2['Monitored for/Tested Organization Country'].str.upper()
Metrics_2['Monitored for/Tested Organization Country'] = Metrics_2['Monitored for/Tested Organization Country'].str.strip()
Metrics_2['Monitored for/Tested Organization'] = Metrics_2['Monitored for/Tested Organization'].str.upper()
Metrics_2['Monitored for/Tested Organization'] = Metrics_2['Monitored for/Tested Organization'].str.strip()
Metrics_2['Date Submitted'] = Metrics_2['Date Submitted'].astype(str).str.upper()
Metrics_2['Date Submitted'] = Metrics_2['Date Submitted'].astype(str).str.strip()
Metrics_2['Threshold Status/Result'] = Metrics_2['Threshold Status/Result'].astype(str).str.upper()
Metrics_2['Threshold Status/Result'] = Metrics_2['Threshold Status/Result'].astype(str).str.strip()

# Metrics_2_HK = Metrics_2[Metrics_2['Monitored for/Tested Organization Country'] == 'HONG KONG']

#Filter
Country_Details = ['GROUP', 'HONG KONG']
Metrics_2 = Metrics_2[Metrics_2['Monitored for/Tested Organization Country'].isin(Country_Details)]
Metrics_2 = Metrics_2[(Metrics_2['Date Submitted']!="null")]
searchfor = ['L3-50005903', 'L3-50005647', 'L3-50005903', 'L3-50005647', 'L6-50005647']
Metrics_2 = Metrics_2[Metrics_2['Monitored for/Tested Organization'].str.contains('|'.join(searchfor), case=False, na=False)]
Metrics_2 = Metrics_2[Metrics_2['Response Value/Exception No'] > 0.0]
Threshold_Details = ['PASSED', 'WITHIN THRESHOLD']
Metrics_2 = Metrics_2[~Metrics_2['Threshold Status/Result'].isin(Threshold_Details)]
print(Metrics_2.shape)

Pivot_Metrics_2 = Metrics_2

#Pivot_Metrics_2 = pd.pivot_table(
#    data = Metrics_2,
#    index = ['Risk ID', 'Risk Name', 'Indicator Type', 'Metric ID/Plan ID', 'Name of Metric/Plan', 'Test Results-CST', 'Exception Details-CST', 'Additional Comments-KCI/KRI', 'Response Value/Exception No', 'Threshold Status/Result', 'Monitored for/Tested Organization Country', 'Monitored for/Tested Organization', 'Date Submitted'],
#    # values='Days to Log',
#    values = 'Metric ID/Plan ID',
#    aggfunc='count'
#).reset_index().replace('dummy',np.nan)

# Pivot_Metrics_2.to_excel("Pivot_Metrics_2.xlsx")

####################### Pivot - Country RAs #######################

Risk_Register = pd.read_excel('1022_Risk Register (1).xlsx')
Risk_Register = pd.DataFrame(Risk_Register)
print(Risk_Register.shape)
Risk_1 = Risk_Register[['Risk ID', 'Risk', 'Final Residual Risk Rating', 'Overall Assessment Rationale', 'Assessed On', 'Assessed Organization']]
print(Risk_1.shape)
Risk_1.columns
Risk_1['Assessed Organization'] = Risk_1['Assessed Organization'].str.upper()
Risk_1['Assessed Organization'] = Risk_1['Assessed Organization'].str.strip()
# Risk_1_HK = Risk_1[Risk_1['Assessed Organization'] == 'HONG KONG']

#Filter
searchfor = ['L3-50005903', 'L1-90000043']
Risk_1 = Risk_1[Risk_1['Assessed Organization'].str.contains('|'.join(searchfor), case=False, na=False)]
print(Risk_1.shape)

Pivot_Risk_1 = Risk_1

#Pivot_Risk_1 = pd.pivot_table(
#    data = Risk_1,
#    index = ['Risk ID', 'Risk', 'Final Residual Risk Rating', 'Overall Assessment Rationale', 'Assessed On'],
#    # values='Days to Log',
#    values = 'Risk ID',
#    aggfunc='count'
#).reset_index().replace('dummy',np.nan)

# Pivot_Risk_1.to_excel("Pivot_Risk_1.xlsx")

####################### Pivot - Group RAs #######################

Risk_Register = pd.read_excel('1022_Risk Register (1).xlsx')
Risk_Register = pd.DataFrame(Risk_Register)
print(Risk_Register.shape)
Risk_2 = Risk_Register[['Risk ID', 'Risk', 'Final Residual Risk Rating', 'Prior Residual Rating', 'Overall Assessment Rationale', 'Assessed On', 'Assessed Organization']]
print(Risk_2.shape)
Risk_2.columns
Risk_2['Assessed Organization'] = Risk_2['Assessed Organization'].str.upper()
Risk_2['Assessed Organization'] = Risk_2['Assessed Organization'].str.strip()
# Risk_2_HK = Risk_2[Risk_2['Assessed Organization'] == 'HONG KONG']

#Filter
searchfor = ['L3-50005903', 'L1-90000043']
Risk_2 = Risk_2[Risk_2['Assessed Organization'].str.contains('|'.join(searchfor), case=False, na=False)]
print(Risk_2.shape)

Pivot_Risk_2 = Risk_2

#Pivot_Risk_2 = pd.pivot_table(
#    data = Risk_2,
#    index = ['Risk ID', 'Risk', 'Final Residual Risk Rating', 'Prior Residual Rating', 'Overall Assessment Rationale', 'Assessed On'],
#    # values='Days to Log',
#    values = 'Risk ID',
#    aggfunc='count'
#).reset_index().replace('dummy',np.nan)

# Pivot_Risk_2.to_excel("Pivot_Risk_2.xlsx")

####################### Writing to Excel #######################

# create a excel writer object
with pd.ExcelWriter("1022_Pivot Template V001.xlsx") as writer:
    # use to_excel function and specify the sheet_name and index
    # to store the dataframe in specified sheet
    Pivot_Issues_1.to_excel(writer, sheet_name='Pivot - Issues', index=False)
    Pivot_Events_1.to_excel(writer, sheet_name='Pivot - Events', index=False)
    Pivot_Metrics_1.to_excel(writer, sheet_name='Pivot - Overdue Metrics', index=False)
    Pivot_Metrics_2.to_excel(writer, sheet_name='Pivot - Metric Exceptions', index=False)
    Pivot_Risk_1.to_excel(writer, sheet_name='Pivot - Country RAs', index=False)
    Pivot_Risk_2.to_excel(writer, sheet_name='Pivot - Group RAs', index=False)
    Risk_Register.to_excel(writer, sheet_name='Risk Ratings', index=False)
    Issues.to_excel(writer, sheet_name='Issues', index=False)
    Events.to_excel(writer, sheet_name='Events', index=False)
    Metrics.to_excel(writer, sheet_name='Metrics', index=False)


####################### Code run time  #########################

# record end time
end = time.time()
 
# print the difference between start
# and end time in milli. secs
print("The time of execution of above program is :",
      (end-start) * 10**3, "ms")
print("The time of execution of above program is :",
      (((end-start) * 10**3)/1000)/60, "minutes")