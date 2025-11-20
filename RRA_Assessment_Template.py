# -*- coding: utf-8 -*-
"""
Created on Thu Dec 22 21:33:23 2022

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


# Read excel file with sheet name

xls = pd.ExcelFile('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Py_Data_RRA_Automation//1022_Pivot Template V001.xlsx')
Pivot_Issues = pd.read_excel(xls, 'Pivot - Issues')
Pivot_Events = pd.read_excel(xls, 'Pivot - Events')
Pivot_Overdue_Metrics = pd.read_excel(xls, 'Pivot - Overdue Metrics')
Pivot_Metric_Exceptions = pd.read_excel(xls, 'Pivot - Metric Exceptions')
Pivot_Country_RAs = pd.read_excel(xls, 'Pivot - Country RAs')
Pivot_Group_RAs = pd.read_excel(xls, 'Pivot - Group RAs')
#Risk_Ratings = pd.read_excel(xls, 'Risk Ratings')
#Issues = pd.read_excel(xls, 'Issues')
#Events = pd.read_excel(xls, 'Events')
#Metrics = pd.read_excel(xls, 'Metrics')

############################# Filter for Country HK ##########################

Country = 'HONG KONG'
Pivot_Issues = Pivot_Issues[Pivot_Issues['Owner Organization - Country / Group'] == Country]
Pivot_Events = Pivot_Events[Pivot_Events['Country'] == Country]
Pivot_Overdue_Metrics = Pivot_Overdue_Metrics[Pivot_Overdue_Metrics['Monitored for/Tested Organization Country'] == Country]
Pivot_Metric_Exceptions = Pivot_Metric_Exceptions[Pivot_Metric_Exceptions['Monitored for/Tested Organization Country'] == Country]
Pivot_Country_RAs = Pivot_Country_RAs[Pivot_Country_RAs['Overall Assessment Rationale'] == Country]
Pivot_Group_RAs = Pivot_Group_RAs[Pivot_Group_RAs['Overall Assessment Rationale'] == Country]
#Risk_Ratings = Risk_Ratings[Risk_Ratings['Overall Assessment Rationale'] == Country]
#Issues = Issues[Issues['Owner Organization - Country / Group'] == Country]
#Events = Events[Events['Country'] == Country]
#Metrics = Metrics[Metrics['Monitored for/Tested Organization Country'] == Country]

####################### Heatmap #######################

Heatmap = pd.read_excel('Heatmap.xlsx')
Heatmap = pd.DataFrame(Heatmap)
print(Heatmap.shape)

# Count the number of Occurrence of Values based on another column
Pivot_Issues['Risk ID'] = Pivot_Issues['Risk ID'].str.upper()
Pivot_Issues['Risk ID'] = Pivot_Issues['Risk ID'].str.strip()
Heatmap['Risk ID'] = Heatmap['Risk ID'].str.upper()
Heatmap['Risk ID'] = Heatmap['Risk ID'].str.strip()
Country_Unique_Risk_Id = Pivot_Issues['Risk ID'].value_counts().reset_index(name='Unique_Country_Specific_Risk_ID_Count')
Country_Unique_Risk_Id.rename(columns = {'index':'Risk ID'}, inplace = True)
Heatmap_Risk_ID = Heatmap['Risk ID'].to_frame()
Heatmap_Risk_ID_vlookup = pd.merge(Heatmap_Risk_ID, Country_Unique_Risk_Id, on ='Risk ID', how ='left')
Heatmap['Issues'] = Heatmap_Risk_ID_vlookup['Unique_Country_Specific_Risk_ID_Count']

####################### Writing to Excel #######################

# create a excel writer object
with pd.ExcelWriter("Assessment_Template_HK.xlsx") as writer:
    # use to_excel function and specify the sheet_name and index
    # to store the dataframe in specified sheet
    Heatmap.to_excel(writer, sheet_name='Heatmap', index=False)
    Pivot_Issues.to_excel(writer, sheet_name='Issues', index=False)
    Pivot_Events.to_excel(writer, sheet_name='Events', index=False)
    Pivot_Overdue_Metrics.to_excel(writer, sheet_name='Overdue Metrics', index=False)
    Pivot_Metric_Exceptions.to_excel(writer, sheet_name='Metric Exceptions', index=False)
    Pivot_Country_RAs.to_excel(writer, sheet_name='HK Ratings', index=False)
    Pivot_Group_RAs.to_excel(writer, sheet_name='Group Ratings', index=False)


####################### Code run time  #########################

# record end time
end = time.time()
 
# print the difference between start
# and end time in milli. secs
print("The time of execution of above program is :",
      (end-start) * 10**3, "ms")
print("The time of execution of above program is :",
      (((end-start) * 10**3)/1000)/60, "minutes")