# -*- coding: utf-8 -*-
"""
Created on Fri Aug  5 17:37:54 2022

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
from datetime import date
import numpy as np
# import the builtin time module
import time


# Grab Currrent Time Before Running the Code
start = time.time()

os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Supervisory Attestation")

# Fetching the Raw data
df1 = pd.read_excel('supdashboard_supervisoryattestaion_2023-01-05 06_13_56.956.xlsx', sheet_name="supdashboard_supervisoryattesta")
# Data length
print(df1.shape)
# Identifying columns
df1.columns
# filtering the non- attested supervisors
# df2 = df1[df1['Overall Attestation Status'] == "Pending Attestation"]
df2 = df1[df1['Overall Attestation Status'].str.contains('Pending|Pending Attestation', na=False, case=False)]
# Choosing only the needed columns
df3 = df2[['PSID','User Name','Attestation Period','Overall Attestation Status','Country']]
# Remove Duplicates
df3 = df3.drop_duplicates()
# Data length
print(df3.shape)
# Drop Index
df3.reset_index(drop=True, inplace=True)
# Identifying columns
# df3.columns
# Checking the data
# df3.head()

#################### Round 2 Emailing ###############

unique_Bank_Ids = df3['PSID'].astype(int).unique()
unique_Bank_Ids
# unique_Bank_Ids.unique()
# unique_Bank_Ids = pd.DataFrame(unique_Bank_Ids)
# unique_Bank_Ids = unique_Bank_Ids.reset_index()
# unique_Bank_Ids = unique_Bank_Ids[['PSID']]

# for each in unique_Bank_Ids:
#     print (each)
    
# Getting Outlook details:
    
import win32com.client

# Get the User, Manager & Boss ids for all users

User_List = []
Manager_List = []

# To Delete Bank Ids which are invalid
# https://www.statology.org/numpy-remove-element-from-array/
len(unique_Bank_Ids)
# unique_Bank_Ids = np.delete(unique_Bank_Ids, np.where(unique_Bank_Ids == 1370665))
# unique_Bank_Ids = np.setdiff1d(unique_Bank_Ids, [1370665, 1412807, 1332116])
# unique_Bank_Ids = np.setdiff1d(unique_Bank_Ids, [1370665])
len(unique_Bank_Ids)

for i in range(len(unique_Bank_Ids)):
    print (unique_Bank_Ids[i])
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    gal = outlook.Session.GetGlobalAddressList()
    entries = gal.AddressEntries
    recipient = outlook.CreateRecipient(unique_Bank_Ids[i])
    user = recipient.AddressEntry.GetExchangeUser()
    user_manager = recipient.AddressEntry.GetExchangeUser().GetExchangeUserManager()
    # Managers_Manager = user_manager.GetExchangeUserManager()
    # Super_Boss_1 = Managers_Manager.GetExchangeUserManager()
    if user is not None:
        print("User Name:", user.Name)
        # print("User Office Location:",user.OfficeLocation)
        # print("User Department:",user.Department)
        print("Manager's Name:", user_manager.Name)
        # print("Managers_Manager:", Managers_Manager.Name)
        # print("Super_Boss_1:", Super_Boss_1.Name)
        #print("User First Name:",user.FirstName)
        #print("User Last Name:",user.LastName)
        #print("User Mail Id:",user.PrimarySmtpAddress)
        #print("User Job Title:",user.JobTitle)
        #print("User BusinessTelephoneNumber:",user.BusinessTelephoneNumber)
        #print("User MobileTelephoneNumber:",user.MobileTelephoneNumber)
        # User_Name = user.Name
        # User_Department = user.Department
        # Manager_Name = user_manager.Name
        user_id = user.Alias
        Manager_id = user_manager.Alias
        # Boss_id = Managers_Manager.Alias
        User_List.append(user_id) 
        Manager_List.append(Manager_id)


# Final mail generation - R1

for i in range(len(User_List)):
    # print (User_List[i])
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    gal = outlook.Session.GetGlobalAddressList()
    entries = gal.AddressEntries
    recipient = outlook.CreateRecipient(User_List[i])
    user = recipient.AddressEntry.GetExchangeUser()
    user_manager = recipient.AddressEntry.GetExchangeUser().GetExchangeUserManager()
    Managers_Manager = user_manager.GetExchangeUserManager()
    # Super_Boss_1 = Managers_Manager.GetExchangeUserManager()
    if user is not None:
        # print("User Name:", user.Name)
        # print("User Office Location:",user.OfficeLocation)
        # print("User Department:",user.Department)
        # print("Manager's Name:", user_manager.Name)
        # print("Managers_Manager:", Managers_Manager.Name)
        # print("Super_Boss_1:", Super_Boss_1.Name)
        #print("User First Name:",user.FirstName)
        #print("User Last Name:",user.LastName)
        #print("User Mail Id:",user.PrimarySmtpAddress)
        #print("User Job Title:",user.JobTitle)
        #print("User BusinessTelephoneNumber:",user.BusinessTelephoneNumber)
        #print("User MobileTelephoneNumber:",user.MobileTelephoneNumber)
        # User_Name = user.Name
        # User_Department = user.Department
        # Manager_Name = user_manager.Name
        # user_id = user.Alias
        # Manager_id = user_manager.Alias
        # Boss_id = Managers_Manager.Alias
        # User_First_Name = user.FirstName
        # User_Last_Name = user.LastName
        # User_Mail_Id = user.PrimarySmtpAddress
        # User_Job_Title = user.JobTitle
        # User_Office_Location = user.OfficeLocation
        # User_BusinessTelephoneNumber = user.BusinessTelephoneNumber
        # User_MobileTelephoneNumber = user.MobileTelephoneNumber
    # Framing data set unique to the Business Issue owner
    # df4 = df3.loc[df1['PSID'] == User_List[i]]
    # Remove Duplicates
    # df4 = df4.drop_duplicates()
    # Reset Index
    # df4.reset_index(drop=True, inplace=True)
    # Id = df4['PSID'].iloc[0]
#    Defect_count = df2['Defect_count'].iloc[0]
#    Date = df2['Date'].iloc[0]
    # Converts date to the needed format
    # Date = Date.strftime('%Y-%m-%d')
    # today = date.today().strftime('%Y-%m-%d')
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
    # Including Transaction date in the subject to prevent overwriting of the mail while saving
    # subject = 'AHA EOD email summary ' + " " + "-" + " "+ str(Id) + ''
        subject = 'FM Supervisory Attestation - Q* 2022 [INTERNAL] - R2'
        # Users
        mail.To = ";".join(str(u)  for u in User_List)
        # mail.To = str(df2['PSID'].iloc[0])
        mail.Subject = subject
        # Recipients for cc
        #recipients = ['Dinesh.Charan@sc.com', 'Xavier.Dionette-Marie@sc.com']
        # recipients = ['1510806']
        # recipients = ['1580298; 1386073; 1226952; 1510806']
        # mail.CC = "; ".join(recipients)
        mail.CC = ";".join(str(m)  for m in Manager_List)
        mail.Sensitivity = 3
        result = df3.to_html()
        mail.HTMLBody = """
        <html>
        <body>
        Dear All,<br></br>
        <br></br>Good Day.<br></br>
        <br></br>The Supervisory Attestation for Q4 2022 is ready on the dashboard, please complete the attestation by Tuesday, 20 Feb 2023 to avoid any overdue cases, thank you.<br></br>
        •<a href="https://supdashboard.gdc.standardchartered.com/#/home"> https://supdashboard.gdc.standardchartered.com/#/home </a><br></br>
        • Supervisory Dashboard > Pending Attestation > Supervisory Attestation <br></br><br></br>
     
<!DOCTYPE html>
 <html>
  <head>
    <title>Title of the document</title>
    <style>
      table,
      th,
      td {
        padding: 10px;
        border: 1px solid black;
        border-collapse: collapse;
      }
    </style>
  </head>
        <th style="text-align:center" > """    + result +     """  </th>
</html>
        <p><b>Please provide your updates by 28 Feb 2023. Please ignore if responded already.</b></p>
        <br>Thanks & Regards,</br>
        <br>CCG BMG</br>
        </body>
        </html>
        """
    name = str("R2")
    # print("HTML & Name step completed - Check")
    #mail.Categories='Red category'
    #mail.Save()
    #mail.Send()
    #mail.Move(Mails_To_Send_AHA_EOD)
    # Declaring directory to store draft mails
mail.SaveAs('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//SUP_Mails'+'//'+name+'.msg')

# Grab Currrent Time After Running the Code
end = time.time()
#Subtract Start Time from The End Time
total_time = end - start
print("\n"+ "Run Time: " +  str(total_time/60) + " mins")