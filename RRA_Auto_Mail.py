# -*- coding: utf-8 -*-
"""
Created on Thu Aug 11 14:19:00 2022

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

#directory = os.chdir("C:/Users/1510806/OneDrive - Standard Chartered Bank/Desktop/RRA_Test/")
#
#from datetime import date
#import numpy as np
#
#df1 = pd.read_excel('BJ Exceptions Aug.xlsx')
##df1 = pd.read_excel('BJ report July.xlsx', index_col=None, usecols=None)
#
## Filter needed columns
#df1 = df1[["meetingId", "meetingTitle", "userName", "endTime", "startTime", "Converted Time", "email", "Bank Id", "Staff Name", "Staff Country"]]
#
#df1 = pd.DataFrame(df1)
#
## Renaming columns in Excel for running the script
#dict = {'Converted Time':'Converted_Time', 'Bank Id': 'Bank_Id', 'Staff Name':'Staff_Name', 'Staff Country': 'Staff_Country'}
#    
#df1.rename(columns=dict,  inplace=True)
#
#unique_BO = df1['Bank_Id'].unique()

#for each in unique_BO:
#    print (str(each))

#unique_BO = df1['Bank_Id'].unique()
#df1['Bank_Id'].dropna()
#df1['Bank_Id'].astype(int)

## auto fit column width
#def get_col_widths(dataframe):
#    # First we find the maximum length of the index column   
#    idx_max = max([len(str(s)) for s in dataframe.index.values] + [len(str(dataframe.index.name))])
#    # Then, we concatenate this to the max of the lengths of column name and its values for each column, left to right
#    return [idx_max] + [max([len(str(s)) for s in dataframe[col].values] + [len(col)]) for col in dataframe.columns]

directory = os.chdir("C:/Users/1510806/OneDrive - Standard Chartered Bank/Desktop/RRA_Mail_Test/")

from datetime import date
import numpy as np

for each in os.listdir(directory):
    print (str(each))

# To access cardnumbers one by one - include in for loop
for i in os.listdir(directory):
    # Framing data set unique to the Business Issue owner
#    df2 = df1.loc[df1['Bank_Id'] == unique_BO[i]]
#    Bank_Id = df2['Bank_Id'].iloc[0]
#    meetingId = df2['meetingId'].iloc[0]
#    meetingTitle  = df2['meetingTitle'].iloc[0]
#    userName  = df2['userName'].iloc[0]
#    endTime  = df2['endTime'].iloc[0]
#    startTime  = df2['startTime'].iloc[0]
#    Converted_Time  = df2['Converted_Time'].iloc[0]
#    email  = df2['email'].iloc[0]
#    Bank_Id  = df2['Bank_Id'].iloc[0]
#    Staff_Name  = df2['Staff_Name'].iloc[0]
#    Staff_Country  = df2['Staff_Country'].iloc[0]
    #today = date.today().strftime("%d-%m-%Y")
    today = date.today().strftime("%b %Y")
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    # BO Data to be attached
#    data_attach = df2.loc[df2['Bank_Id'] == unique_BO[i]]
#    # Reset Index
#    data_attach.reset_index(drop=True, inplace=True)
#    # Check Shape of data for the BO
#    print(data_attach.shape)

    # Dynamic name change as per BO and excel creation fro attachment
#    data_attach.to_excel(r'C:/Users/1510806/OneDrive - Standard Chartered Bank/Desktop/CCG/AHA/BJ_Auto_Mailer/RRA_Mail_Test/' + ' BJ Tracking List ' + "- " + "Date as of - " + str(dt.datetime.today().strftime("%d%m%Y")) + "  " + "-" + "  "+ "User Id - " + str(Bank_Id) + "  " + "-" + "  " +  str(Staff_Name) + '.xlsx')
#    
    # Adding the Data for each BO
    
    # mail.Attachments.Add(i)
    mail.Attachments.Add(r'C:/Users/1510806/OneDrive - Standard Chartered Bank/Desktop/RRA_Mail_Test/' +  i)
    # Including Transaction date in the subject to prevent overwriting of the mail while saving
    star = "**"
    subject =  'RRA East ' + i[:-5]
    #subject =  'Bluejeans Recording Check - Aug - 2022 - ' + ' User' + "  " + "-" + "  " +  str(Bank_Id) + "  " + "-" + "  " +  str(Staff_Name)
    # mail.To = str(df2['Bank_Id'].iloc[0])
    recipients = ['Harinis.Prabu@sc.com; Dinesh.Charan@sc.com; Mohamed-Yousuf.Sarabudeen@sc.com']
    mail.To = "; ".join(recipients)
    mail.Subject = subject
    # Recipients for cc
    #recipients = ['Dinesh.Charan@sc.com', 'Xavier.Dionette-Marie@sc.com']
    recipients = ['Harinis.Prabu@sc.com; Dinesh.Charan@sc.com; Mohamed-Yousuf.Sarabudeen@sc.com']
    mail.CC = "; ".join(recipients)
    mail.Sensitivity = 3
    #df2 = df2.reset_index(drop=True)
    #result = df2.to_html()
    mail.HTMLBody = """
        <html>
        <body>
        Good Day,
        <br></br>
        <br>Hope you are doing great. </br>
        <br></br>
        <br>Please find the details of RRA East files attached here. </br>
        <br></br>
        <br>Please find the details of RRA East files attached here. </br>
        <u style="text-decoration-color:red">
        <p style="color:red;"> <b> Please find the details of RRA East files attached here.   </b></p>
        </u>
        Please find the details of RRA East files attached here. Thanks.
      <br>
     
        <p><b>Please provide your updates by 15 Dec 2022.</b></p>
        <br>Thanks & Regards,</br>
        <br>CCG BMG</br>
        </body>
        </html>
        """
        # mail.Display()
    name = str(subject)
    # Declaring directory to store draft mails
    mail.SaveAs('C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//RRA_Mail_Test//'+'//'+name+'.msg')
    