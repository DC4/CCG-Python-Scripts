###############################################################################################################################
##################################################       BJ       #############################################################
###############################################################################################################################


# -*- coding: utf-8 -*-
"""
Created on Thu Aug 25 16:56:25 2022

@author: 1659765
"""
import pandas as pd
import numpy as np
import datetime
import os
from datetime import date

def previous_quarter(ref):
    if ref.month < 4:
        return datetime.date(ref.year - 2, 12, 31)
    elif ref.month < 7:
        return datetime.date(ref.year-1, 3, 31)
    elif ref.month < 10:
        return datetime.date(ref.year-1, 6, 30)
    return datetime.date(ref.year-1, 9, 30)

cut_offdate=previous_quarter(date.today())
os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct_Py_Scripts//BJ")
BJ_rawdata=pd.ExcelFile("BJ Exceptions - FM and TM.xlsx")
BJ=BJ_rawdata.parse('BJ Data')


BJ["Type of Breach"]="Market Abuse (BlueJeans Recording)"
BJ["Severity"]=""
BJ["Accountability"]="NA"
BJ["Comments"]=""
BJ["Q_Y"]=BJ["Quarter"].astype(str)+" "+BJ["Year"].astype(str)
BJ=BJ[["Bank ID","Type of Breach","Severity","Accountability","Q_Y","Comments","Defect Count","Product Desk"]]
BJ=BJ[np.invert(BJ["Product Desk"].astype(str).str.contains("Treasury"))]

BJ=BJ[["Bank ID","Type of Breach","Severity","Accountability","Q_Y","Comments","Defect Count"]]
BJ=BJ.groupby(["Bank ID","Type of Breach","Severity","Accountability","Comments","Q_Y"]).sum().reset_index()
BJ["Comments"]=BJ["Defect Count"].astype(str)+" breach(es) in a quarter"
BJ=BJ[["Bank ID","Type of Breach","Severity","Accountability","Q_Y","Comments"]]
BJ = BJ.set_axis(axis=1,labels=["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment"])

BJ["Date"]=np.where(BJ["Date"].str[:2]=="Q1","31 March "+BJ["Date"].str[-4:],BJ["Date"])
BJ["Date"]=np.where(BJ["Date"].str[:2]=="Q2","30 June "+BJ["Date"].str[-4:],BJ["Date"])
BJ["Date"]=np.where(BJ["Date"].str[:2]=="Q3","30 September "+BJ["Date"].str[-4:],BJ["Date"])
BJ["Date"]=np.where(BJ["Date"].str[:2]=="Q4","31 December "+BJ["Date"].str[-4:],BJ["Date"])


BJ=BJ[["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment"]]
BJ.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\BJ\BJ_Output.xlsx",index=False)
BJ.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Final\BJ_Output.xlsx",index=False)



###############################################################################################################################
##################################################  Data Breach  ##############################################################
###############################################################################################################################




# -*- coding: utf-8 -*-
"""
Created on Tue Aug 16 18:01:58 2022

@author: 1659765
"""

import pandas as pd
import numpy as np

#%%As the sender is different every quarter the format will change too
#==============================================================================
# data_breach=pd.read_excel("CCIB FM Q3 2021 - Q2 2022_Contacts.xlsx")
# data_breach_name=data_breach[["Incident ID","Involved Staff ID","Involved Staff Name"]]
# data_breach_name=data_breach_name.dropna()
# data_breach_comment=pd.read_excel("CCIB FM Q3 2021 - Q2 2022_Reporter's summary.xlsx",skiprows=2)
# data_breach_comment=data_breach_comment.drop(['Incident Status'],axis=1)
# data_breach=data_breach.drop(["Involved Staff ID","Involved Staff Name"],axis=1)
# data_breach=data_breach.fillna(method='ffill')
# data_breach=data_breach.join(data_breach_name.set_index('Incident ID'),on='Incident ID')
# data_breach=data_breach[data_breach['Involved Staff ID'].notnull()]
# data_breach=data_breach.join(data_breach_comment.set_index('Incident ID'),on='Incident ID')
# data_breach["Type of Breach"]="Data Breach"
# data_breach["Accountability"]="NA"
# data_breach=data_breach[["Involved Staff ID","Type of Breach","Business Risk Rating by BRM","Accountability","Month","Reporter's Summary"]]
# data_breach.set_axis(axis=1,labels=["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment"])
# 
# data_breach["Severity"]=data_breach["Severity"].str.split('(',expand=True)[1].str.split(')',expand=True)[0].str.strip()
# data_breach["Employee PSID"]=data_breach["Employee PSID"].astype(int).astype(str)
# 
# 
#==============================================================================
os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct_Py_Scripts//Data Breach")
data_breach=pd.read_excel("Data Breach.xlsx")
data_breach["Severity"]=data_breach["Severity"].str.split('(',expand=True)[1].str.split(')',expand=True)[0].str.strip()
data_breach["Employee PSID"]=data_breach["Employee PSID"].astype(str)
data_breach.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Data Breach\Data_Breach_Output.xlsx",index=False)
data_breach.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Final\Data_Breach_Output.xlsx",index=False)


###############################################################################################################################
##################################################  Data Leakage  ##############################################################
###############################################################################################################################

# -*- coding: utf-8 -*-
"""
Created on Tue Aug 16 18:01:58 2022

@author: 1659765
"""

import pandas as pd
import numpy as np
import datetime
import os
from datetime import date

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

os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct_Py_Scripts//Data Leakage")
data_leakage=pd.read_excel("MI Reports by Business Function L2 - New HR Format.xlsx")
data_leakage["Type of Breach"]="Data Leakage"
data_leakage["Accountability"]="NA"
data_leakage=data_leakage[["Case ID","UserID","Type of Breach","Grade","Accountability","Closed Date","Recommendation"]]
data_leakage.drop_duplicates(inplace=True)
data_leakage=data_leakage.drop(['Case ID'],axis=1)
data_leakage = data_leakage.set_axis(axis=1,labels=["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment"])
data_leakage["Date"]=pd.to_datetime(data_leakage["Date"],format=r'%m/%d/%Y %I:%M:%S %p')

data_leakage=data_leakage[(data_leakage["Date"]>cut_offdate)]
data_leakage=data_leakage[data_leakage['Date'].notnull()]
data_leakage=data_leakage[np.invert(data_leakage["Severity"]=='Non-Event')]
data_leakage["Employee PSID"]=data_leakage["Employee PSID"].astype(int).astype(str)

data_leakage.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Data Leakage\Data_Leakage_Output.xlsx",index=False)
data_leakage.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Final\Data_Leakage_Output.xlsx",index=False)

###############################################################################################################################
######################################################  EOD  ##################################################################
###############################################################################################################################

# -*- coding: utf-8 -*-
"""
Created on Thu Aug 25 16:56:25 2022

@author: 1659765
"""
import pandas as pd
import numpy as np
import datetime
import os
from datetime import date

os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct_Py_Scripts//EOD")

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


EOD_rawdata=pd.ExcelFile("EOD Exceptions - Consolidated East and West Results.xlsx")
EOD_East=EOD_rawdata.parse('EAST EOD Recap')
EOD_East=EOD_East.loc[EOD_East.index.repeat(EOD_East["Defect count"])]
EOD_East["Type of Breach"]="Market Abuse (EOD Update)"
EOD_East["Severity"]=""
EOD_East["Accountability"]="NA"
EOD_East=EOD_East[["Bank ID","Type of Breach","Severity","Accountability","Period 2","EOD email summary","Product Desk"]]
EOD_East = EOD_East.set_axis(axis=1,labels=["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment","Product Desk"])


EOD_West=EOD_rawdata.parse('WEST EOD Recap')
EOD_West["Q_Y"]=EOD_West["Quarter"].astype(str)+EOD_West["Year"].astype(str)
EOD_West=EOD_West.loc[EOD_West.index.repeat(EOD_West["Defect Count"])]
EOD_West["Type of Breach"]="Market Abuse (EOD Update)"
EOD_West["Severity"]=""
EOD_West["Accountability"]="NA"
EOD_West=EOD_West[["Bank ID","Type of Breach","Severity","Accountability","Period 2","Comments","Product Desk"]]
EOD_West = EOD_West.set_axis(axis=1,labels=["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment","Product Desk"])

EOD=pd.concat([EOD_West,EOD_East])

EOD=EOD[(EOD["Date"]>cut_offdate)]

EOD=EOD.sort_values(by=["Employee PSID","Date"]).reset_index(drop=True)
EOD["Cumulative Count"]=EOD.groupby("Employee PSID").cumcount()+1

EOD_reverse=EOD.sort_values(by=["Employee PSID","Date"],ascending=[True,False]).reset_index(drop=True)
# EOD["Time Diff"]=EOD["Date"].dt.to_period('M')-EOD_reverse["Date"].dt.to_period('M')
EOD["Time Diff"]=EOD["Date"].dt.to_period('M').astype(int)-EOD_reverse["Date"].dt.to_period('M').astype(int)
EOD["Severity"]=np.where(np.logical_and(EOD["Time Diff"]<=6,EOD["Cumulative Count"]>=5),"5 or more period 6 mths","")
EOD["Severity"]=np.where(EOD["Cumulative Count"]>=7,"7 or more period 12 mths",EOD["Severity"])

EOD=EOD[np.invert(EOD["Product Desk"].str.contains("Treasury"))]
EOD=EOD[np.invert(EOD["Product Desk"].str.contains("TM"))]

EOD=EOD[["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment"]]

EOD.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\EOD\EOD_Output.xlsx",index=False)
EOD.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Final\EOD_Output.xlsx",index=False)

###############################################################################################################################
##################################################     FLM      ###############################################################
###############################################################################################################################

# -*- coding: utf-8 -*-
"""
Created on Wed Sep 28 12:45:05 2022

@author: 1659765
"""

import pandas as pd
import numpy as np

os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct_Py_Scripts//FLM")
FLM=pd.read_excel("FLM.xlsx")

FLM["Employee PSID"]=FLM["Employee PSID"].astype(str)

FLM.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\FLM\FLM_Output.xlsx",index=False)
FLM.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Final\FLM_Output.xlsx",index=False)

###############################################################################################################################
#####################################################    OPS    ###############################################################
###############################################################################################################################

# -*- coding: utf-8 -*-
"""
Created on Wed Sep 28 12:45:05 2022

@author: 1659765
"""

import pandas as pd
import numpy as np

os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct_Py_Scripts//Operation Losses")
Ops=pd.read_excel("Ops.xlsx")

Ops["Employee PSID"]=Ops["Employee PSID"].astype(str)

Ops.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Operation Losses\Ops_Output.xlsx",index=False)
Ops.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Final\Ops_Output.xlsx",index=False)

###############################################################################################################################
#####################################################    Others    ############################################################
###############################################################################################################################

# -*- coding: utf-8 -*-
"""
Created on Tue Oct 25 14:26:34 2022

@author: 1659765
"""

import pandas as pd
import numpy as np

os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct_Py_Scripts//Others")
Others=pd.read_excel("Others.xlsx")

Others["Employee PSID"]=Others["Employee PSID"].astype(str)

Others.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Others\Others_Output.xlsx",index=False)
Others.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Final\Others_Output.xlsx",index=False)

###############################################################################################################################
#######################################################     PAD   #############################################################
###############################################################################################################################

# -*- coding: utf-8 -*-
"""
Created on Tue Aug 16 18:01:58 2022

@author: 1659765
"""

import pandas as pd
import numpy as np
import datetime
import os
from datetime import date

os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct_Py_Scripts//PAD")

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

PAD=pd.read_excel("CCIB FM Consolidated Monthly Breaches Reports - New.xls")
PAD_main=pd.read_excel("PAD_MainData.xlsx")

PAD["Type of Breach"]="PAD breaches"
PAD["Accountability"]="Not Available"
PAD["Risk Classification"]=np.where(PAD["Risk Classification"]=="Very Low","Very Low (1)",PAD["Risk Classification"])
PAD["Risk Classification"]=np.where(PAD["Risk Classification"]=="Low","Low (2)",PAD["Risk Classification"])
PAD["Risk Classification"]=np.where(PAD["Risk Classification"]=="Medium","Medium (3)",PAD["Risk Classification"])
PAD["Risk Classification"]=np.where(PAD["Risk Classification"]=="High","High (4)",PAD["Risk Classification"])

PAD["Remarks"]=PAD["PAD Breaches"]+ " "+PAD["Remarks"]
PAD=PAD[["Bank ID","Type of Breach","Risk Classification","Accountability","Breach Issued on","Remarks","Transaction ID"]]
PAD.set_axis(axis=1,labels=["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment","Transaction ID"])
PAD_main=pd.concat([PAD_main, PAD], ignore_index=True)
PAD_main=PAD_main.drop_duplicates()
PAD_main=PAD_main[PAD_main["Date"]>cut_offdate]

PAD=PAD_main[["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment"]]
PAD_main.to_excel("PAD_MainData.xlsx",index=False)

PAD_main.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\PAD\PAD_Output.xlsx",index=False)
PAD_main.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Final\PAD_Output.xlsx",index=False)

PAD.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\PAD\PAD.xlsx",index=False)
PAD.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Final\PAD.xlsx",index=False)


###############################################################################################################################
#####################################################   TCIW   ################################################################
###############################################################################################################################


# -*- coding: utf-8 -*-
"""
Created on Tue Aug 16 18:01:58 2022

@author: 1659765
"""

import pandas as pd
import numpy as np
import datetime
import os
from datetime import date

os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct_Py_Scripts//TCIW")

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

TCIW=pd.read_excel("CCIB FM Consolidated Monthly Breaches Reports - New.xls")
TCIW_main=pd.read_excel("TCIW_MainData.xlsx")
    
TCIW["Type of Breach"]="TCIW"
TCIW["Accountability"]="Not Available"
TCIW["Final Severity Rating"]=np.where(TCIW["Final Severity Rating"]=="Very Low","Very Low (1)",TCIW["Final Severity Rating"])
TCIW["Final Severity Rating"]=np.where(TCIW["Final Severity Rating"]=="Low","Low (2)",TCIW["Final Severity Rating"])
TCIW["Final Severity Rating"]=np.where(TCIW["Final Severity Rating"]=="Medium","Medium (3)",TCIW["Final Severity Rating"])
TCIW["Final Severity Rating"]=np.where(TCIW["Final Severity Rating"]=="High","High (4)",TCIW["Final Severity Rating"])

TCIW=TCIW[["Staff ID","Type of Breach","Final Severity Rating","Accountability","Month of Close Date","Breach Description"]]
TCIW.set_axis(axis=1,labels=["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment"])
TCIW_main=pd.concat([TCIW_main, TCIW], ignore_index=True)
TCIW_main=TCIW_main.drop_duplicates()
TCIW=TCIW_main[TCIW_main["Date"]>cut_offdate]

TCIW["Employee PSID"]=TCIW["Employee PSID"].astype(int).astype(str)
TCIW_main.to_excel("TCIW_MainData.xlsx",index=False)

TCIW_main.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\TCIW\TCIW_Output.xlsx",index=False)
TCIW_main.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Final\TCIW_Output.xlsx",index=False)

TCIW.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\TCIW\TCIW.xlsx",index=False)
TCIW.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Final\TCIW.xlsx",index=False)


###############################################################################################################################
##################################################      FINAL    ##############################################################
###############################################################################################################################

# -*- coding: utf-8 -*-
"""
Created on Mon Sep 26 13:48:50 2022

@author: 1659765
"""

#Import library
import pandas as pd
import numpy as np

os.chdir("C://Users//1510806//OneDrive - Standard Chartered Bank//Desktop//Conduct_Py_Scripts//Final")

def quarter_list_gen(ref):
    if ref.month < 4:
        return ["Total","Q1 "+str(ref.year-1),"Q2 "+str(ref.year-1),"Q3 "+str(ref.year-1),"Q4 "+str(ref.year-1)]
    elif ref.month < 7:
        return ["Total","Q2 "+str(ref.year-1),"Q3 "+str(ref.year-1),"Q4 "+str(ref.year-1),"Q1 "+str(ref.year)]
    elif ref.month < 10:
        return ["Total","Q3 "+str(ref.year-1),"Q4 "+str(ref.year-1),"Q1 "+str(ref.year),"Q2 "+str(ref.year)]
    return ["Total","Q4 "+str(ref.year-1),"Q1 "+str(ref.year),"Q2 "+str(ref.year),"Q3 "+str(ref.year)]

#Change this every quarter
Quarter_list=quarter_list_gen(date.today())

#Import raw data from FMSW
FMSW_rawdata=pd.ExcelFile("FMSWData_12.xlsx")

#Remove breaches that are not reported in forum, to review in every cycle
list_of_headers=FMSW_rawdata.sheet_names[1:]
list_to_remove=['Transaction Rate Review','Dodd Frank','Gifts & Entertainment']
for i in list_to_remove:
    list_of_headers.remove(i)

#%% Consolidate all FMSW tab into 1 file
final=pd.DataFrame(columns=["Employee PSID - Name","Type of Breach","Severity","Accountability","Date","Comment","Workflow Status","Attestation Date"])
for i in list_of_headers:
        
    #Obtaining Employee PSID - Name column
    if i=="VC - Mismark":
        column_2=FMSW_rawdata.parse(i).iloc[:,14].rename("Employee PSID - Name").to_frame()
    else:
        column_2=FMSW_rawdata.parse(i).iloc[:,0].rename("Employee PSID - Name").to_frame()
    
    #Obtaining Type of Breach column
    column_3=[]
    for j in range(len(column_2)):
        column_3.append(i)
    column_3=pd.Series(column_3).rename("Type of Breach").to_frame()
    
    #Obtaining Severity column
    if "Severity" in list(FMSW_rawdata.parse(i)):
        column_4=FMSW_rawdata.parse(i)["Severity"]
    elif "SEVERITY" in list(FMSW_rawdata.parse(i)):
        column_4=FMSW_rawdata.parse(i)["SEVERITY"].rename("Severity")
    elif "MI RAG Status" in list(FMSW_rawdata.parse(i)):
        column_4=FMSW_rawdata.parse(i)["MI RAG Status"].rename("Severity")
    elif "Rag Status" in list(FMSW_rawdata.parse(i)):
        column_4=FMSW_rawdata.parse(i)["Rag Status"].rename("Severity")
    elif "Breach Category" in list(FMSW_rawdata.parse(i)):
        column_4=FMSW_rawdata.parse(i)["Breach Category"].rename("Severity")
    elif "Fair Accountability Severity Rating" in list(FMSW_rawdata.parse(i)):
        column_4=FMSW_rawdata.parse(i)["Fair Accountability Severity Rating"].rename("Severity")
    elif "Issue Category" in list(FMSW_rawdata.parse(i)):
        column_4=FMSW_rawdata.parse(i)["Issue Category"].rename("Severity")
    elif i=="VC - Mismark":
        column_4=[]
        for j in range(len(column_2)):
            if FMSW_rawdata.parse(i)["Mismark"][j]=="Yes":
                column_4.append("RED")
            elif FMSW_rawdata.parse(i)["Potential Mismark"][j]=="Yes" and FMSW_rawdata.parse(i)["Mismark"][j]!="No":
                column_4.append("RED")
            elif FMSW_rawdata.parse(i)["Potential Mismark"][j]=="yes"and FMSW_rawdata.parse(i)["Mismark"][j]!="No":
                column_4.append("RED")
            else:
                column_4.append("NA")
        column_4=pd.Series(column_4).rename("Severity").to_frame()
    elif "Overall Attestation Status" in list(FMSW_rawdata.parse(i)):
        column_4=FMSW_rawdata.parse(i)["Overall Attestation Status"].rename("Severity")
    else:
        column_4=[]
        for j in range(len(column_2)):
            column_4.append("NA")
        column_4=pd.Series(column_4).rename("Severity").to_frame()
        

    #Obtaining Accountability column
    if "Disciplinary Action Taken" in list(FMSW_rawdata.parse(i)):
        column_5=FMSW_rawdata.parse(i)["Disciplinary Action Taken"].rename("Accountability")
    elif "Fair Accountability Outcome" in list(FMSW_rawdata.parse(i)):
        column_5=FMSW_rawdata.parse(i)["Fair Accountability Outcome"].rename("Accountability")
    else:
        column_5=[]
        for j in range(len(column_2)):
            column_5.append("NA")
        column_5=pd.Series(column_5).rename("Accountability").to_frame()
 
    #Obtaining Month column
    if "As Of Date" in list(FMSW_rawdata.parse(i)):
        column_6=FMSW_rawdata.parse(i)["As Of Date"].rename("Date")
    elif "MTCR Review Date" in list(FMSW_rawdata.parse(i)):
        column_6=FMSW_rawdata.parse(i)["MTCR Review Date"].rename("Date")
    elif "Workflow Initiated (UTC)" in list(FMSW_rawdata.parse(i)):
        column_6=FMSW_rawdata.parse(i)["Workflow Initiated (UTC)"].rename("Date")
    elif "Deal Date" in list(FMSW_rawdata.parse(i)):
        column_6=FMSW_rawdata.parse(i)["Deal Date"].rename("Date")
    elif "Review Month" in list(FMSW_rawdata.parse(i)):
        column_6=FMSW_rawdata.parse(i)["Review Month"].rename("Date")
    elif "LAST_MODIFIED_DATE" in list(FMSW_rawdata.parse(i)):
        column_6=FMSW_rawdata.parse(i)["LAST_MODIFIED_DATE"].rename("Date")
    elif "Month-end" in list(FMSW_rawdata.parse(i)):
        column_6=FMSW_rawdata.parse(i)["Month-end"].rename("Date")        
    elif "Month of Report" in list(FMSW_rawdata.parse(i)):
        column_6=FMSW_rawdata.parse(i)["Month of Report"].rename("Date")   
    elif i=="FM Supervisory Attestation Brea":
        column_6=FMSW_rawdata.parse(i)["Attestation Period"].rename("Date") 
    else:
        column_6=[]
        for j in range(len(column_2)):
            column_6.append("NA")
        column_6=pd.Series(column_6).rename("Date").to_frame()
    
    #Obtaining Commentary column
    if i=="Group Mandatory E - Learning Ov":
        column_7=(FMSW_rawdata.parse(i)["E-Learning Name"]+" with ageing of "+ FMSW_rawdata.parse(i)["Ageing"].apply(str)).rename("Comment").to_frame()
    elif i=="Missed Trade":
        column_7=((FMSW_rawdata.parse(i)["Trade Reference"]).astype(str)+(FMSW_rawdata.parse(i)["Staff Responsible Comments"])).rename("Comment").to_frame()
    elif i=="Pre & Post Trade Mid Mark":
        column_7=(FMSW_rawdata.parse(i)["Exception"]).rename("Comment").to_frame()        
    elif i=="Client Categorisation":
        column_7=(FMSW_rawdata.parse(i)["FO Staff Justification"]).rename("Comment").to_frame()    
    elif i=="Trader Mandate Exception":
        column_7=(FMSW_rawdata.parse(i)["Trader Justification Code"]).rename("Comment").to_frame()    
    elif i=="Credit Risk Excess":
        column_7=(FMSW_rawdata.parse(i)["Explanation / Details of Control Break Down"]).rename("Comment").to_frame()   
    elif i=="Termsheet Findings":
        column_7=(FMSW_rawdata.parse(i)["Remarks (Sales' explanations and whether repeat offender)"]).rename("Comment").to_frame()   
    elif i=="VC - Mismark":
        column_7=(FMSW_rawdata.parse(i)["Comments"]+FMSW_rawdata.parse(i)["Curve / Security Description"]).rename("Comment").to_frame()   
    elif i=="Best Execution Alerts":
        column_7=(FMSW_rawdata.parse(i)["Justification Comments"]).rename("Comment").to_frame()   
    elif i=="FM Surveillance Alerts":
        column_7=(FMSW_rawdata.parse(i)["Summary of Issue"]).rename("Comment").to_frame()   
    elif i=="Market Excess":
        column_7=(FMSW_rawdata.parse(i)["Market Risk Remarks"]).rename("Comment").to_frame()   
    else:
        column_7=[]
        for j in range(len(column_2)):
            column_7.append("NA")
        column_7=pd.Series(column_7).rename("Comment").to_frame()
        
    #Removing recall rows
    if "Workflow Status" in list(FMSW_rawdata.parse(i)):
        column_a=FMSW_rawdata.parse(i)["Workflow Status"]
    elif "Workflow Event Status" in list(FMSW_rawdata.parse(i)):
        column_a=FMSW_rawdata.parse(i)["Workflow Event Status"].rename("Workflow Status")
    else:
        column_a=[]
        for j in range(len(column_2)):
            column_a.append("NA")
        column_a=pd.Series(column_a).rename("Workflow Status").to_frame()

    #Change attestation status rows from late to not attested
    if i=="FM Supervisory Attestation Brea":
        column_b=FMSW_rawdata.parse(i)["Attestation Date"]
    else:
        column_b=[]
        for j in range(len(column_2)):
            column_b.append(np.nan)
        column_b=pd.Series(column_b).rename("Attestation Date").to_frame()
        
    combine=pd.concat([column_2,column_3,column_4,column_5,column_6,column_7,column_a,column_b],axis=1)
    final=pd.concat([final,combine]).reset_index(drop=True)
#%%
#To merge the FMSW data with staff data
final=pd.concat([final["Employee PSID - Name"].str.split(" - ",n=1,expand=True),final.iloc[:,1:]],axis=1).reset_index(drop=True)
# final.set_axis(axis=1,labels=["Employee PSID","Name","Type of Breach","Severity","Accountability","Date","Comment","Workflow Status","Attestation Date"])
final = final.set_axis(axis=1,labels=["Employee PSID","Name","Type of Breach","Severity","Accountability","Date","Comment","Workflow Status","Attestation Date"])
final=final.drop("Name",axis=1)
final["FMSW/non-FMSW"]="FMSW"

#%%
#G&E
G_E=FMSW_rawdata.parse('Gifts & Entertainment',parse_dates= ["As of Date"])
G_E=G_E[G_E["G&E Request Status"]!="Approved"]
G_E=G_E[G_E["G&E Delayed Ageing"]>7]
G_E=G_E[np.logical_or(np.logical_and(G_E["Value per Individual Participant in USD"]>100,
                       G_E["Is Public Official involved in this G&E"]=="Yes")
,np.logical_and(G_E["Value per Individual Participant in USD"]>200,
                       G_E["Is Public Official involved in this G&E"]=="No"))]
G_E["Name"]=G_E["Name"].str.split(' -',expand=True)[0]
G_E=G_E[["Name","Reason for Delayed Registration","As of Date"]]

G_E["Type of Breach"]="G&E"
G_E["Accountability"]="Not Available"
G_E["Severity"]="High"
G_E=G_E[["Name","Type of Breach","Severity","Accountability","As of Date","Reason for Delayed Registration"]]
# G_E.set_axis(axis=1,labels=["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment"])
G_E = G_E.set_axis(axis=1,labels=["Employee PSID","Type of Breach","Severity","Accountability","Date","Comment"])

G_E["Employee PSID"]=G_E["Employee PSID"].astype(int).astype(str)

#%%Collect non-FMSW data
list_non_fmsw=[data_breach,data_leakage,G_E,PAD,TCIW,FLM,EOD,BJ,Ops,Others]
non_FMSW=pd.concat(list_non_fmsw)
non_FMSW["FMSW/non-FMSW"]="non-FMSW"
non_FMSW['Employee PSID'] = non_FMSW['Employee PSID'].astype(str)
#%%Merge FMSW and Non-FMSW data
final=pd.concat([final,non_FMSW], ignore_index=True)
#%%
#Remove workflow status that have recall in it
final["Workflow Status"]=final["Workflow Status"].fillna("NA")
final=final[np.invert(final["Workflow Status"].str.contains("Recall"))]
#Remove best execution alert that have status of green
final=final[np.logical_not(np.logical_and(final["Type of Breach"]=="Best Execution Alerts",final["Severity"].str.contains("GREEN")))]

#Import stafflist
Staff_rawdata=pd.read_excel("Stafflist.xlsx")
columns_keep=["Bank Id","Staff Name","Staff Country","Staff Region","Supervisor Id","Supervisor Name","Business Function Level 6 Desc"]
Staff_rawdata=Staff_rawdata[columns_keep]
# Staff_rawdata["Bank Id"]=Staff_rawdata.astype("str")
Staff_rawdata["Bank Id"]=Staff_rawdata["Bank Id"].astype("str")
Staff_rawdata["Supervisor Id"]=Staff_rawdata["Supervisor Id"].astype("str")
# Staff_rawdata.set_axis(axis=1,labels=["Employee PSID","Name","Location","Region","LM PSID","LM Name","Business (Lvl 6)"])
Staff_rawdata = Staff_rawdata.set_axis(axis=1,labels=["Employee PSID","Name","Location","Region","LM PSID","LM Name","Business (Lvl 6)"])

#Join staff details
testing=final.join(Staff_rawdata.set_index('Employee PSID'),on='Employee PSID')

#Join LM details
LM_rawdata=Staff_rawdata[['Employee PSID','Location','Region']]
LM_rawdata.columns=[['LM PSID','LM Location','LM Region']]
testing=testing.join(LM_rawdata.set_index('LM PSID'),on='LM PSID')

#Split EU and US location
testing["Region"]=np.where(testing["Location"]=="United States","Americas",testing["Region"])
testing["Region"]=np.where(testing["Region"]=="Europe & Americas","Europe",testing["Region"])
# testing = testing.set_axis(axis=1,labels=["Employee PSID", "Type of Breach", "Severity", "Accountability", "Date", "Comment", "Workflow Status", "Attestation Date", "FMSW/non-FMSW", "Incident ID", "Staff ID", "Final Severity Rating", "Month of Close Date", "Breach Description", "Name", "Location", "Region", "LM PSID", "LM Name", "Business (Lvl 6)", "LM Location", "LM Region"])
# testing["LM Region"]=np.where(testing["LM Location"]=="United States","Americas",testing["LM Region"])
# testing["LM Region"]=np.where(testing["LM Region"]=="Europe & Americas","Europe",testing["LM Region"])
testing[[('LM Region',)]]=np.where(testing[[('LM Location',)]]=="United States","Americas",testing[[('LM Region',)]])
testing[[('LM Region',)]]=np.where(testing[[('LM Region',)]]=="Europe & Americas","Europe",testing[[('LM Region',)]])

#Convert quarter date to actual date as attestation is lag by 1 quarter
testing["Date"]=np.where(testing["Date"].str[:2]=="Q4","1 March "+testing["Date"].str[-4:],testing["Date"])
testing["Date"]=np.where(testing["Date"].str[:2]=="Q1","1 June "+testing["Date"].str[-4:],testing["Date"])
testing["Date"]=np.where(testing["Date"].str[:2]=="Q2","1 September "+testing["Date"].str[-4:],testing["Date"])
testing["Date"]=np.where(testing["Date"].str[:2]=="Q3","1 December "+testing["Date"].str[-4:],testing["Date"])
testing["Date"]=pd.to_datetime(testing["Date"])
#Lag the attestation date by 1 quarter
testing["Attestation Date"]=pd.to_datetime(testing["Attestation Date"],dayfirst=True)
testing["Date"]=np.where(np.logical_and(testing["Type of Breach"]=="FM Supervisory Attestation Brea",testing["Date"].dt.month==3),testing["Date"]+pd.offsets.DateOffset(years=1),testing["Date"])
#Remove PAD breaches from Q2 2022 onward as these are moved to non-FMSW
testing=testing[np.invert(np.logical_and(testing["Type of Breach"]=="PAD Breaches",testing["Date"]>'2022-03-31'))]

#Amend the name of the breaches
testing["Type of Breach"]=np.where(testing["Type of Breach"]=="Group Mandatory E - Learning Ov","Group Mandatory e-Learning Overdue",testing["Type of Breach"])
testing["Type of Breach"]=np.where(testing["Type of Breach"]=="Best Execution Alerts","Best Execution",testing["Type of Breach"])
testing["Type of Breach"]=np.where(testing["Type of Breach"]=="Gifts & Entertainment","G&E",testing["Type of Breach"])
testing["Type of Breach"]=np.where(testing["Type of Breach"]=="Pre & Post Trade Mid Mark","PTMM",testing["Type of Breach"])
testing["Type of Breach"]=np.where(testing["Type of Breach"]=="Missed Trade","Missed trade",testing["Type of Breach"])
testing["Type of Breach"]=np.where(testing["Type of Breach"]=="FM Supervisory Attestation Brea","FM Supervisory Attestation Breach",testing["Type of Breach"])
testing["Type of Breach"]=np.where(testing["Type of Breach"]=="PAD Breaches","PAD breaches",testing["Type of Breach"])
#Amend the severity
testing["Severity"]=np.where(testing["Severity"]=="RED","Rag - RED",testing["Severity"])
testing["Severity"]=np.where(testing["Severity"]=="AMBER","Rag - AMBER",testing["Severity"])
testing["Severity"]=np.where(testing["Severity"]=="MATERIAL","Material",testing["Severity"])
testing["Severity"]=np.where(testing["Severity"]=="Very Low (1)","Very Low",testing["Severity"])
testing["Severity"]=np.where(testing["Severity"]=="Low (2)","Low",testing["Severity"])
testing["Severity"]=np.where(testing["Severity"]=="Medium (3)","Medium",testing["Severity"])
testing["Severity"]=np.where(testing["Severity"]=="High (4)","High",testing["Severity"])
#Create the materiality column
testing["Materiality"]=testing["Type of Breach"]+" "+testing["Severity"]
testing["Materiality"]=np.where(testing["Materiality"].str.contains("NA"),"NA",testing["Materiality"])

#%%
#Import the mapping file for sub categories
Mapping_rawdata=pd.read_excel("Mapping.xlsx")
Mapping_rawdata=Mapping_rawdata[["Breach type","Sub categories"]]
Mapping_rawdata=Mapping_rawdata.rename(columns={"Breach type":"Type of Breach"})
testing=testing.join(Mapping_rawdata.set_index('Type of Breach'),on='Type of Breach')

#Join the mapping file for materiality with main data
Mapping_rawdata=pd.read_excel("Mapping.xlsx")
Materiality_mapping=Mapping_rawdata[["Materiality","Material?"]]
testing=testing.join(Materiality_mapping.set_index('Materiality'),on='Materiality')
testing.drop_duplicates(inplace=True)

#Manually tag a few materiality
testing["Material?"]=np.where(testing["Type of Breach"]=="Trader Mandate Exception","Material",testing["Material?"])
testing["Material?"]=np.where(testing["Type of Breach"]=="ISDA Stays","Material",testing["Material?"])
testing["Material?"]=np.where(testing["Type of Breach"]=="PTMM","non-Material",testing["Material?"])
testing["Material?"]=np.where(np.logical_and(testing["Type of Breach"]=="FM Supervisory Attestation Breach",testing["Date"]<testing["Attestation Date"]),"Material",testing["Material?"])

#PTMM need to be more than 3 in a year to be considered as material
testing_a=testing[['Employee PSID','Type of Breach','Material?']]
testing_a=testing_a.groupby(testing_a.columns.tolist()).size().reset_index().rename(columns={0:'Records'})
testing=pd.merge(testing,testing_a,how='left')
testing['Material?']=np.where(np.logical_and(testing["Type of Breach"]=="PTMM",testing['Records']>3),"Material",testing['Material?'])

#Remove staff that already left the bank
testing=testing[testing['Region'].notnull()]
# testing=testing[np.invert(testing["Business (Lvl 6)"].str.contains("null"))]
testing=testing[np.invert(testing["Business (Lvl 6)"].str.contains("null", na=False))]
testing=testing[np.logical_and(testing['Business (Lvl 6)'].notnull(),np.invert(testing["Name"].str.contains("Korea")))]

#Remove non-FM individuals
non_fm=testing[np.invert(testing["Business (Lvl 6)"].str.contains("FM|CF"))]
testing=testing[testing["Business (Lvl 6)"].str.contains("FM|CF")]
non_fm.to_excel("non_fm.xlsx",index=False)

# testing=testing[[ 'Date','Employee PSID', 'Name','Location','Region', 'LM PSID', 'LM Name','LM Location','LM Region', 'Business (Lvl 6)', 'Type of Breach', 'Sub categories', 'Severity','Accountability','Materiality','Material?','Comment',"FMSW/non-FMSW"]]
testing=testing[[ 'Date','Employee PSID', 'Name','Location','Region', 'LM PSID', 'LM Name', ('LM Location',), ('LM Region',), 'Business (Lvl 6)', 'Type of Breach', 'Sub categories', 'Severity','Accountability','Materiality','Material?','Comment',"FMSW/non-FMSW"]]


#Manual amendment
testing=testing[np.invert(np.logical_and(np.logical_or(testing['Type of Breach']=="Group Mandatory e-Learning Overdue",testing['Type of Breach']=="FM Supervisory Attestation Breach"),testing['Material?']=="non-Material"))]
testing=testing[np.invert(testing['Name']=="Mohammad Naeem Khan")] #staff is just a driver
testing=testing[np.invert(np.logical_and(np.logical_and(testing['Type of Breach']=="Market Abuse (BlueJeans Recording)",testing['Name']=="Hala Tayyarah"),testing['Date']<'2022-06-30'))] #false exception for Q2 2022 conduct forum
testing=testing[np.invert(np.logical_and(np.logical_and(testing['Type of Breach']=="Best Execution",testing['Name']=="Moses Ndungu Kiboi"),testing['Date']<'2022-06-30'))] #false exception for Q2 2022 conduct forum
testing=testing[np.invert(np.logical_and(np.logical_and(testing['Type of Breach']=="Best Execution",testing['Name']=="Moses Ndungu Kiboi"),testing['Date']<'2022-12-30'))] #false exception for Q4 2022 conduct forum
testing=testing[np.invert(np.logical_and(np.logical_and(testing['Type of Breach']=="Best Execution",testing['Name']=="Ramish Shahid"),testing['Date']<'2022-06-30'))] #false exception for Q2 2022 conduct forum
testing=testing[np.invert(np.logical_and(np.logical_and(testing['Type of Breach']=="FM Supervisory Attestation Breach",testing['Name']=="Cosette Reczek"),testing['Date']<='2022-09-30'))] #false exception for Q3 2022 conduct forum
consolidated=testing[np.logical_and(np.logical_and(testing['Type of Breach']=="Missed trade",testing['Name']=="Kenneth Kabutha Ndirangu"),testing['Date']<'2022-06-30')].iloc[0:1]
testing=testing[np.invert(np.logical_and(np.logical_and(testing['Type of Breach']=="Missed trade",testing['Name']=="Kenneth Kabutha Ndirangu"),testing['Date']<'2022-06-30'))] #consolidate to a single item
testing=pd.concat([testing,consolidated], ignore_index=True)
testing=testing[np.invert(np.logical_and(np.logical_and(testing['Type of Breach']=="FM Supervisory Attestation Breach",testing['Business (Lvl 6)']=="CCIB FM MAG Core"),testing['Date']<='2022-09-30'))] #false exception for MAG team
testing=testing[np.invert(testing['Name']=="Sandeep Bahl")] #staff has left the bank
testing['Material?']=np.where(np.logical_and(np.logical_and(testing["Type of Breach"]=="Missed trade",testing['Name']=="Kenneth Kabutha Ndirangu"),testing['Date']<='2022-09-30'),"non-Material",testing['Material?'])
testing=testing[np.invert(np.logical_and(np.logical_and(testing['Type of Breach']=="FM Supervisory Attestation Breach",testing['Name']=="Chandrasekar Sridharan"),testing['Date']<='2022-09-30'))] #not a LM as of Q3 2022
testing=testing[np.invert(np.logical_and(np.logical_and(testing['Type of Breach']=="Group Mandatory e-Learning Overdue",testing['Name']=="Georgina Marie Higgins"),testing['Date']<'2022-09-30'))] #false exception for Q3 2022 conduct forum
testing=testing[np.invert(np.logical_and(np.logical_and(testing['Type of Breach']=="Best Execution",testing['Name']=="Veronica Vasco"),testing['Date']<'2022-09-30'))] #false exception for Q3 2022 conduct forum
testing=testing[np.invert(np.logical_and(np.logical_and(testing['Type of Breach']=="FM Supervisory Attestation Breach",testing['Name']=="Svetlana Stepcenkova"),testing['Date']<='2022-09-30'))] #staff was on maternity leave
testing=testing[np.invert(np.logical_and(np.logical_and(testing['Type of Breach']=="Data Breach",testing['Name']=="Chloe Friot"),testing['Date']<='2022-12-31'))] #20 - false exception Non material for Q4 Conduct Forum
testing=testing[np.invert(np.logical_and(np.logical_and(testing['Type of Breach']=="Client KYC Exceptions",testing['Name']=="Zubin M Sethna"),testing['Date']<='2022-12-31'))] #22 - false exception for Q4 Conduct Forum
testing=testing[np.invert(testing['Name']=="Gloria Gam-Ikon")] #staff has left the bank

#Merge Breach type into 1
consolidated=testing[np.logical_and(np.logical_and(testing['Type of Breach']=="Missed trade",testing['Name']=="Anjli Kaushik Bajaria"),testing['Date']<'2022-12-31')].iloc[0:1]
testing=testing[np.invert(np.logical_and(np.logical_and(testing['Type of Breach']=="Missed trade",testing['Name']=="Anjli Kaushik Bajaria"),testing['Date']<'2022-12-31'))] #consolidate to a single item
testing=pd.concat([testing,consolidated], ignore_index=True) #Merge Anjli's 8 Missed Trade into 1 - Q4

consolidated=testing[np.logical_and(np.logical_and(testing['Type of Breach']=="Missed trade",testing['Name']=="Sunday Matthew Olumoroti"),testing['Date']<'2022-12-31')].iloc[0:1]
testing=testing[np.invert(np.logical_and(np.logical_and(testing['Type of Breach']=="Missed trade",testing['Name']=="Sunday Matthew Olumoroti"),testing['Date']<'2022-12-31'))] #consolidate to a single item
testing=pd.concat([testing,consolidated], ignore_index=True) #Merge Sunday's 2 Missed Trade into 1- Q4


#BlueJeans temporary materiality change
testing['Material?']=np.where(np.logical_and(testing["Type of Breach"]=="Market Abuse (BlueJeans Recording)",testing['Comment']=="1 breach(es) in a quarter"),"non-Material",testing['Material?'])


#Create the final current 12 month data
testing['Employee PSID']=testing['Employee PSID'].astype(str)
testing['LM PSID']=testing['LM PSID'].astype(str)
full_12_month=testing
full_12_month["Quarter"]=pd.PeriodIndex(full_12_month["Date"],freq='Q')
full_12_month["Quarter"]=full_12_month["Quarter"].astype(str).str[-2:]+" "+full_12_month["Date"].dt.year.astype(str)

# full_12_month=full_12_month[['Quarter','Date','Employee PSID', 'Name','Location','Region', 'LM PSID', 'LM Name','LM Location','LM Region', 'Business (Lvl 6)', 'Type of Breach', 'Sub categories', 'Severity','Accountability','Materiality','Material?','Comment',"FMSW/non-FMSW"]]
full_12_month=full_12_month[['Quarter','Date','Employee PSID', 'Name','Location','Region', 'LM PSID', 'LM Name',('LM Location',),('LM Region',), 'Business (Lvl 6)', 'Type of Breach', 'Sub categories', 'Severity','Accountability','Materiality','Material?','Comment',"FMSW/non-FMSW"]]

#%%Show Trader and Sales
Staff_rawdata=pd.read_excel("Stafflist.xlsx")
columns_keep=["Bank Id","Role"]
Staff_rawdata=Staff_rawdata[columns_keep]
#Staff_rawdata["Bank Id"]=Staff_rawdata.astype("str")
Staff_rawdata["Bank Id"]=Staff_rawdata["Bank Id"].astype("str")
Staff_rawdata.set_axis(axis=1,labels=["Employee PSID","Role"])
Staff_rawdata = Staff_rawdata.set_axis(axis=1,labels=["Employee PSID","Role"])
Staff_rawdata["Role"]=np.where(Staff_rawdata["Role"]=="Local Sales","Sales",Staff_rawdata["Role"])
Staff_rawdata["Role"]=np.where(Staff_rawdata["Role"]=="Global Sales","Sales",Staff_rawdata["Role"])
Staff_rawdata["Role"]=np.where(Staff_rawdata["Role"]=="Sales","Sales","Trader")
full_12_month=full_12_month.join(Staff_rawdata.set_index('Employee PSID'),on='Employee PSID')

#%%Manual addition of breaches (Korea etc.)
manual_raw=pd.read_excel("manual_addition.xlsx")
full_12_month=pd.concat([full_12_month,manual_raw], ignore_index=True)

#%%Comparison with previous 12 months data
#Import previous 12-month data
full_12_month_prev=pd.read_excel("FMSWDataprev_12.xlsx",converters={'Employee PSID':str,'LM PSID':str})
full_12_month_prev=full_12_month_prev.fillna("NA")
full_12_month=full_12_month.fillna("NA")
full_12_month=full_12_month.replace(r'^\s*$',"NA",regex=True)
full_12_month_prev=full_12_month_prev[full_12_month_prev['Date']>cut_offdate]
full_12_month=full_12_month[full_12_month['Date']>cut_offdate]

checking=pd.merge(full_12_month,full_12_month_prev,how='outer',indicator=True)
checking_2=checking[checking['_merge']=='left_only']
checking_3=checking[checking['_merge']=='both']
checking_4=checking[checking['_merge']=='right_only']

checking_2_pivot=checking_2.pivot_table(index=["Type of Breach"],values="Employee PSID",aggfunc='count')
checking_4_pivot=checking_4.pivot_table(index=["Type of Breach"],values="Employee PSID",aggfunc='count')
checking_3_pivot=checking_3.pivot_table(index=["Type of Breach"],values="Employee PSID",aggfunc='count')

checking_2.to_excel("cur.xlsx",index=False)
checking_4.to_excel("prev.xlsx",index=False)

checking_2=checking_2[checking_2["Quarter"]==Quarter_list[-1]]

#%%
checking_5=checking_2
testing=checking_2

#For Material credit risk, market risk and missed trade add severity
checking_2["Type of Breach"]=np.where(checking_2["Type of Breach"]=="Credit Risk Excess",checking_2["Materiality"],checking_2["Type of Breach"])
checking_2["Type of Breach"]=np.where(checking_2["Type of Breach"]=="Market Risk Excess",checking_2["Materiality"],checking_2["Type of Breach"])
checking_2["Type of Breach"]=np.where(checking_2["Type of Breach"]=="Missed trade",checking_2["Materiality"],checking_2["Type of Breach"])


#%% Merge Name and Create Material breach in Q
testing_3=testing[['Employee PSID', 'Name','Type of Breach','Material?']]
testing_3=testing_3.groupby(testing_3.columns.tolist()).size().reset_index().rename(columns={0:'Records'})
testing_3['Type of Breach']=testing_3['Type of Breach'].astype(str)+' ('+testing_3['Records'].astype(str)+")"
testing_3=testing_3[testing_3['Material?']=="Material"]
#Serve as a testing file in case no setaff with repeated material breach
#==============================================================================
# test_row=[["1635756","Craig Makin","Testing","Material"],["1635756","Craig Makin","Testing","Material"]]
# test_row=pd.DataFrame(test_row,columns=['Employee PSID', 'Name','Type of Breach','Material?'])
# testing_3=testing_3.append(test_row,ignore_index=True)
#==============================================================================
testing_3=testing_3.drop_duplicates()

testing_4=testing_3[['Employee PSID','Type of Breach']]
testing_4=testing_4.groupby("Employee PSID")['Type of Breach'].apply(', '.join).reset_index()
testing_3=testing_3[['Employee PSID','Name']]
testing_4=testing[['Employee PSID', 'Name']].drop_duplicates().join(testing_4.set_index('Employee PSID'),on='Employee PSID')   
testing_4=testing_4.drop_duplicates() 
testing_4.columns=[[('Employee PSID', ''),('Name', ''),('Material breach in 12 months', '')]]
#  final_three_months=testing_4[testing_4[('Material breach in 12 months', '')].notnull()]
final_three_months=testing_4[testing_4[(('Material breach in 12 months', ''),)].notnull()]


#%% Merge Name and Create Staff and Material breach in Q
testing_3=testing[['LM Name','Type of Breach','Material?']]
testing_3=testing_3.groupby(testing_3.columns.tolist()).size().reset_index().rename(columns={0:'Records'})
testing_3['Type of Breach']=testing_3['Type of Breach'].astype(str)+' ('+testing_3['Records'].astype(str)+")"
testing_3=testing_3[testing_3['Material?']=="Material"]

testing_3=testing_3.drop_duplicates()
testing_4=testing_3.groupby("LM Name")['Type of Breach'].apply(', '.join).reset_index()

testing_5=testing[['LM Name','Name','Material?']]
testing_5=testing_5[testing_5['Material?']=="Material"]

testing_5=testing_5.drop_duplicates()
testing_6=testing_5.groupby("LM Name")['Name'].apply(' , '.join).reset_index()

testing_6=testing_6.join(testing_4.set_index('LM Name'),on='LM Name')
testing_6=testing_6.drop_duplicates() 
testing_6.columns=[[('LM Name', ''),('Staff with Material Breaches in 12 months',''),('Material breach in 12 months', '')]]
final_lm_three_months=testing_6



#%% Continuation of step 2
testing=full_12_month

# final_three_months = jaja
final_three_months.columns=[['Employee PSID','Name','Material breach in Q']]
# testing=testing.join(final_three_months[['Employee PSID','Material breach in Q']].set_index('Employee PSID'),on='Employee PSID')
testing = testing.merge(final_three_months[['Employee PSID','Material breach in Q']], left_index=True, right_index=True, how='left')

# testing=testing[testing['Material breach in Q'].notnull()]
testing=testing[testing[('Material breach in Q',)].notnull()]

#%% Merge Name and Create Material breach in Year
testing_3=testing[['Employee PSID', 'Name','Type of Breach','Material?']]
testing_3=testing_3.groupby(testing_3.columns.tolist()).size().reset_index().rename(columns={0:'Records'})
testing_3['Type of Breach']=testing_3['Type of Breach'].astype(str)+' ('+testing_3['Records'].astype(str)+")"
testing_3=testing_3[testing_3['Material?']=="Material"]
#==============================================================================
# test_row=[["1635756","Craig Makin","Testing","Material"],["1635756","Craig Makin","Testing","Material"]]
# test_row=pd.DataFrame(test_row,columns=['Employee PSID', 'Name','Type of Breach','Material?'])
# testing_3=testing_3.append(test_row,ignore_index=True)
#==============================================================================
testing_3=testing_3.drop_duplicates()

testing_4=testing_3[['Employee PSID','Type of Breach']]
testing_4=testing_4.groupby("Employee PSID")['Type of Breach'].apply(', '.join).reset_index()

testing_3=testing_3[['Employee PSID','Name']]
testing_4=testing[['Employee PSID', 'Name']].drop_duplicates().join(testing_4.set_index('Employee PSID'),on='Employee PSID')   
testing_4=testing_4.drop_duplicates() 
testing_4.columns=[[('Employee PSID', ''),('Name', ''),('Material breach in 12 months', '')]]
# [('Material breach in 12 months', '')]="Material: "+testing_4[('Material breach in 12 months', '')].astype(str)
testing_4[(('Material breach in 12 months', ''),)]="Material: "+testing_4[(('Material breach in 12 months', ''),)].astype(str)
testing_4a=testing_4
#%% Merge Name and Create non-Material breach in Year
testing_3=testing[['Employee PSID', 'Name','Type of Breach','Material?']]
testing_3=testing_3.groupby(testing_3.columns.tolist()).size().reset_index().rename(columns={0:'Records'})
testing_3['Type of Breach']=testing_3['Type of Breach'].astype(str)+' ('+testing_3['Records'].astype(str)+")"
testing_3=testing_3[testing_3['Material?']=="non-Material"]

testing_3=testing_3.drop_duplicates()

testing_4=testing_3[['Employee PSID','Type of Breach']]
testing_4=testing_4.groupby("Employee PSID")['Type of Breach'].apply(', '.join).reset_index()

testing_3=testing_3[['Employee PSID','Name']]
testing_4=testing[['Employee PSID', 'Name']].drop_duplicates().join(testing_4.set_index('Employee PSID'),on='Employee PSID')   
testing_4=testing_4.drop_duplicates() 
testing_4.columns=[[('Employee PSID', ''),('Name', ''),('non-Material breach in 12 months', '')]]
testing_4=testing_4.dropna()
# testing_4[('non-Material breach in 12 months', '')]="non-Material: "+testing_4[('non-Material breach in 12 months', '')].astype(str)
testing_4[(('non-Material breach in 12 months', ''),)]="non-Material: "+testing_4[(('non-Material breach in 12 months', ''),)].astype(str)
testing_4=pd.merge(testing_4a,testing_4,how='left')

# testing_4[('Breaches in 12 months','')]=np.where(pd.isnull(testing_4[('non-Material breach in 12 months', '')]),testing_4[('Material breach in 12 months', '')],testing_4[('Material breach in 12 months', '')].astype(str)+"\n"+testing_4[('non-Material breach in 12 months', '')].astype(str))
testing_4[(('Breaches in 12 months', ''),)]=np.where(pd.isnull(testing_4[(('non-Material breach in 12 months', ''),)]),testing_4[(('Material breach in 12 months', ''),)],testing_4[(('Material breach in 12 months', ''),)].astype(str)+"\n"+testing_4[(('non-Material breach in 12 months', ''),)].astype(str))

testing_4.columns=[[('Employee PSID'),('Name'),('Material breach in 12 months'),('non-Material breach in 12 months'),('Breaches in 12 months')]]
testing_4=testing_4[[('Employee PSID', ''),('Name', ''),('Breaches in 12 months','')]]
# testing_4=testing_4[(('Employee PSID', ''),), (('Name', ''),),(('Breaches in 12 months', ''),)]

#%% pivot for employee
full_sub_cat_listing=pd.read_excel("focus_breach_type.xlsx")
full_sub_cat=full_sub_cat_listing["Breach Type"]
full_material=["Material","non-Material"]
testing["Focus"]=np.where(testing["Type of Breach"].isin(full_sub_cat),testing["Type of Breach"],"Others")

full_column_list=[(p1,p2) for p1 in full_material for p2 in full_sub_cat]
testing_2=testing.pivot_table(index=["Region","Location","Employee PSID"],columns=["Material?","Focus"],values="Name",aggfunc='count')
testing_2.index=pd.MultiIndex.from_tuples(testing_2.index,names=["Region","Location","Employee PSID"])

testing_2=testing_2.reindex(columns=full_column_list)

testing_2[('Total', 'Breaches in last 12m')]=testing_2.sum(axis=1)
testing_2[('Total', 'Material breaches in last 12m')]=testing_2['Material'].sum(axis=1)
testing_2[('Total', 'Non-material breaches in last 12m')]=testing_2['non-Material'].sum(axis=1)

# testing_3=testing_2.reset_index()

testing_3=testing_2.reset_index()
final_three_months.columns=[[('Employee PSID', ''),('Name', ''),('Material breach in Q', '')]]
# ORIGINAL COMMAND
testing_3=pd.merge(testing_3,final_three_months,how='left')
# NEW COMMAND
# testing_3 = pd.concat([testing_3, final_three_months], axis=1).reindex(testing_3.index)
# final_three_months = final_three_months.reset_index()
# testing_3["Employee PSID"] = testing_3["Employee PSID"].astype(int)
# final_three_months["Employee PSID"] = final_three_months["Employee PSID"].astype(int)
testing_3.to_excel("testing_3.xlsx")
final_three_months.to_excel("final_three_months.xlsx")
print (""" Send to output excel - testing_3,final_three_months - for merging - testing_3,final_three_months 
       due to error and build new testing_3 in excel by merging - testing_3,final_three_months
       then fetch the new testing_3 dataset - To check with Vidya the output format 
       and change column names if needed in testing_3 """)
# testing_3=pd.merge(testing_3,final_three_months,how='left',left_on='Employee PSID', right_on='Employee PSID')
testing_3 = pd.DataFrame(testing_3)
testing_3 = pd.read_excel('testing_3.xlsx')


testing_4.to_excel("testing_4.xlsx")
print("Remove the line below column names in testing_4 and read the dataset again, due to issue with merging")
testing_4 = pd.read_excel('testing_4.xlsx')
testing_3=pd.merge(testing_3,testing_4,how='left')
# Staff_rawdata.columns=[[('Employee PSID', ''),('Role', '')]]
# Staff_rawdata.columns=[[('Employee PSID'),('Role')]]
testing_3["Employee PSID"] = testing_3["Employee PSID"].fillna(0)
testing_3["Employee PSID"] = testing_3["Employee PSID"].astype(int)
Staff_rawdata["Employee PSID"] = Staff_rawdata["Employee PSID"].fillna(0)
Staff_rawdata["Employee PSID"] = Staff_rawdata["Employee PSID"].astype(int)
# testing_3=testing_3.join(Staff_rawdata.set_index('Employee PSID'),on='Employee PSID')
testing_3 = pd.merge(testing_3,Staff_rawdata,on='Employee PSID',how='left')

testing_5=list(testing_3.columns)
reorder_column=list(testing_5[:3])+list(testing_5[-4:-3])+list(testing_5[-1:])+list(testing_5[-3:-1])+list(testing_5[-7:-4])+list(testing_5[3:13])
testing_5=testing_3[reorder_column]
appendix_staff=testing_5

testing_5=testing_5[np.logical_and(testing_5[('Total', 'Material breaches in last 12m')]>0,testing_5[('Total', 'Breaches in last 12m')]>1)]
testing_5.to_excel("testing_5.xlsx")
#Sort
testing_5=testing_5.sort_values(by=[('Total', 'Material breaches in last 12m'),('Total', 'Non-material breaches in last 12m')],ascending=[False,False])
#Arranging the columns
testing_5=testing_5.drop([('Total', 'Material breaches in last 12m')],axis=1)
testing_5=testing_5.drop([('Total', 'Non-material breaches in last 12m')],axis=1)

testing_5=testing_5.drop([('Employee PSID', '')],axis=1)
testing_5["Name"]=testing_5["Name"]+" ("+testing_5["Location"]+")"
testing_5=testing_5.rename(columns={('Total', 'Non-material breaches in last 12m'):('Total breaches in last 12m', '')})

Country_head=pd.read_excel("FM Head.xlsx")
Country_head=Country_head[["FM Head","Location"]]
Country_head.columns=[[('Country Head', ''),('Location', '')]]
testing_5=pd.merge(testing_5,Country_head,how='left')
reorder_column=list((testing_5.columns)[-1:])+list((testing_5.columns)[:-1])
testing_5=testing_5[reorder_column]
testing_5=testing_5.drop([('Location', '')],axis=1)

#%%Appendix Staff
#Sort
appendix_staff=appendix_staff.sort_values(by=[('Total', 'Material breaches in last 12m'),('Total', 'Non-material breaches in last 12m')],ascending=[False,False])
#tidying
appendix_staff=appendix_staff.drop([('Total', 'Material breaches in last 12m')],axis=1)
appendix_staff=appendix_staff.drop([('Total', 'Non-material breaches in last 12m')],axis=1)

appendix_staff=appendix_staff.drop([('Employee PSID', '')],axis=1)
appendix_staff["Name"]=appendix_staff["Name"]+" ("+appendix_staff["Location"]+")"
appendix_staff=appendix_staff.rename(columns={('Total', 'Non-material breaches in last 12m'):('Total breaches in last 12m', '')})

Country_head=pd.read_excel("FM Head.xlsx")
Country_head=Country_head[["FM Head","Location"]]
Country_head.columns=[[('Country Head', ''),('Location', '')]]
appendix_staff=pd.merge(appendix_staff,Country_head,how='left')
reorder_column=list((appendix_staff.columns)[-1:])+list((appendix_staff.columns)[:-1])
appendix_staff=appendix_staff[reorder_column]
appendix_staff=appendix_staff.drop([('Location', '')],axis=1)
#%% Splitting and tidying

AME=testing_5[testing_5[('Region','')]=="Africa & Middle East"]
Asia=testing_5[testing_5[('Region','')]=="Asia"]
Europe=testing_5[testing_5[('Region','')]=="Europe"]
Americas=testing_5[testing_5[('Region','')]=="Americas"]

AME=AME.set_index(('Region',''))
AME.index.names=['Region']
AME.columns.names=(None,None)

Asia=Asia.set_index(('Region',''))
Asia.index.names=['Region']
Asia.columns.names=(None,None)

Europe=Europe.set_index(('Region',''))
Europe.index.names=['Region']
Europe.columns.names=(None,None)

Americas=Americas.set_index(('Region',''))
Americas.index.names=['Region']
Americas.columns.names=(None,None)
#%% Splitting and tidying (Appendix)
AME_appendix=appendix_staff[appendix_staff[('Region','')]=="Africa & Middle East"]
Asia_appendix=appendix_staff[appendix_staff[('Region','')]=="Asia"]
Europe_appendix=appendix_staff[appendix_staff[('Region','')]=="Europe"]
Americas_appendix=appendix_staff[appendix_staff[('Region','')]=="Americas"]

AME_appendix=AME_appendix.set_index(('Region',''))
AME_appendix.index.names=['Region']
AME_appendix.columns.names=(None,None)

Asia_appendix=Asia_appendix.set_index(('Region',''))
Asia_appendix.index.names=['Region']
Asia_appendix.columns.names=(None,None)

Europe_appendix=Europe_appendix.set_index(('Region',''))
Europe_appendix.index.names=['Region']
Europe_appendix.columns.names=(None,None)

Americas_appendix=Americas_appendix.set_index(('Region',''))
Americas_appendix.index.names=['Region']
Americas_appendix.columns.names=(None,None)

#%% Merge Name and Create Staff and Material breach in Q
testing_3=testing[['LM Name','Type of Breach','Material?']]
testing_3=testing_3.groupby(testing_3.columns.tolist()).size().reset_index().rename(columns={0:'Records'})
testing_3['Type of Breach']=testing_3['Type of Breach'].astype(str)+' ('+testing_3['Records'].astype(str)+")"
testing_3=testing_3[testing_3['Material?']=="Material"]

testing_3=testing_3.drop_duplicates()
testing_4=testing_3.groupby("LM Name")['Type of Breach'].apply(', '.join).reset_index()

testing_5=testing[['LM Name','Name','Material?']]
testing_5=testing_5[testing_5['Material?']=="Material"]

testing_5=testing_5.drop_duplicates()
testing_6=testing_5.groupby("LM Name")['Name'].apply(', '.join).reset_index()

testing_6=testing_6.join(testing_4.set_index('LM Name'),on='LM Name')
testing_6=testing_6.drop_duplicates() 
testing_6.columns=[[('LM Name', ''),('Staff with Material Breaches in 12 months',''),('Material breach in 12 months', '')]]
testing_6[('Material breach in 12 months', '')]="Material: "+testing_6[('Material breach in 12 months', '')].astype(str)
testing_6a=testing_6


testing_3=testing[['LM Name','Type of Breach','Material?']]
testing_3=testing_3.groupby(testing_3.columns.tolist()).size().reset_index().rename(columns={0:'Records'})
testing_3['Type of Breach']=testing_3['Type of Breach'].astype(str)+' ('+testing_3['Records'].astype(str)+")"
testing_3=testing_3[testing_3['Material?']=="non-Material"]

testing_3=testing_3.drop_duplicates()
testing_4=testing_3.groupby("LM Name")['Type of Breach'].apply(', '.join).reset_index()

testing_5=testing[['LM Name','Name','Material?']]
testing_5=testing_5[testing_5['Material?']=="Material"]

testing_5=testing_5.drop_duplicates()
testing_6=testing_5.groupby("LM Name")['Name'].apply(', '.join).reset_index()

testing_6=testing_6.join(testing_4.set_index('LM Name'),on='LM Name')
testing_6=testing_6.drop_duplicates() 
testing_6.columns=[[('LM Name', ''),('Staff with Material Breaches in 12 months',''),('non-Material breach in 12 months', '')]]
testing_6=testing_6.dropna()
testing_6[('non-Material breach in 12 months', '')]="non-Material: "+testing_6[('non-Material breach in 12 months', '')].astype(str)

testing_6=pd.merge(testing_6a,testing_6,how='left')

testing_6[('Breaches in 12 months','')]=np.where(pd.isnull(testing_6[('non-Material breach in 12 months', '')]),testing_6[('Material breach in 12 months', '')],testing_6[('Material breach in 12 months', '')].astype(str)+"\n"+testing_6[('non-Material breach in 12 months', '')].astype(str))
testing_6=testing_6[[('LM Name', ''),('Staff with Material Breaches in 12 months', ''),('Breaches in 12 months','')]]


#%% pivot for LM
full_sub_cat_listing=pd.read_excel("focus_breach_type.xlsx")
full_sub_cat=full_sub_cat_listing["Breach Type"]
full_material=["Material","non-Material"]
testing["Focus"]=np.where(testing["Type of Breach"].isin(full_sub_cat),testing["Type of Breach"],"Others")

full_column_list=[(p1,p2) for p1 in full_material for p2 in full_sub_cat]
testing_2=testing.pivot_table(index=["LM Region","LM Location","LM Name"],columns=["Material?","Focus"],values="Name",aggfunc='count')
testing_2.index=pd.MultiIndex.from_tuples(testing_2.index,names=["LM Region","LM Location","LM Name"])

testing_2=testing_2.reindex(columns=full_column_list)

testing_2[('Total', 'Breaches in last 12m')]=testing_2.sum(axis=1)
testing_2[('Total', 'Material breaches in last 12m')]=testing_2['Material'].sum(axis=1)
testing_2[('Total', 'Non-material breaches in last 12m')]=testing_2['non-Material'].sum(axis=1)

testing_3=testing_2.reset_index()
final_lm_three_months.columns=[[('LM Name', ''),('Staff with Material Breaches in Q', ''),('Material breach in Q', '')]]

testing_3=pd.merge(testing_3,final_lm_three_months,how='left')
testing_3=pd.merge(testing_3,testing_6,how='left')
testing_5=list(testing_3.columns)
reorder_column=list(testing_5[:3])+list(testing_5[-4:-3])+list(testing_5[-2:-1])+list(testing_5[-3:-2])+list(testing_5[-1:])+list(testing_5[-7:-4])+list(testing_5[3:13])
testing_5=testing_3[reorder_column]
appendix_LM=testing_5
testing_5=testing_5[np.logical_and(testing_5[('Total', 'Material breaches in last 12m')]>1,testing_5[('Total', 'Breaches in last 12m')]>1)]
#Sort
testing_5=testing_5.sort_values(by=[('Total', 'Material breaches in last 12m'),('Total', 'Non-material breaches in last 12m')],ascending=[False,False])

testing_5=testing_5.drop([('Staff with Material Breaches in 12 months','')],axis=1)
testing_5=testing_5.drop([('Total', 'Material breaches in last 12m')],axis=1)
testing_5=testing_5.drop([('Total', 'Non-material breaches in last 12m')],axis=1)

testing_5["LM Name"]=testing_5["LM Name"]+" ("+testing_5["LM Location"]+")"
testing_5=testing_5.rename(columns={('Total', 'Non-material breaches in last 12m'):('Total breaches in last 12m', '')})

Country_head=pd.read_excel("FM Head.xlsx")
Country_head=Country_head[["FM Head","Location"]]
Country_head.columns=[[('Country Head', ''),('LM Location', '')]]
testing_5=pd.merge(testing_5,Country_head,how='left')
reorder_column=list((testing_5.columns)[-1:])+list((testing_5.columns)[:-1])
testing_5=testing_5[reorder_column]
testing_5=testing_5[testing_5[('Staff with Material Breaches in Q', '')].notnull()]
testing_5=testing_5[testing_5[('Staff with Material Breaches in Q', '')].str.contains(",")]
testing_5=testing_5.drop([('LM Location', ''),('Staff with Material Breaches in Q', '')],axis=1)

#%%Appendix LM
#Sort
appendix_LM=appendix_LM.sort_values(by=[('Total', 'Material breaches in last 12m'),('Total', 'Non-material breaches in last 12m')],ascending=[False,False])
appendix_LM=appendix_LM.drop([('Staff with Material Breaches in 12 months','')],axis=1)
appendix_LM=appendix_LM.drop([('Total', 'Material breaches in last 12m')],axis=1)
appendix_LM=appendix_LM.drop([('Total', 'Non-material breaches in last 12m')],axis=1)

appendix_LM["LM Name"]=appendix_LM["LM Name"]+" ("+appendix_LM["LM Location"]+")"
appendix_LM=appendix_LM.rename(columns={('Total', 'Non-material breaches in last 12m'):('Total breaches in last 12m', '')})

Country_head=pd.read_excel("FM Head.xlsx")
Country_head=Country_head[["FM Head","Location"]]
Country_head.columns=[[('Country Head', ''),('LM Location', '')]]
appendix_LM=pd.merge(appendix_LM,Country_head,how='left')
reorder_column=list((appendix_LM.columns)[-1:])+list((appendix_LM.columns)[:-1])
appendix_LM=appendix_LM[reorder_column]
appendix_LM=appendix_LM.drop([('LM Location', ''),('Staff with Material Breaches in Q', '')],axis=1)
appendix_LM=appendix_LM.sort_values(by=('Total', 'Breaches in last 12m'),ascending=False)

#%% Splitting and tidying

AME_LM=testing_5[testing_5[('LM Region','')]=="Africa & Middle East"]
Asia_LM=testing_5[testing_5[('LM Region','')]=="Asia"]
Europe_LM=testing_5[testing_5[('LM Region','')]=="Europe"]
Americas_LM=testing_5[testing_5[('LM Region','')]=="Americas"]

AME_LM=AME_LM.set_index(('LM Region',''))
AME_LM.index.names=['LM Region']
AME_LM.columns.names=(None,None)

Asia_LM=Asia_LM.set_index(('LM Region',''))
Asia_LM.index.names=['LM Region']
Asia_LM.columns.names=(None,None)

Europe_LM=Europe_LM.set_index(('LM Region',''))
Europe_LM.index.names=['LM Region']
Europe_LM.columns.names=(None,None)

Americas_LM=Americas_LM.set_index(('LM Region',''))
Americas_LM.index.names=['LM Region']
Americas_LM.columns.names=(None,None)

#%% Splitting and tidying (Appendix)
AME_appendix_LM=appendix_LM[appendix_LM[('LM Region','')]=="Africa & Middle East"]
Asia_appendix_LM=appendix_LM[appendix_LM[('LM Region','')]=="Asia"]
Europe_appendix_LM=appendix_LM[appendix_LM[('LM Region','')]=="Europe"]
Americas_appendix_LM=appendix_LM[appendix_LM[('LM Region','')]=="Americas"]

AME_appendix_LM=AME_appendix_LM.set_index(('LM Region',''))
AME_appendix_LM.index.names=['LM Region']
AME_appendix_LM.columns.names=(None,None)

Asia_appendix_LM=Asia_appendix_LM.set_index(('LM Region',''))
Asia_appendix_LM.index.names=['LM Region']
Asia_appendix_LM.columns.names=(None,None)

Europe_appendix_LM=Europe_appendix_LM.set_index(('LM Region',''))
Europe_appendix_LM.index.names=['LM Region']
Europe_appendix_LM.columns.names=(None,None)

Americas_appendix_LM=Americas_appendix_LM.set_index(('LM Region',''))
Americas_appendix_LM.index.names=['LM Region']
Americas_appendix_LM.columns.names=(None,None)


#%% output 
with pd.ExcelWriter('Group.xlsx') as writer:
    checking_5.to_excel(writer, sheet_name='Master Data',index=False)
    full_12_month.to_excel(writer, sheet_name='12-month Master Data',index=False)
    Asia.to_excel(writer, sheet_name='Asia')
    AME.to_excel(writer, sheet_name='AME')
    Europe.to_excel(writer, sheet_name='Europe')
    Americas.to_excel(writer, sheet_name='Americas')
    Asia_LM.to_excel(writer, sheet_name='Asia_Supervisor')
    AME_LM.to_excel(writer, sheet_name='AME_Supervisor')
    Europe_LM.to_excel(writer, sheet_name='Europe_Supervisor')
    Americas_LM.to_excel(writer, sheet_name='Americas_Supervisor')
    
    Asia_appendix.to_excel(writer, sheet_name='Asia_appendix')
    AME_appendix.to_excel(writer, sheet_name='AME_appendix')
    Europe_appendix.to_excel(writer, sheet_name='Europe_appendix')
    Americas_appendix.to_excel(writer, sheet_name='Americas_appendix')
    Asia_appendix_LM.to_excel(writer, sheet_name='Asia_Supervisor_app')
    AME_appendix_LM.to_excel(writer, sheet_name='AME_Supervisor_app')
    Europe_appendix_LM.to_excel(writer, sheet_name='Europe_Supervisor_app')
    Americas_appendix_LM.to_excel(writer, sheet_name='Americas_Supervisor_app')
writer.save()

with pd.ExcelWriter('Asia.xlsx') as writer:
    checking_5[checking_5["Region"]=="Asia"].to_excel(writer, sheet_name='Master Data',index=False)
    full_12_month[full_12_month["Region"]=="Asia"].to_excel(writer, sheet_name='12-month Master Data',index=False)
    Asia.to_excel(writer, sheet_name='Asia')
    Asia_LM.to_excel(writer, sheet_name='Asia_Supervisor')
    Asia_appendix.to_excel(writer, sheet_name='Asia_appendix')
    Asia_appendix_LM.to_excel(writer, sheet_name='Asia_Supervisor_app')
writer.save()

with pd.ExcelWriter('Europe.xlsx') as writer:
    checking_5[checking_5["Region"]=="Europe"].to_excel(writer, sheet_name='Master Data',index=False)
    full_12_month[full_12_month["Region"]=="Europe"].to_excel(writer, sheet_name='12-month Master Data',index=False)
    Europe.to_excel(writer, sheet_name='Europe')
    Europe_LM.to_excel(writer, sheet_name='Europe_Supervisor')
    Europe_appendix.to_excel(writer, sheet_name='Europe_appendix')
    Europe_appendix_LM.to_excel(writer, sheet_name='Europe_Supervisor_app')
writer.save()

with pd.ExcelWriter('AME.xlsx') as writer:
    checking_5[checking_5["Region"]=="Africa & Middle East"].to_excel(writer, sheet_name='Master Data',index=False)
    full_12_month[full_12_month["Region"]=="Africa & Middle East"].to_excel(writer, sheet_name='12-month Master Data',index=False)
    AME.to_excel(writer, sheet_name='AME')
    AME_LM.to_excel(writer, sheet_name='AME_Supervisor')
    AME_appendix.to_excel(writer, sheet_name='AME_appendix')
    AME_appendix_LM.to_excel(writer, sheet_name='AME_Supervisor_app')
writer.save()

with pd.ExcelWriter('Americas.xlsx') as writer:
    checking_5[checking_5["Region"]=="Americas"].to_excel(writer, sheet_name='Master Data',index=False)
    full_12_month[full_12_month["Region"]=="Americas"].to_excel(writer, sheet_name='12-month Master Data',index=False)
    Americas.to_excel(writer, sheet_name='Americas')
    Americas_LM.to_excel(writer, sheet_name='Americas_Supervisor')
    Americas_appendix.to_excel(writer, sheet_name='Americas_appendix')
    Americas_appendix_LM.to_excel(writer, sheet_name='Americas_Supervisor_app')
writer.save()


#%% Risk Profile
staff_count=pd.read_excel("Headcount Details Base Report.xlsx")
staff_count=staff_count[staff_count["Frontline/Non Frontline"]=="Frontline"]
staff_count["Region Name"]=np.where(staff_count["Country Name"]=="United States","Americas",staff_count["Region Name"])
staff_count["Region Name"]=np.where(staff_count["Region Name"]=="Europe & Americas","Europe",staff_count["Region Name"])
staff_count_pivot=staff_count.pivot_table(index=["Country Name"],values="Employee Bank ID",aggfunc='count')
staff_count_pivot=staff_count_pivot.reset_index()
staff_count_pivot.columns=["Location","Staff Count"]

full_12_month_mat=full_12_month[full_12_month["Material?"]=="Material"]
last_q_matbreach=full_12_month_mat[full_12_month_mat["Quarter"]==Quarter_list[-1]]

full_12_month_mat=full_12_month_mat.pivot_table(index=["Region","Location"],columns=["Quarter"],values="Name",aggfunc='count')
full_12_month_mat=full_12_month_mat.reset_index()
Country_head=pd.read_excel("FM Head.xlsx")
full_12_month_mat=pd.merge(Country_head,full_12_month_mat,how='outer')
full_12_month_mat=full_12_month_mat.join(staff_count_pivot[["Location","Staff Count"]].set_index('Location'),on='Location')
full_12_month_mat=full_12_month_mat[["Region","FM Head","Location","Staff Count"]+Quarter_list[1:]]
last_q_matbreach=last_q_matbreach[["Location","Type of Breach"]]
last_q_matbreach=last_q_matbreach.groupby(last_q_matbreach.columns.tolist()).size().reset_index().rename(columns={0:'Records'})
last_q_matbreach["Records"]=last_q_matbreach["Type of Breach"]+ " ("+last_q_matbreach["Records"].astype(str)+")"
last_q_matbreach=last_q_matbreach.drop(["Type of Breach"],axis=1)

last_q_matbreach=last_q_matbreach.groupby("Location")['Records'].apply(', '.join).reset_index()
full_12_month_mat=pd.merge(full_12_month_mat,last_q_matbreach,how='left')
full_12_month_mat=full_12_month_mat.set_index('Region')

#%% Risk Profile Appendix
full_12_month_matbreach=full_12_month[full_12_month["Material?"]=="Material"]
full_12_month_matbreach=full_12_month_matbreach.pivot_table(index=["Region","Type of Breach"],columns=["Quarter"],values="Name",aggfunc='count')
full_12_month_matbreach["Total"]=full_12_month_matbreach.sum(axis=1)
full_12_month_matbreach=full_12_month_matbreach[Quarter_list]
full_12_month_matbreach=full_12_month_matbreach.sort_values([Quarter_list[-1],Quarter_list[0]], ascending=[False,False]).sort_index(level=0, ascending=[True])

testing_groupby=full_12_month[full_12_month["Material?"]=="Material"].pivot_table(index=["Region"],values="Name",aggfunc='count')
full_12_month_matbreach=pd.merge(full_12_month_matbreach,pd.DataFrame(testing_groupby["Name"].rename("Percentage Quarter")),left_index=True, right_index=True)
full_12_month_matbreach["Percentage Quarter"]=full_12_month_matbreach["Total"]/full_12_month_matbreach["Percentage Quarter"]

full_12_month_regbreach=full_12_month[full_12_month["Material?"]=="Material"]
full_12_month_regbreach["sub_region"]=np.where(full_12_month_regbreach["Location"].isin(["Hong Kong","Taiwan","China","Korea (the Republic of)","Japan"]),"GCNA",np.nan)
full_12_month_regbreach["sub_region"]=np.where(full_12_month_regbreach["Location"].isin(["Singapore","Malaysia","Thailand","Philippines","Vietnam","Indonesia"]),"ASEAN",full_12_month_regbreach["sub_region"])
full_12_month_regbreach["sub_region"]=np.where(full_12_month_regbreach["Location"].isin(["Bangladesh","Sri Lanka","Nepal","India"]),"South Asia",full_12_month_regbreach["sub_region"])

full_12_month_regbreach["sub_region"]=np.where(full_12_month_regbreach["Location"].isin(["Iraq", "Oman", "Saudi Arabia", "United Arab Emirates", "Bahrain", "Jordan", "Pakistan", "Qatar"]),"MENAP",full_12_month_regbreach["sub_region"])
full_12_month_regbreach["sub_region"]=np.where(full_12_month_regbreach["Location"].isin(["Kenya","Tanzania, United Republic of","Uganda"]),"East Africa",full_12_month_regbreach["sub_region"])
full_12_month_regbreach["sub_region"]=np.where(full_12_month_regbreach["Location"].isin(["Nigeria", "Cameroon", "Cote D'Ivoire", "Gambia", "Ghana", "Sierra Leone"]),"West Africa",full_12_month_regbreach["sub_region"])
full_12_month_regbreach["sub_region"]=np.where(full_12_month_regbreach["Location"].isin(["Angola", "Botswana", "Mauritius", "South Africa", "Zimbabwe", "Zambia"]),"Southern Africa",full_12_month_regbreach["sub_region"])


full_12_month_regbreach=full_12_month_regbreach.pivot_table(index=["Region","sub_region"],columns=["Quarter"],values="Name",aggfunc='count')
full_12_month_regbreach["Total"]=full_12_month_regbreach.sum(axis=1)
full_12_month_regbreach=full_12_month_regbreach[Quarter_list]

#%% Group Risk Profile
full_12_month_matbreach_group=full_12_month[full_12_month["Material?"]=="Material"]
full_12_month_matbreach_group=full_12_month_matbreach_group.pivot_table(index=["Type of Breach"],columns=["Quarter"],values="Name",aggfunc='count')
full_12_month_matbreach_group["Total"]=full_12_month_matbreach_group.sum(axis=1)
full_12_month_matbreach_group=full_12_month_matbreach_group[Quarter_list]
full_12_month_matbreach_group=full_12_month_matbreach_group.sort_values([Quarter_list[-1],Quarter_list[0]], ascending=[False,False])
full_12_month_matbreach_group.loc['Total'] = full_12_month_matbreach_group.sum(numeric_only=True)


matbreach_region=full_12_month[full_12_month["Material?"]=="Material"]
matbreach_region=matbreach_region.pivot_table(index=["Region"],columns=["Quarter"],values="Name",aggfunc='count')
matbreach_region["Q_Q_change"]=(matbreach_region[Quarter_list[-1]]-matbreach_region[Quarter_list[-2]])/matbreach_region[Quarter_list[-2]]
matbreach_region["Last_Qdist"]=matbreach_region[Quarter_list[-1]]/sum(matbreach_region[Quarter_list[-1]])

#%%Europe regional ask
#==============================================================================
# full_12_month_eu=full_12_month[full_12_month["Region"]=="Europe"]
# full_12_month_eu['Month_Date'] = pd.to_datetime(full_12_month_eu['Date']).dt.strftime('%m/%Y')
# full_12_month_eu.to_excel("EU 12month_check.xlsx",index=False)
# full_12_month_eu=full_12_month_eu.replace(np.nan,'-')
# full_12_month_eu=full_12_month_eu.pivot_table(index=["Name","LM Name","Type of Breach","Severity","Accountability"],columns=["Month_Date"],values="Date",aggfunc='count')
# #==============================================================================
# # full_12_month_eu=pd.concat([y.append(y.sum().rename((x, 'Total'))) for x, y in full_12_month_eu.groupby(level=0)]).append(full_12_month_eu.sum().rename(('Grand', 'Total')))
# # 
# #==============================================================================
# full_12_month_eu.to_excel("EU 12month.xlsx") 
# 
#==============================================================================


#%% output 
with pd.ExcelWriter('Risk Profile.xlsx') as writer:
    full_12_month_mat.to_excel(writer, sheet_name='Risk Profile')
    full_12_month_matbreach.to_excel(writer, sheet_name='Risk Profile App')
    full_12_month_regbreach.to_excel(writer, sheet_name='Risk Profile App2')
    full_12_month_matbreach_group.to_excel(writer, sheet_name='Risk Profile(Group)')
    matbreach_region.to_excel(writer, sheet_name='Risk Profile(Group App)')
writer.save()

group=full_12_month_matbreach.groupby("Region")


