# -*- coding: utf-8 -*-
"""
Created on Wed Sep 28 12:45:05 2022

@author: 1659765
"""

import pandas as pd
import numpy as np

Ops=pd.read_excel("Ops.xlsx")

Ops["Employee PSID"]=Ops["Employee PSID"].astype(str)

Ops.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Operation Losses\Ops_Output.xlsx",index=False)
Ops.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Final\Ops_Output.xlsx",index=False)