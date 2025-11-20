# -*- coding: utf-8 -*-
"""
Created on Tue Oct 25 14:26:34 2022

@author: 1659765
"""

import pandas as pd
import numpy as np

Others=pd.read_excel("Others.xlsx")

Others["Employee PSID"]=Others["Employee PSID"].astype(str)

Others.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Others\Others_Output.xlsx",index=False)
Others.to_excel(r"C:\Users\1510806\OneDrive - Standard Chartered Bank\Desktop\Conduct_Py_Scripts\Final\Others_Output.xlsx",index=False)