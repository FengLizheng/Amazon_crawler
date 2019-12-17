# -*- coding: utf-8 -*-
"""
Created on Tue Dec 17 11:12:00 2019

@author: FENG Lizheng The university of Hong Kong

@E-mail:fenglz@hku.hk 

"""


import pandas as pd
import os
Folder_Path = r'C:\Users\V330\Desktop\combine'         
SaveFile_Path =  r'C:\Users\V330\Desktop\combine'       
SaveFile_Name = r'AmazonReview.csv'              
 

os.chdir(Folder_Path)

file_list = os.listdir()
 

df = pd.read_csv(Folder_Path +'\\'+ file_list[0])   
 

df.to_csv(SaveFile_Path+'\\'+ SaveFile_Name,encoding="utf_8_sig",index=False)
 

for i in range(1,len(file_list)):
    df = pd.read_csv(Folder_Path + '\\'+ file_list[i])
    df.to_csv(SaveFile_Path+'\\'+ SaveFile_Name,encoding="utf_8_sig",index=False, header=False, mode='a+')
