# -*- coding: utf-8 -*-
"""
Created on Fri Apr 24 12:10:19 2020

@author: TungNguyen
"""



import pandas as pd



dir_data = r'C:\Users\TungNguyen\Ionomr Innovations\Ionomr - Documents\R&D\Combined MFG and Test data/'
dir_archive = r'Archived/'
dir_MFG = r'C:\Users\TungNguyen\Ionomr Innovations\Ionomr - Documents\R&D\Reinforced Membrane Development\Non-woven\Mirwec - Labo Coating Trials/'
dir_exsitu = r'C:\Users\TungNguyen\Ionomr Innovations\Ionomr - Documents\R&D\Reinforced membrane development\Non-woven\Mirwec - Labo Coating Trials\Ex-situ data/'
dir_insitu = r'C:\Users\TungNguyen\Ionomr Innovations\Ionomr - Documents\R&D\2020 Testing\Miho/'


MFG_file = 'Labo coatings.xlsx'
Exsitu_file = 'Labo membranes_ex-situ.xlsx'
Insitu_file = 'Test Summary.xlsx'
# Data from the previous testing
AEM_file_old = 'AEM_output_Archive data April 24 2020.xlsx'
PEM_file_old = 'PEM_output_Archive data April 24 2020.xlsx'

# Read the present data
MFG_data = pd.read_excel(dir_MFG+MFG_file, index_col = 1, sheet_name ='LABO', header = 5,na_filter=False )
Exsitu_data = pd.read_excel(dir_exsitu+Exsitu_file, index_col = 0, sheet_name ='Sheet1', header = 0,na_filter=False )
PEM_Insitu_data = pd.read_excel(dir_insitu+Insitu_file, index_col = 0, sheet_name ='PEM', header = 0,na_filter=False )
AEM_Insitu_data = pd.read_excel(dir_insitu+Insitu_file, index_col = 0, sheet_name ='AEM', header = 0,na_filter=False )

# Read the archived data
# Old data before April 24, 2020 was messy and unorganized. 
# Old data from the master file was processed to remove douplicate. 
# Test summary and Exsitu file still carries old data while the MFG file starts with afresh sheet. 
AEM_old_data = pd.read_excel(dir_data+dir_archive+AEM_file_old, index_col = 0, sheet_name ='AEM_raw_data', header = 0,na_filter=False )
PEM_old_data = pd.read_excel(dir_data+dir_archive+PEM_file_old, index_col = 0, sheet_name ='PEM_raw_data', header = 0,na_filter=False )

# Merge exsitu and MFG data
# Using left because exsitu data can be late updated 
Exsitu_raw_data_left = pd.merge(MFG_data, Exsitu_data, how='left', on =['Sample'])
# We add the recent updated MFG and exsitu data with the old MFG and exsitu data from the previous master file  
PEM_raw_data = Exsitu_raw_data_left.append(PEM_old_data)
AEM_raw_data = Exsitu_raw_data_left.append(AEM_old_data)

# Merge insitu and exsitu data for each product
PEM_raw_data = pd.merge(PEM_raw_data, PEM_Insitu_data, how='right', on =['Sample'])
AEM_raw_data = pd.merge(AEM_raw_data, AEM_Insitu_data, how='right', on =['Sample'])
# delete rows with no index  (blank rows)
PEM_raw_data = PEM_raw_data[PEM_raw_data.index != r'']
AEM_raw_data = AEM_raw_data[AEM_raw_data.index != r'']
# save the raw data file to excel 
AEM_raw_data.sort_values(by = ['Sample']).to_excel(dir_data+"AEMION.xlsx",sheet_name='AEM_raw_data')  
PEM_raw_data.sort_values(by = ['Sample']).to_excel(dir_data+"PEMION.xlsx",sheet_name='PEM_raw_data')  

#TODO
# delete the old  excel data files in the Archived folder
# save the current excel data files to the Archived folder
