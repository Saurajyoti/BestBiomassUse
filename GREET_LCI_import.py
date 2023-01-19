# -*- coding: utf-8 -*-
"""
Created on Mon Jan  9 14:00:34 2023

@author: skar

Collect LCI data by year from GREET1 as per correspondence file.

"""

#%%

import pandas as pd
import xlwings as xw

#%%
# GREET1 model run for study years

class GREET_LCI_import:
    
    def __init__(self, model_path_prefix, file_model, sheet_input,
                 corr_path_prefix, fcorr_LCI, sheet_corr_LCI,
                 start_year, end_year, increment,
                 file_save_sim):
        
        self.model_path_prefix = model_path_prefix
        self.file_model = file_model
        self.sheet_input = sheet_input
        
        self.corr_path_prefix = corr_path_prefix
        self.fcorr_LCI = fcorr_LCI
        self.sheet_corr_LCI = sheet_corr_LCI
        
        self.start_year = start_year
        self.end_year = end_year
        self.increment = increment      
        
        self.param_input_cell = 'E9' # Cell address of year parameter in the GREET1 2022 Inputs tab
        
        self.sim_df = pd.DataFrame() # initialize data frame to save runs
        
        self.file_save_sim = file_save_sim
        
    
    def modify_GREET2_and_run(self, param_input_cell, param_val):
        
        with xw.App(visible=False) as app:            
            wb = xw.Book(model_path_prefix + '/' + file_model)
            sheet = wb.sheets[self.sheet_input]
            sheet[param_input_cell].value = param_val
            wb.app.calculate()
            wb.save()
            
            wb2 = xw.Book(self.corr_path_prefix + '/' + self.fcorr_LCI)
            wb2.app.calculate()
            wb2.save()
            
        self.temp_corr_LCI = pd.read_excel(self.corr_path_prefix + '/' + self.fcorr_LCI, 
                                           self.sheet_corr_LCI, header=3, index_col=None)
    
    def save_sim_to_file(self):
        
        self.sim_df.to_csv(self.corr_path_prefix + '/' + self.file_save_sim)
    
    def sim_model(self):
        
        for y in range(self.start_year, self.end_year+1, self.increment):
            print(f'Currently extracting data for year: {y}')            
            self.modify_GREET2_and_run(self.param_input_cell, y)
            self.temp_corr_LCI['Year'] = y           
            self.sim_df = pd.concat([self.sim_df, self.temp_corr_LCI], ignore_index=True)
       
        self.save_sim_to_file()
        

if __name__ == '__main__':
    
    model_path_prefix = 'C:/Users/skar/Box/saura_self/GREET_2022'
    file_model = 'GREET1_2022.xlsm'
    sheet_input = 'Inputs'
    
    corr_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - Best use of biomass/data/correspondence_files'
    fcorr_LCI = 'corr_LCI_GREET_pathway.xlsx'
    sheet_corr_LCI = 'GREET_mappings'
    
    file_save_sim = 'corr_LCI_GREET_temporal.csv'
    
    
    start_year = 2020
    end_year = 2050
    increment = 1
    
    obj = GREET_LCI_import(model_path_prefix, file_model, sheet_input, 
                           corr_path_prefix, fcorr_LCI, sheet_corr_LCI,
                           start_year, end_year, increment,
                           file_save_sim)
    
    obj.sim_model()
