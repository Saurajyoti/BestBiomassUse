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
       
    def sim_model(self):
        
        with xw.App(visible=False) as app: 
            
            wb = xw.Book(model_path_prefix + '/' + file_model)
            wb_sheet = wb.sheets[self.sheet_input]
            
            wb2 = xw.Book(self.corr_path_prefix + '/' + self.fcorr_LCI)
            wb2_sheet = wb2.sheets[self.sheet_corr_LCI]
            
            for y in range(self.start_year, self.end_year+1, self.increment):
                print(f'Currently extracting data for year: {y}')            
                
                #self.modify_GREET2_and_run(self.param_input_cell, y)      
                wb_sheet[self.param_input_cell].value = y
                wb.app.calculate()
                wb.save() # not sure if the output Excel file with reference will access Model Excel file from disk or memory, so saving the updated model file everytime.              
                
                wb2.app.calculate()
                #wb2.save()
                
                #self.temp_corr_LCI  = wb2_sheet[range_of_output_sheet].options(pd.DataFrame).value.reset_index()
                self.temp_corr_LCI  = wb2_sheet.used_range.options(pd.DataFrame).value.reset_index()
                
                self.temp_corr_LCI['Year'] = y           
                self.sim_df = pd.concat([self.sim_df, self.temp_corr_LCI], ignore_index=True)
       
        self.sim_df.to_csv(self.corr_path_prefix + '/' + self.file_save_sim)        

if __name__ == '__main__':
    
    model_path_prefix = 'C:/Users/skar/Box/saura_self/GREET_2022'
    file_model = 'GREET1_2022.xlsm'
    sheet_input = 'Inputs'
    
    corr_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - Best use of biomass/data/correspondence_files'
    fcorr_LCI = 'corr_LCI_GREET_pathway_03_24_2023.xlsx'
    sheet_corr_LCI = 'GREET_mappings'
    # Please update the range of cells to extract the data if the table changes
    #range_of_output_sheet='A4:L3070'
    
    file_save_sim = 'corr_LCI_GREET_temporal_03_24_2023.csv'   
    
    start_year = 2021
    end_year = 2021
    increment = 1
    
    obj = GREET_LCI_import(model_path_prefix, file_model, sheet_input, 
                           corr_path_prefix, fcorr_LCI, sheet_corr_LCI,
                           start_year, end_year, increment,
                           file_save_sim)    
    obj.sim_model()
