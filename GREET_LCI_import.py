# -*- coding: utf-8 -*-
"""
Created on Mon Jan  9 14:00:34 2023

@author: skar

Collect LCI data by year from GREET1 as per correspondence file.

"""

#%%

import pandas as pd
import xlwings as xw
from datetime import datetime

#%%
# GREET1 model run for study years

class GREET_LCI_import:
    
    def __init__(self, model_path_prefix, fmodel,
                 corr_path_prefix, fcorr_LCI, sheet_input, sheet_output,
                 gparam, sheet_gparam, cell_gparam,
                 fsave_sim):
        
        self.model_path_prefix = model_path_prefix
        self.fmodel = fmodel        
        
        self.corr_path_prefix = corr_path_prefix
        self.fcorr_LCI = fcorr_LCI
        self.sheet_input = sheet_input
        self.sheet_output = sheet_output
        
        self.gparam = gparam        
        self.sheet_gparam = sheet_gparam
        self.cell_gparam = cell_gparam
             
        self.fsave_sim = fsave_sim
        
        self.sim_df = pd.DataFrame() # initialize data frame to save runs
        
        # read parameter value sets
        self.sim_params = pd.read_excel(self.corr_path_prefix + '/' + self.fcorr_LCI, 
                                        self.sheet_input, header=3, index_col=None)
        
    
    def modify_model_and_run(self, gparam_index, gparam_val, df_params):
        
        with xw.App(visible=False) as app:            
            
            # open model workbook
            wb = xw.Book(model_path_prefix + '/' + fmodel)
            
            # modify model with global parameters
            for idx, val in enumerate(gparam_val):                
                sheet = wb.sheets[self.sheet_gparam[idx]]
                sheet[self.cell_gparam[idx]].value = val
                wb.app.calculate()
                wb.save()
            
            # modify model with set of parameters            
            for r in range(df_params.shape[0]):
                sheet = wb.sheets[df_params.iloc[r,0]]
                sheet[df_params.iloc[r,1]].value = df_params.iloc[r,2]
            wb.app.calculate()
            wb.save()
            
            # update output sheet
            wb2 = xw.Book(self.corr_path_prefix + '/' + self.fcorr_LCI)
            wb2.app.calculate()
            wb2.save()
            
        self.temp_corr_LCI = pd.read_excel(self.corr_path_prefix + '/' + self.fcorr_LCI, 
                                           self.sheet_output, header=3, index_col=None)
    
    def save_sim_to_file(self, mode, header):
        
        self.sim_df.to_csv(self.corr_path_prefix + '/' + self.fsave_sim, 
                           mode=mode, header=header, index=False)
    
    def sim_model(self):
        
        n_param_sets = self.sim_params.shape[1] - 3
        
        # truncate if file exists and create
        self.save_sim_to_file(mode='w', header=False)
        write_header = True
        
        for gparam_index, gparam_val in enumerate(gparam):
            
            for param_set in range(0, n_param_sets):
                
                df_params = self.sim_params.iloc[:,[1,2,param_set+3]]
                
                print(f'Executing global parameter set: {gparam_val} and parameter set {param_set+1} out of {n_param_sets}') 
                print( '    Elapsed time: ' + str(datetime.now() - init_time))
                
                self.modify_model_and_run(gparam_index, gparam_val, df_params)                
                self.temp_corr_LCI['gparam_val'] = '-'.join(map(str,gparam_val))
                self.temp_corr_LCI['param_set'] = param_set + 1
                
                #self.sim_df = pd.concat([self.sim_df, self.temp_corr_LCI.copy()], ignore_index=True)
                self.sim_df = self.temp_corr_LCI.copy()
                
                if write_header:
                    self.save_sim_to_file(mode='a', header=write_header) # append output to file
                    write_header = False
                else:
                    self.save_sim_to_file(mode='a', header=write_header) # append output to file
        

if __name__ == '__main__':
    
    init_time = datetime.now()
    
    model_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - Algae/data/model'
    fmodel = 'GREET_2022 Algae harmonization project_HTL_paper_1.xlsm'
    
    corr_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - Algae/data/correspondence_files'
    fcorr_LCI = 'corr_LCI_GREET_pathway_Algae_v1_run.xlsx'
    sheet_input = 'PC'
    sheet_output = 'map_outputs'
    
    fsave_sim = 'GREET_Algae_sims_PC.csv'
    
    # Global parameter declarations
    sheet_gparam = ['Algae', 'Algae'] # the sheets in fmodel that has the parameters
    cell_gparam = ['AI556', 'AF555'] # the cells in fmodel sheet_gparam where parameters are located
    gparam = [[1,1],  # value set of global parameters
              #[2,1],
              #[3,1],
              #[1,2],
              #[2,2],
              #[3,2],
              #[1,3],
              #[2,3],
              #[3,3]
              ]        
    
    obj = GREET_LCI_import(model_path_prefix, fmodel,
                           corr_path_prefix, fcorr_LCI, sheet_input, sheet_output,
                           gparam, sheet_gparam, cell_gparam,
                           fsave_sim)
    
    obj.sim_model()
    
    print( '    Total run time: ' + str(datetime.now() - init_time)) 
