# -*- coding: utf-8 -*-
"""
Created on Mon Jan  9 14:00:34 2023

@author: skar

Collect LCI data by year from GREET1 as per correspondence file.

"""

#%%

import pandas as pd
import numpy as np
import xlwings as xw
from datetime import datetime

# %% Customize the Excel Instance

class ExcelApp(xw.App):
    """override xw.App default properties"""
    calculation = 'manual'
    display_alerts = False
    enable_events = False
    screen_updating = False
    visible = False
    
#%%
# GREET1 model run for study years

class GREET_LCI_import:
    
    def __init__(self, model_path_prefix, fmodel,
                 corr_path_prefix, fcorr_LCI, sheet_input, 
                 gparam, sheet_gparam, cell_gparam,
                 fsave_sim):
        
        self.model_path_prefix = model_path_prefix
        self.fmodel = fmodel        
        
        self.corr_path_prefix = corr_path_prefix
        self.fcorr_LCI = fcorr_LCI
        self.sheet_input = sheet_input
        
        self.gparam = gparam        
        self.sheet_gparam = sheet_gparam
        self.cell_gparam = cell_gparam
             
        self.fsave_sim = fsave_sim
        
        self.sim_df = pd.DataFrame() # initialize data frame to save runs
        
        # read parameter value sets
        self.sim_params = pd.read_excel(self.corr_path_prefix + '/' + self.fcorr_LCI, 
                                        self.sheet_input, header=0, index_col=None)
        
    
    def save_sim_to_file(self, mode, header):
        
        self.sim_df.to_csv(self.corr_path_prefix + '/' + self.fsave_sim, 
                           mode=mode, header=header, index=False)
    
    def sim_model(self):
        
        n_param_sets = self.sim_params.shape[1] - 3
        
        # truncate if file exists and create
        #self.save_sim_to_file(mode='w', header=False)
        write_header = True
        
        
        with ExcelApp() as app:            
            
            # open model workbook
            wb = xw.Book(model_path_prefix + '/' + fmodel)
        
            for gparam_index, gparam_val in enumerate(gparam):
                
                for param_set in range(0, n_param_sets):
                    
                    df_params = self.sim_params.iloc[:,[1,2,param_set+3]]
                    
                    print(f'Executing for sheet {self.sheet_input}, global parameter set: {gparam_val} and parameter set {param_set+1} out of {n_param_sets}') 
                    print( '    Elapsed time: ' + str(datetime.now() - init_time))
                    
                    # modify model with global parameters
                    for idx, val in enumerate(gparam_val):                
                        sheet = wb.sheets[self.sheet_gparam[idx]]
                        sheet[self.cell_gparam[idx]].value = val                
                    
                    # modify model with set of parameters            
                    for r in range(df_params.shape[0]):
                        if ~np.isnan(df_params.iloc[r,2]):
                            sheet = wb.sheets[df_params.iloc[r,0]]
                            sheet[df_params.iloc[r,1]].value = df_params.iloc[r,2]
                    
                    wb.app.calculate()
                               
                    # Extract data
                    sheet = wb.sheets['Algae_results']
                    self.sim_df = sheet['A1:E25'].options(pd.DataFrame).value 
                    
                    self.sim_df['gparam_val'] = '_'.join(map(str,gparam_val))
                    self.sim_df['sim_index'] = self.sim_params.columns[param_set+3]
                    self.sim_df['iter'] = param_set + 1                              
                                
                    check_time = datetime.now()
                    if write_header:
                        self.save_sim_to_file(mode='w', header=write_header) # append output to file
                        write_header = False
                    else:
                        self.save_sim_to_file(mode='a', header=write_header) # append output to file
                    print(f'File save time {str(datetime.now() - check_time)}')
        

if __name__ == '__main__':
    
    init_time = datetime.now()
    
    model_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - GREET sim/Proj_Algae/data/model/02-06-2024'
    fmodel = 'GREET_2022 Algae_harmonization_individual_approach_2050.xlsm'
    
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
   
    corr_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - GREET sim/Proj_Algae/data/correspondence_files/02-06-2024'
    fcorr_LCI = 'corr_LCI_GREET_Algae_individual_2050.xlsx'
    sheets_input = ['PC_disp', 
                    'PC_mass',
                    'PC_proc_alloc',
                    'Fuel'
                    ]   
    
    for sheet_1 in sheets_input:
        
        fsave_sim = 'GREET_Algae_sims_' + sheet_1 + '_02_06_2024' + '.csv'         
        
        obj = GREET_LCI_import(model_path_prefix, fmodel,
                               corr_path_prefix, fcorr_LCI, sheet_1,
                               gparam, sheet_gparam, cell_gparam,
                               fsave_sim)
        
        obj.sim_model()
    
    print( '    Total run time: ' + str(datetime.now() - init_time)) 
