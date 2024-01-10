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
    display_alerts = True
    enable_events = True
    screen_updating = True
    visible = True
    
#%%
# GREET1 model run for study years

class GREET_LCI_import:
    
    def __init__(self, model_path_prefix, fmodel,
                 corr_path_prefix, fcorr_LCI, if_params, sheet_input, 
                 gparam, sheet_gparam, cell_gparam,
                 out_path_prefix, fsave_sim):
        
        self.model_path_prefix = model_path_prefix
        self.fmodel = fmodel        
        
        self.corr_path_prefix = corr_path_prefix
        self.fcorr_LCI = fcorr_LCI
        
        
        self.gparam = gparam        
        self.sheet_gparam = sheet_gparam
        self.cell_gparam = cell_gparam
             
        self.out_path_prefix = out_path_prefix
        self.fsave_sim = fsave_sim
        
        self.sim_df = pd.DataFrame() # initialize data frame to save runs
        
        # read parameter value sets
        self.if_params = if_params
        if self.if_params:
            self.sheet_input = sheet_input
            self.sim_params = pd.read_excel(self.corr_path_prefix + '/' + self.fcorr_LCI, 
                                            self.sheet_input, header=0, index_col=None)
        
    
    def save_sim_to_file(self, mode, header):
        
        fname = self.out_path_prefix + '/' + self.fsave_sim
        
        if mode == 'a':
            with pd.ExcelWriter(fname, mode=mode, if_sheet_exists='overlay') as writer:  
                self.sim_df.to_excel(writer, sheet_name='Sheet1', header=header, index=False)
        else:
            with pd.ExcelWriter(fname, mode=mode) as writer:  
                self.sim_df.to_excel(writer, sheet_name='Sheet1', header=header, index=False)
            
    def sim_model(self):
        
        
        # truncate if file exists and create
        #self.save_sim_to_file(mode='w', header=False)
        write_header = True
        
        if self.if_params:
            
            n_param_sets = self.sim_params.shape[1] - 3
            
            with ExcelApp() as app:            
                
                # open model workbook
                wb = xw.Book(model_path_prefix + '/' + fmodel)
            
                for gparam_index, gparam_val in enumerate(gparam):
                    
                    for param_set in range(0, n_param_sets):
                        
                        df_params = self.sim_params.iloc[:,[1,2,param_set+3]]
                        
                        print(f'Executing for sheet {self.sheet_input}, global parameter set: {gparam_val} and parameter set {param_set+1} out of {n_param_sets}') 
                        #print(f'Executing for global parameter {gparam_index+1} out of {len(gparam)} global parameters') 
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
                        sheet = wb.sheets['Autonomie_results']
                        self.sim_df = sheet['A4:G30'].options(pd.DataFrame, index=False).value                         
                        
                        self.sim_df['gparam_val'] = '_'.join(map(str,gparam_val))
                       # self.sim_df['sim_index'] = self.sim_params.columns[param_set+3]
                       # self.sim_df['iter'] = param_set + 1                              
                                    
                        check_time = datetime.now()
                        if write_header:
                            self.save_sim_to_file(mode='w', header=write_header) # append output to file
                            write_header = False
                        else:
                            self.save_sim_to_file(mode='a', header=write_header) # append output to file
                        print(f'File save time {str(datetime.now() - check_time)}')
                #wb.save()
                wb.close()         
        else:
            
            with ExcelApp() as app:            
                
                # open model workbook
                wb = xw.Book(model_path_prefix + '/' + fmodel)
            
                for gparam_index, gparam_val in enumerate(gparam):
                   
                        print(f'Executing for global parameter {gparam_index+1} out of {len(gparam)} global parameters') 
                        print( '    Elapsed time: ' + str(datetime.now() - init_time))
                        
                        # modify model with global parameters
                        for idx, val in enumerate(gparam_val):                
                            sheet = wb.sheets[self.sheet_gparam[idx]]
                            sheet[self.cell_gparam[idx]].value = val                
                        
                        wb.app.calculate()                        
                                   
                        # Extract data
                        sheet = wb.sheets['Autonomie_results']
                        self.sim_df = sheet['A4:H82'].options(pd.DataFrame, index=False).value 
                        print(self.sim_df['Feedstock'])
                        
                        self.sim_df['gparam_val'] = '_'.join(map(str,gparam_val))
                                                           
                        check_time = datetime.now()
                        if write_header:
                            self.save_sim_to_file(mode='w', header=write_header) # append output to file
                            write_header = False
                        else:
                            self.save_sim_to_file(mode='a', header=write_header) # append output to file
                        print(f'File save time {str(datetime.now() - check_time)}')
                #wb.save()
                wb.close() 


if __name__ == '__main__':
    
    init_time = datetime.now()
    
    model_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - GREET sim/Proj_Autonomie/data/model/01-09-2024'
    # Make sure macros are enabled in Excel, after opening the GREET1 file, under Development/Macro settings
    fmodel = 'GREET1_2023.xlsm'
    
    # Global parameter declarations
    sheet_gparam = ['Inputs', 'Inputs', 'Inputs'] # the sheets in fmodel that has the parameters
    cell_gparam = ['E9', 'F746', 'F747'] # the cells in fmodel sheet_gparam where parameters are located
    gparam = [[2005,1,1],  # value set of global parameters
              [2005,2,2],
              [2005,3,3],
              [2005,4,4],
              [2005,5,5],
              [2005,6,6],
              [2005,7,7],
              [2005,8,8],
              [2005,9,9],
              [2005,10,10],
              [2005,11,11],
              [2005,12,12],
              
                [2010,1,1],  
                [2010,2,2],
                [2010,3,3],
                [2010,4,4],
                [2010,5,5],
                [2010,6,6],
                [2010,7,7],
                [2010,8,8],
                [2010,9,9],
                [2010,10,10],
                [2010,11,11],
                [2010,12,12],
                  
                [2013,1,1],  
                [2013,2,2],
                [2013,3,3],
                [2013,4,4],
                [2013,5,5],
                [2013,6,6],
                [2013,7,7],
                [2013,8,8],
                [2013,9,9],
                [2013,10,10],
                [2013,11,11],
                [2013,12,12],      
                  
                [2014,1,1],  
                [2014,2,2],
                [2014,3,3],
                [2014,4,4],
                [2014,5,5],
                [2014,6,6],
                [2014,7,7],
                [2014,8,8],
                [2014,9,9],
                [2014,10,10],
                [2014,11,11],
                [2014,12,12], 
                
                [2015,1,1],  
                [2015,2,2],
                [2015,3,3],
                [2015,4,4],
                [2015,5,5],
                [2015,6,6],
                [2015,7,7],
                [2015,8,8],
                [2015,9,9],
                [2015,10,10],
                [2015,11,11],
                [2015,12,12],  
                
                [2016,1,1],  
                [2016,2,2],
                [2016,3,3],
                [2016,4,4],
                [2016,5,5],
                [2016,6,6],
                [2016,7,7],
                [2016,8,8],
                [2016,9,9],
                [2016,10,10],
                [2016,11,11],
                [2016,12,12],  
                
                [2017,1,1],  
                [2017,2,2],
                [2017,3,3],
                [2017,4,4],
                [2017,5,5],
                [2017,6,6],
                [2017,7,7],
                [2017,8,8],
                [2017,9,9],
                [2017,10,10],
                [2017,11,11],
                [2017,12,12], 
                
                [2018,1,1],  
                [2018,2,2],
                [2018,3,3],
                [2018,4,4],
                [2018,5,5],
                [2018,6,6],
                [2018,7,7],
                [2018,8,8],
                [2018,9,9],
                [2018,10,10],
                [2018,11,11],
                [2018,12,12],  
                
                [2019,1,1],  
                [2019,2,2],
                [2019,3,3],
                [2019,4,4],
                [2019,5,5],
                [2019,6,6],
                [2019,7,7],
                [2019,8,8],
                [2019,9,9],
                [2019,10,10],
                [2019,11,11],
                [2019,12,12],  
                
                [2020,1,1],  
                [2020,2,2],
                [2020,3,3],
                [2020,4,4],
                [2020,5,5],
                [2020,6,6],
                [2020,7,7],
                [2020,8,8],
                [2020,9,9],
                [2020,10,10],
                [2020,11,11],
                [2020,12,12], 
                  
                [2021,1,1],  
                [2021,2,2],
                [2021,3,3],
                [2021,4,4],
                [2021,5,5],
                [2021,6,6],
                [2021,7,7],
                [2021,8,8],
                [2021,9,9],
                [2021,10,10],
                [2021,11,11],
                [2021,12,12], 
                  
                [2023,1,1],  
                [2023,2,2],
                [2023,3,3],
                [2023,4,4],
                [2023,5,5],
                [2023,6,6],
                [2023,7,7],
                [2023,8,8],
                [2023,9,9],
                [2023,10,10],
                [2023,11,11],
                [2023,12,12], 
                
                [2025,1,1],  
                [2025,2,2],
                [2025,3,3],
                [2025,4,4],
                [2025,5,5],
                [2025,6,6],
                [2025,7,7],
                [2025,8,8],
                [2025,9,9],
                [2025,10,10],
                [2025,11,11],
                [2025,12,12], 
                  
                [2030,1,1],  
                [2030,2,2],
                [2030,3,3],
                [2030,4,4],
                [2030,5,5],
                [2030,6,6],
                [2030,7,7],
                [2030,8,8],
                [2030,9,9],
                [2030,10,10],
                [2030,11,11],
                [2030,12,12], 
                
                [2035,1,1],  
                [2035,2,2],
                [2035,3,3],
                [2035,4,4],
                [2035,5,5],
                [2035,6,6],
                [2035,7,7],
                [2035,8,8],
                [2035,9,9],
                [2035,10,10],
                [2035,11,11],
                [2035,12,12],
                  
                [2040,1,1],  
                [2040,2,2],
                [2040,3,3],
                [2040,4,4],
                [2040,5,5],
                [2040,6,6],
                [2040,7,7],
                [2040,8,8],
                [2040,9,9],
                [2040,10,10],
                [2040,11,11],
                [2040,12,12],
                
                [2045,1,1],  
                [2045,2,2],
                [2045,3,3],
                [2045,4,4],
                [2045,5,5],
                [2045,6,6],
                [2045,7,7],
                [2045,8,8],
                [2045,9,9],
                [2045,10,10],
                [2045,11,11],
                [2045,12,12],
                
                [2050,1,1],  
                [2050,2,2],
                [2050,3,3],
                [2050,4,4],
                [2050,5,5],
                [2050,6,6],
                [2050,7,7],
                [2050,8,8],
                [2050,9,9],
                [2050,10,10],
                [2050,11,11],
                [2050,12,12]
                    
              ]  
    
    # test set
    #sheet_gparam = ['Inputs', 'Inputs', 'Inputs']
    #cell_gparam = ['E9', 'F746', 'F747']
    #gparam = [[2005,1,1],  # value set of global parameters
    #          [2005,2,2],
    #          [2005,3,3],
    #          [2005,4,4]]
   
    corr_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - GREET sim/Proj_Autonomie/data/correspondence_files/01-09-2023'
    fcorr_LCI = 'corr_LCI_GREET_Autonomie.xlsx'
    sheets_input = ['param_update'] 
    if_params = False
    
    out_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - GREET sim/Proj_Autonomie/data/output'
    
    if if_params:        
        for sheet_1 in sheets_input:
            
            #fsave_sim = 'GREET_2023_' + sheet_1 + '_09_20_2023' + '.csv'         
            fsave_sim = 'GREET_2023_sim_out.xlsx'
            
            obj = GREET_LCI_import(model_path_prefix, fmodel,
                                   corr_path_prefix, fcorr_LCI, if_params, sheet_1,
                                   gparam, sheet_gparam, cell_gparam,
                                   out_path_prefix, fsave_sim)
            
            obj.sim_model()
    
    else:
        fsave_sim = 'GREET_2023_sim_out.xlsx'
        sheet_1=''
        obj = GREET_LCI_import(model_path_prefix, fmodel,
                               corr_path_prefix, fcorr_LCI, if_params, sheet_1,
                               gparam, sheet_gparam, cell_gparam,
                               out_path_prefix, fsave_sim)
        
        obj.sim_model()
    
    print( '    Total run time: ' + str(datetime.now() - init_time)) 
