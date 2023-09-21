# -*- coding: utf-8 -*-
"""
Created on Tue Feb 28 08:29:28 2023

@author: skar
"""
# this script is not required to get back the simulation index, as its already
# stored in the output tables in GREET_LCI_import.py

import pandas as pd

path_in = 'C:/Users/skar/Box/saura_self/Proj - Algae/data/correspondence_files/09-19-2023'
f_in = 'corr_LCI_GREET_Algae_individual.xlsx'
sheet_PC = 'Fuel'

path_out = 'C:/Users/skar/Box/saura_self/Proj - Algae/data/output/09-19-2023'
fs_out = [
    'GREET_Algae_sims_PC_proc_alloc_09_20_2023.csv',
    'GREET_Algae_sims_PC_mass_09_20_2023.csv',
    'GREET_Algae_sims_PC_disp_09_20_2023.csv',
    'GREET_Algae_sims_Fuel_09_20_2023.csv'
    ]

f_formatted = 'GREET_Algae_sims_sk.xlsx'

param_set_fs_out = [['1_1',
                     #'2-1', '3-1',
                     #'1-2', '2-2', '3,2',
                     #'1,3', '2,3', '3,3'
                     ],
                    ['1_1'],
                    ['1_1'],
                    ['1_1']]

df_all = pd.DataFrame()
for idx in range(len(fs_out)):
    df = pd.read_csv(path_out+'/'+fs_out[idx])
    df['gparam_val'] = df['gparam_val'].astype('string')
    df = df.loc[df['gparam_val'].isin(param_set_fs_out[idx]), : ]
    df_all = pd.concat([df, df_all]).reset_index(drop=True)
    
df_all[['disp_sc', 'FU_match']] = df_all['gparam_val'].str.split('_', expand=True)
df_all.drop(columns=['gparam_val'], inplace=True)

df_in = pd.read_excel(path_in+'/'+f_in, sheet_name=sheet_PC, header=0)
df_in = pd.DataFrame({'sim_index' : df_in.columns[3:],
                       'run_index' : range(1, len(df_in.columns[3:])+1,1)}) 

df_all = df_all.merge(df_in, how='left', left_on='param_set',
                      right_on = 'run_index').reset_index(drop=True)   
df_all.drop(columns=['param_set', 'run_index'], inplace=True)

df_all.to_excel(path_out+'/'+f_formatted, index=False)