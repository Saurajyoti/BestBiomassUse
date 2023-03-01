# -*- coding: utf-8 -*-
"""
Created on Tue Feb 28 08:29:28 2023

@author: skar
"""

import pandas as pd

path_in = 'C:/Users/skar/Box/saura_self/Proj - Algae/data/correspondence_files'
f_in = 'corr_LCI_GREET_pathway_Algae_v1_run.xlsx'
sheet_PC = 'PC'

path_out = 'C:/Users/skar/Box/saura_self/Proj - Algae/data/correspondence_files/output files'
fs_out = ['GREET_Algae_sims_PC_[1,1].csv',
          'GREET_Algae_sims_PC_[2,1]_[3,1].csv',
          'GREET_Algae_sims_PC_[1,2].csv',
          'GREET_Algae_sims_PC_[2,2].csv']

f_formatted = 'GREET_Algae_sims_PC_sk.xlsx'

param_set_fs_out = [['1-1'],
                    ['2-1', '3-1'],
                    ['1-2'],
                    ['2-2']]

df_all = pd.DataFrame()
for idx in range(len(fs_out)):
    df = pd.read_csv(path_out+'/'+fs_out[idx])
    df = df.loc[df['gparam_val'].isin(param_set_fs_out[idx]), : ]
    df_all = pd.concat([df, df_all]).reset_index(drop=True)
    
df_all[['disp_sc', 'FU_match']] = df_all['gparam_val'].str.split('-', expand=True)
df_all.drop(columns=['gparam_val'], inplace=True)

df_in = pd.read_excel(path_in+'/'+f_in, sheet_name=sheet_PC, header=3)
df_in = pd.DataFrame({'sim_index' : df_in.columns[3:],
                       'run_index' : range(1, len(df_in.columns[3:])+1,1)}) 

df_all = df_all.merge(df_in, how='left', left_on='param_set',
                      right_on = 'run_index').reset_index(drop=True)   
df_all.drop(columns=['param_set', 'run_index'], inplace=True)

df_all.to_excel(path_out+'/'+f_formatted, index=False)