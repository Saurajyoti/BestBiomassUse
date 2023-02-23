# -*- coding: utf-8 -*-
"""
Created on Wed Feb 22 16:50:36 2023

@author: skar
"""

# Calculate MAC for reported TEA and LCA studies

# Calculate percentage GHG difference, percentage MFSP difference

# Implement selection logic: if LCA_alternative pathway > LCA_conventional pathway: ignore the pathway, otherwise select pathway for acceptance ranking
#%%

input_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - Best use of biomass/data'

data_path_prefix = input_path_prefix + '/QA'
unit_path_prefix = input_path_prefix + '/Unit'

f_pathways = 'reported_TEA_LCA.xlsx'
sheet_lca = 'LCA'
sheet_tea = 'TEA'

f_save = 'mac_for_pathways.csv'

#%%

import pandas as pd

#%%

# pre-defined data structures

replaced_CI ={
    'gasoline' : 89.90, # g CO2e/MJ, GREET and Decarb 2b study
    'diesel' : 90.30, # g CO2e/MJ, GREET and Decarb 2b study
    'diesel, gasoline' : 90.30, # g CO2e/MJ, GREET and Decarb 2b study  
    'jet fuel' : 84.5, # g CO2e/MJ, GREET and Decarb 2b study
    'naptha, jet fuel' : 84.5, # g CO2e/MJ, GREET and Decarb 2b study
    'plastic' : 2600 # # g CO2e/kg, fossil derived HDPE plastic, Benavides et al., 2020
    }
replaced_CI_unit_numerator ={
    'gasoline' : 'g CO2e', 
    'diesel' : 'g CO2e', 
    'diesel, gasoline' : 'g CO2e', 
    'jet fuel' : 'g CO2e', 
    'naptha, jet fuel' : 'g CO2e', 
    'plastic' : 'g CO2e' 
    }
replaced_CI_unit_denominator ={
    'gasoline' : 'MJ', 
    'diesel' : 'MJ', 
    'diesel, gasoline' : 'MJ', 
    'jet fuel' : 'MJ', 
    'naptha, jet fuel' : 'MJ', 
    'plastic' : 'kg'
    }

replaced_mfsp ={
    'gasoline' : 0.03, # g CO2e/MJ, GREET and Decarb 2b study
    'diesel' : 0.03, # g CO2e/MJ, GREET and Decarb 2b study
    'diesel, gasoline' : 0.03, # g CO2e/MJ, GREET and Decarb 2b study  
    'jet fuel' : 0.02, # g CO2e/MJ, GREET and Decarb 2b study
    'naptha, jet fuel' : 0.02, # g CO2e/MJ, GREET and Decarb 2b study
    'plastic' : 0.81 # # g CO2e/kg, fossil derived HDPE plastic, Benavides et al., 2020
    }
replaced_mfsp_unit_numerator ={
    'gasoline' : 'USD', 
    'diesel' : 'USD', 
    'diesel, gasoline' : 'USD', 
    'jet fuel' : 'USD', 
    'naptha, jet fuel' : 'USD', 
    'plastic' : 'USD' 
    }
replaced_mfsp_unit_denominator ={
    'gasoline' : 'MJ', 
    'diesel' : 'MJ', 
    'diesel, gasoline' : 'MJ', 
    'jet fuel' : 'MJ', 
    'naptha, jet fuel' : 'MJ', 
    'plastic' : 'kg'
    }

#%%

lca = pd.read_excel(data_path_prefix+'/'+f_pathways, sheet_name=sheet_lca, index_col=None)
tea = pd.read_excel(data_path_prefix+'/'+f_pathways, sheet_name=sheet_tea, index_col=None)
tea = tea[['Links to reports', 'Case/Scenario', 'MFSP',       
           'MFSP_unit_numerator', 'MFSP_unit_denominator', 'MFSP_Year']]

# map TEA and LCA
mac = pd.concat([lca.merge(tea, how='inner', 
                           left_on='TEA_mapping_1', right_on='Case/Scenario').reset_index(drop=True),
                lca.merge(tea, how='inner',
                          left_on='TEA_mapping_2', right_on='Case/Scenario').reset_index(drop=True)]).reset_index(drop=True)

# map CI of replaced commodities
mac['CI_replaced'] = [replaced_CI[x] for x in mac['Fuel pool']]
mac['CI_replaced_unit_numerator'] = [replaced_CI_unit_numerator[x] for x in mac['Fuel pool']]
mac['CI_replaced_unit_denominator'] = [replaced_CI_unit_denominator[x] for x in mac['Fuel pool']] 

# map mfsp of replaced commodities
mac['mfsp_replaced'] = [replaced_mfsp[x] for x in mac['Fuel pool']]
mac['mfsp_replaced_unit_numerator'] = [replaced_mfsp_unit_numerator[x] for x in mac['Fuel pool']]
mac['mfsp_replaced_unit_denominator'] = [replaced_mfsp_unit_denominator[x] for x in mac['Fuel pool']] 

# ad-hoc unit convertion, GGE to MJ
mac.loc[mac['MFSP_unit_denominator'].isin(['GGE']), 'MFSP'] =\
    mac.loc[mac['MFSP_unit_denominator'].isin(['GGE']), 'MFSP'] / 121.2  # 1 GGE = 121.3 MJ
mac.loc[mac['MFSP_unit_denominator'].isin(['GGE']), 'MFSP_unit_denominator'] = 'MJ'

# unit check
inconsistent_units = (mac['MFSP_unit_numerator'] != mac['mfsp_replaced_unit_numerator']) |\
                        (mac['MFSP_unit_denominator'] != mac['mfsp_replaced_unit_denominator']) |\
                        (mac['CI_unit_numerator'] != mac['CI_unit_numerator']) |\
                        (mac['CI_unit_denominator'] != mac['CI_unit_denominator']) |\
                        (mac['MFSP_unit_denominator'] != mac['CI_unit_denominator'])
mac_inconsistent_units = mac.loc[inconsistent_units, :]
if sum(inconsistent_units):
    print('The following pathways are ignored as units are not consistent for MAC calculation:')
    print(mac_inconsistent_units)
mac = mac.loc[~inconsistent_units, :].copy()                     

# calculate percentage carbon intensity abatement by alternative pathways
# (CI_replaced - CI_replacing)/CI_replaced * 100
mac['percent_CI_abated'] = (mac['CI_replaced']-mac['CI'])/mac['CI_replaced']*100

# filter out alternative pathways with higher environmental impact than conventional pathways
not_effective = mac.loc[mac['CI'] > mac['CI_replaced'], : ]
mac = mac.loc[mac['CI'] <= mac['CI_replaced'], : ]

if not_effective.shape[0]>0:
    print('The following pathways have higher carbon intensity than conventional pathways, hence not considered as feasible options:')
    print(not_effective)

# calculate MAC
mac['mac'] = (mac['MFSP']-mac['mfsp_replaced']) / (mac['CI_replaced']-mac['CI'])
mac['mac_unit_numerator'] = mac['MFSP_unit_numerator']
mac['mac_unit_denominator'] = mac['CI_unit_numerator']

# ad-hoc unit convert g to kg
mac.loc[mac['mac_unit_denominator'].isin(['g CO2e']), 'mac'] =\
    mac.loc[mac['mac_unit_denominator'].isin(['g CO2e']), 'MFSP'] *1E3  # g to kg
mac.loc[mac['mac_unit_denominator'].isin(['g CO2e']), 'mac_unit_denominator'] = 'kg'

# save to file
mac.to_csv(data_path_prefix+'/'+f_save, index=False)