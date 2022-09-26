# -*- coding: utf-8 -*-
"""
Created on Wed Jul 27 11:34:10 2022

@author: skar
"""

'''
This is the main script to call data processing scripts, process data, perform calculations, and 
save output files. 
'''

#%%

# import packages
import pandas as pd
import numpy as np
import os
from datetime import datetime
import cpi

#%%
# Declare data input and other parameters

input_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - Best use of biomass/data'
output_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - Best use of biomass/data/interim'

input_path_TEA = input_path_prefix + '/TEA'
input_path_LCA = input_path_prefix + '/LCA'
input_path_GREET = input_path_prefix + '/GREET'
input_path_corr = input_path_prefix + '/correspondence_files'

f_TEA = 'TEA Database_09_09_2022.xlsx'
sheet_TEA = 'Biofuel'

f_out_itemized_mfsp = 'mfsp_itemized.csv'
f_out_agg_mfsp = 'mfsp_agg.csv'

f_GREET_efs = 'GREET_EF_EERE.csv'

# declare correspondence files
f_corr_ref_fuel_biofuel = 'corr_ref_fuel_biofuel.csv'
f_corr_fuel_replaced_GREET = 'corr_fuel_replaced_GREET.csv'
f_corr_biofuel_replacing_GREET = 'corr_biofuel_replacing_GREET.csv'

save_interim_files = True

#%%
# Step: Load data file and select columns for computation

df_econ = pd.read_excel(input_path_TEA + '/' + f_TEA, sheet_name = sheet_TEA, header = 3)

df_econ = df_econ[['Case/Scenario', 'Parameter',
       'Item', 'Stream Description', 'Flow Name', 'Flow: Units (numerator)',
       'Flow: Units (denominator)', 'Flow', 'Cost Item',
       'Cost: Units (numerator)', 'Cost: Units (denominator)', 'Unit Cost',
       'Operating Time: Units', 'Operating Time', 'Operating Time (%)',
       'Total Cost: Units (numerator)', 'Total Cost: Units (denominator)',
       'Total Cost', 'Total Flow: Units (numerator)',
       'Total Flow: Units (denominator)', 'Total Flow', 'Cost Year']]

ef = pd.read_csv(input_path_GREET + '\\' + f_GREET_efs, header = 3).drop_duplicates()

corr_ref_fuel_biofuel = pd.read_csv(input_path_corr + '/' + f_corr_ref_fuel_biofuel, header = 3, index_col=None)
corr_fuel_replaced_GREET = pd.read_csv(input_path_corr + '/' + f_corr_fuel_replaced_GREET, header = 3, index_col=None)
corr_biofuel_replacing_GREET = pd.read_csv(input_path_corr + '/' + f_corr_biofuel_replacing_GREET, header = 3, index_col=None)

#%%

# Step: Create Cost Item table

# Subset cost items to use for itemized MFSP calculation
cost_items = df_econ.loc[df_econ['Item'].isin(['Purchased Inputs',
                                               'Waste Disposal',
                                               'Coproducts',
                                               'Fixed Costs',
                                               'Capital Depreciation',
                                               'Average Income Tax',
                                               'Average Return on Investment']), : ].copy()
cost_items.drop_duplicates(inplace=True)

# Separate feedstock demand yearly flows
cost_feedstocks = df_econ.loc[df_econ['Item'] == 'Feedstock Cost', 
                             ['Case/Scenario', 'Stream Description', 'Flow Name', 
                              'Flow: Units (numerator)', 'Flow: Units (denominator)', 'Flow']].copy()
cost_feedstocks.rename(columns={'Stream Description' : 'Feedstock Stream Description',
                                'Flow Name' : 'Feedstock',
                                'Flow: Units (numerator)' : 'Feedstock Flow: Units (numerator)', 
                                'Flow: Units (denominator)' : 'Feedstock Flow: Units (denominator)',
                                'Flow' : 'Feedstock Flow'}, inplace=True)
#cost_feedstocks.drop_duplicates(inplace=True)

# Merge with the cost items df
cost_items = pd.merge(cost_items, cost_feedstocks, how='left', on='Case/Scenario').reset_index(drop=True)

#%%

# Step: Create Biofuel Yield table and merge with Cost Item table

# Separate biofuel yield flows
biofuel_yield = df_econ.loc[df_econ['Item'] == 'Final Product',
                            ['Case/Scenario', 'Flow Name', 'Item', 'Total Flow: Units (numerator)', 
                             'Total Flow: Units (denominator)', 'Total Flow']].reset_index().copy()
biofuel_yield.rename(columns={'Flow Name' : 'Biofuel Flow Name',
                              'Total Flow: Units (numerator)' : 'Biofuel Flow: Units (numerator)', 
                              'Total Flow: Units (denominator)' : 'Biofuel Flow: Units (denominator)',
                              'Total Flow' : 'Biofuel Flow'}, inplace=True)

# For co-produced flows, summarize the flow data to one output
biofuel_yield2 = biofuel_yield.groupby(['Case/Scenario', 'Item', 'Biofuel Flow: Units (numerator)', 
                                        'Biofuel Flow: Units (denominator)']).agg({'Biofuel Flow' : 'sum'}).reset_index()

# Merge with the cost items df
cost_items = pd.merge(cost_items, biofuel_yield2, how='left', on='Case/Scenario').reset_index(drop=True)

#%%

# Step: Correct for inflation to the year of study

# drop blanks
cost_items = cost_items.loc[~cost_items['Total Cost'].isin(['-']), : ]

study_year = 2021

#cpi.update()

cost_items['Adjusted Total Cost'] = cost_items.apply(lambda x: cpi.inflate(x['Total Cost'], x['Cost Year'], to=study_year), axis=1)
cost_items['Adjusted Cost Year'] = study_year

#%%

# Step: Calculate itemized Marginal Fuel Selling Price (MFSP)

cost_items['Itemized MFSP'] = cost_items['Adjusted Total Cost'].astype(float) / cost_items['Biofuel Flow'].astype(float)
cost_items['Itemized MFSP: Units (numerator)'] = cost_items['Total Cost: Units (numerator)']
cost_items['Itemized MFSP: Units (denominator)'] = cost_items['Biofuel Flow: Units (numerator)']

MFSP_agg = cost_items.groupby(['Case/Scenario',
                               'Feedstock',
                               'Itemized MFSP: Units (numerator)', 
                               'Itemized MFSP: Units (denominator)',
                               'Adjusted Cost Year']).agg({'Itemized MFSP' : 'sum'}).reset_index()

# Getting back the Final Product column
MFSP_agg = pd.merge(biofuel_yield[['Case/Scenario', 'Biofuel Flow Name']].drop_duplicates(), 
                    MFSP_agg, how='left', on='Case/Scenario').reset_index(drop=True)

# Save interim data tables
if save_interim_files == True:
    cost_items.to_csv(output_path_prefix + '/' + f_out_itemized_mfsp)
    MFSP_agg.to_csv(output_path_prefix + '/' + f_out_agg_mfsp)

#%%

# Step: Merge correspondence tables and GREET emission factors

# map conventional fuels those are replaced with biofuels
MAC_df = pd.merge(MFSP_agg, corr_ref_fuel_biofuel, how = 'left', 
               on=['Case/Scenario', 'Biofuel Flow Name', 'Feedstock']).reset_index(drop=True) 

# map biofuels with GREET pathways
MAC_df = pd.merge(MAC_df, corr_biofuel_replacing_GREET, how='left',
               on=['Case/Scenario', 'Biofuel Flow Name', 'Feedstock']).reset_index(drop=True)
MAC_df.rename(columns={'GREET Pathway' : 'GREET Pathway for replacing fuel'}, inplace=True)

# map conventional fuels with GREET pathways
MAC_df = pd.merge(MAC_df, corr_fuel_replaced_GREET, how='left', on=['Reference Case Fuel']).reset_index(drop=True)
MAC_df.rename(columns={'GREET Pathway' : 'GREET Pathway for replaced fuel'}, inplace=True)

# map GREET carbon intensities for replaced fuels
MAC_df = pd.merge(MAC_df, ef, how='left', 
                  left_on=['GREET Pathway for replaced fuel'],
                  right_on=['GREET Pathway']).reset_index(drop=True)
MAC_df.rename(columns={'Flow Name' : 'Flow Name_replaced fuel',
                       'Formula' :'Formula_replaced fuel',
                       'Unit (Numerator)' : 'Unit (Numerator)_replaced fuel',
                       'Unit (Denominator)' : 'Unit (Denominator)_replaced fuel',
                       'Case' : 'Case_replaced fuel',
                       'Scope' : 'Scope_replaced fuel',
                       'Reference case' : 'Reference case_replaced fuel',
                       'Elec0' : 'Elec0_replaced fuel'}, inplace=True)
MAC_df.drop(['GREET Pathway'], axis=1, inplace=True)

# map GREET carbon intensities for replacing fuels

# Considering reference case carbon intensities only
ef = ef.loc[ef['Case'] == 'Reference case', : ]

MAC_df = pd.merge(MAC_df, ef, how='left', 
                  left_on=['GREET Pathway for replacing fuel', 'Year'],
                  right_on=['GREET Pathway', 'Year']).reset_index(drop=True)
MAC_df.rename(columns={'Flow Name' : 'Flow Name_replacing fuel',
                       'Formula' :'Formula_replacing fuel',
                       '	Unit (Numerator)	' : 'Unit (Numerator)_replacing fuel',
                       'Unit (Denominator)' : 'Unit (Denominator)_replacing fuel',
                       'Case' : 'Case_replacing fuel',
                       'Scope' : 'Scope_replacing fuel',
                       'Reference case' : 'Reference case_replacing fuel',
                       'Elec0' : 'Elec0_replacing fuel'}, inplace=True)
MAC_df.drop(['GREET Pathway'], axis=1, inplace=True)


#%%
# Step: Calculate MAC by Cost Items

# MAC = (MFSP_biofuel - MFSP_ref) / (CI_ref - CI_biofuel)



numeric_cols = ['Flow', 'Unit Cost', 'Feedstock Flow', 'Operating Time', 'Biofuel Flow']
for col_name in numeric_cols:
    cost_items.loc[cost_items[col_name] == '-', col_name] = '0'

cost_items[numeric_cols] = cost_items[numeric_cols].apply(pd.to_numeric)

# (lb/hr) * (usd/lb) / (US dry ton/yr) * (hr/yr) / (GGE/US dry ton)
cost_items['MAC Value'] = cost_items['Flow'] * cost_items['Unit Cost'] / \
                            cost_items['Feedstock Flow'] * cost_items['Operating Time'] /  cost_items['Biofuel Flow']

# Aggregrate MAC for each feedstock-biofuel conversion pathways
cost_items_agg = cost_items.groupby(['Case/Scenario', 'Feedstock', 'Biofuel Flow Name']).agg({'MAC Value' : 'sum'}).reset_index()

# Save interim data tables
if save_interim_files == True:
    cost_items.to_csv(output_path_prefix + '/' + f_out_itemized_mfsp)
    cost_items_agg.to_csv(output_path_prefix + '/' + f_out_mfsp)
    
#%%