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

input_path_prefix = 'C:\\Users\\skar\\Box\\saura_self\\Proj - Best use of biomass\\data'
output_path_prefix = 'C:\\Users\\skar\\Box\\saura_self\\Proj - Best use of biomass\\data\\interim'

input_path_TEA = input_path_prefix + '\\TEA'
input_path_LCA = input_path_prefix + '\\LCA'

f_TEA = 'TEA Database_09_09_2022.xlsx'
sheet_TEA = 'Biofuel'

f_out_itemized_mfsp = 'mfsp_itemized.csv'
f_out_agg_mfsp = 'mfsp_agg.csv'

save_interim_files = True

#%%
# Step: Load data file and select columns for computation

df_econ = pd.read_excel(input_path_TEA + '\\' + f_TEA, sheet_name = sheet_TEA, header = 3)

df_econ = df_econ[['Case/Scenario', 'Parameter',
       'Item', 'Stream Description', 'Flow Name', 'Flow: Units (numerator)',
       'Flow: Units (denominator)', 'Flow', 'Cost Item',
       'Cost: Units (numerator)', 'Cost: Units (denominator)', 'Unit Cost',
       'Operating Time: Units', 'Operating Time', 'Operating Time (%)',
       'Total Cost: Units (numerator)', 'Total Cost: Units (denominator)',
       'Total Cost', 'Total Flow: Units (numerator)',
       'Total Flow: Units (denominator)', 'Total Flow', 'Cost Year']]

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

# Step: Correct for inflatino to the year of study

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
    cost_items.to_csv(output_path_prefix + '\\' + f_out_itemized_mfsp)
    MFSP_agg.to_csv(output_path_prefix + '\\' + f_out_agg_mfsp)

#%%

# Step: Calculate MAC by Cost Items

numeric_cols = ['Flow', 'Unit Cost', 'Feedstock Flow', 'Operating Time', 'Biofuel Flow']
for col_name in numeric_cols:
    cost_items.loc[cost_items[col_name] == '-', col_name] = '0'

cost_items[numeric_cols] = cost_items[numeric_cols].apply(pd.to_numeric)

# (lb/hr) * (usd/lb) / (US dry ton/yr) * (hr/yr) / (GGE/US dry ton)
cost_items['MAC Value'] = cost_items['Flow'] * cost_items['Unit Cost'] / cost_items['Feedstock Flow'] * cost_items['Operating Time'] /  cost_items['Biofuel Flow']

# Aggregrate MAC for each feedstock-biofuel conversion pathways
cost_items_agg = cost_items.groupby(['Case/Scenario', 'Feedstock', 'Biofuel Flow Name']).agg({'MAC Value' : 'sum'}).reset_index()

# Save interim data tables
if save_interim_files == True:
    cost_items.to_csv(output_path_prefix + '\\' + f_out_itemized_mfsp)
    cost_items_agg.to_csv(output_path_prefix + '\\' + f_out_mfsp)
    
#%%