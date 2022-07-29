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

# Declare data input and other parameters

input_path_prefix = 'C:\\Users\\skar\\Box\\saura_self\\Proj - Best use of biomass\\data'
input_path_TEA = input_path_prefix + '\\TEA'
input_path_LCA = input_path_prefix + '\\LCA'

f_TEA = 'TEA Database_07_27_2022.xlsx'
sheet_TEA = 'Biofuel'


#%%

# import packages
import pandas as pd
import numpy as np
import os
from datetime import datetime

#%%

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

# Subset cost items to use for itemized MFSP calculation
cost_items = df_econ.loc[df_econ['Item'].isin(['Purchased Inputs',
                                               'Waste Disposal']), : ].copy()
cost_items.drop_duplicates(inplace=True)

# Separate feedstock demand flows
cost_feedstocks = df_econ.loc[df_econ['Item'] == 'Feedstock Cost', 
                         ['Case/Scenario', 'Stream Description', 'Flow Name', 'Flow: Units (numerator)', 'Flow: Units (denominator)', 'Flow']].copy()
cost_feedstocks.rename(columns={'Stream Description' : 'Feedstock Stream Description',
                                'Flow Name' : 'Feedstock',
                                'Flow: Units (numerator)' : 'Feedstock Flow: Units (numerator)', 
                                'Flow: Units (denominator)' : 'Feedstock Flow: Units (denominator)',
                                'Flow' : 'Feedstock Flow'}, inplace=True)
cost_feedstocks.drop_duplicates(inplace=True)

# Merge with the cost items df
cost_items = pd.merge(cost_items, cost_feedstocks, how='left', on='Case/Scenario').reset_index(drop=True)

#%%
# Separate biofuel yield flows
biofuel_yield = df_econ.loc[df_econ['Item'] == 'Final Product',
                            ['Case/Scenario', 'Flow Name', 'Flow: Units (numerator)', 'Flow: Units (denominator)', 'Flow']].copy()
biofuel_yield.rename(columns={'Flow Name' : 'Biofuel Flow Name',
                              'Flow: Units (numerator)' : 'Biofuel Flow: Units (numerator)', 
                              'Flow: Units (denominator)' : 'Biofuel Flow: Units (denominator)',
                              'Flow' : 'Biofuel Flow'}, inplace=True)
biofuel_yield.drop_duplicates(inplace=True)

# Select one flow unit for specific pathways
tmp_case = '2013 Biochemical Design Case: Corn Stover-Derived Sugars to Diesel'
tmp_flow = 'Renewable Diesel Blendstock'
tmp_unit = 'GGE'
tempdf = biofuel_yield.loc[(biofuel_yield['Case/Scenario'] == tmp_case) &
                           (biofuel_yield['Biofuel Flow Name'] == tmp_flow), : ].copy()
biofuel_yield = biofuel_yield.loc[(biofuel_yield['Case/Scenario'] != tmp_case) |
                                  (biofuel_yield['Biofuel Flow Name'] != tmp_flow), : ].copy()
tempdf = tempdf.loc[(tempdf['Case/Scenario'] == tmp_case) &
                    (tempdf['Biofuel Flow Name'] == tmp_flow) &
                    (tempdf['Biofuel Flow: Units (numerator)'] == tmp_unit), : ]
biofuel_yield = pd.concat([biofuel_yield, tempdf], axis=0).reset_index(drop=True)

tmp_case = '2022 Target Case'
tmp_flow = 'High-Octane Gasoline Blendsock'
tmp_unit = 'GGE'
tempdf = biofuel_yield.loc[(biofuel_yield['Case/Scenario'] == tmp_case) &
                           (biofuel_yield['Biofuel Flow Name'] == tmp_flow), : ].copy()
biofuel_yield = biofuel_yield.loc[(biofuel_yield['Case/Scenario'] != tmp_case) |
                                  (biofuel_yield['Biofuel Flow Name'] != tmp_flow), : ].copy()
tempdf = tempdf.loc[(tempdf['Case/Scenario'] == tmp_case) &
                    (tempdf['Biofuel Flow Name'] == tmp_flow) &
                    (tempdf['Biofuel Flow: Units (numerator)'] == tmp_unit), : ]
biofuel_yield = pd.concat([biofuel_yield, tempdf], axis=0).reset_index(drop=True)


# Merge with the cost items df
cost_items = pd.merge(cost_items, biofuel_yield, how='left', on='Case/Scenario').reset_index(drop=True)

#%%

# Calculate MAC by Cost Item

numeric_cols = ['Flow', 'Unit Cost', 'Feedstock Flow', 'Operating Time', 'Biofuel Flow']
for col_name in numeric_cols:
    cost_items.loc[cost_items[col_name] == '-', col_name] = '0'

cost_items[numeric_cols] = cost_items[numeric_cols].apply(pd.to_numeric)

# (lb/hr) * (usd/lb) / (US dry ton/yr) * (hr/yr) / (GGE/US dry ton)
cost_items['MAC Value'] = cost_items['Flow'] * cost_items['Unit Cost'] / cost_items['Feedstock Flow'] * cost_items['Operating Time'] /  cost_items['Biofuel Flow']

cost_items_agg = cost_items.groupby(['Case/Scenario', 'Feedstock', 'Biofuel Flow Name']).agg({'MAC Value' : 'sum'}).reset_index()