# -*- coding: utf-8 -*-
"""
Created on Wed Jan  4 10:32:14 2023

@author: skar
"""


'''
This is the main script to call data processing scripts, process data, perform calculations, and 
save output files. This version of main implements itemized LCA assessment of biofuel pathways.
'''

#%%
# Declare data input and other parameters

code_path_prefix = 'C:/Users/skar/repos/BestBiomassUse' # psth to the Github local repository

input_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - Best use of biomass/data'
output_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - Best use of biomass/data/interim'

input_path_TEA = input_path_prefix + '/TEA'
input_path_LCA = input_path_prefix + '/LCA'
input_path_GREET = input_path_prefix + '/GREET'
input_path_EIA_price = input_path_prefix + '/EIA'
input_path_corr = input_path_prefix + '/correspondence_files'
input_path_units = input_path_prefix + '/Units'

f_TEA = 'TEA Database_12_15_2022.xlsx'
sheet_TEA = 'Biofuel'

f_out_itemized_mfsp = 'mfsp_itemized.csv'
f_out_agg_mfsp = 'mfsp_agg.csv'
f_out_MAC = 'mac.csv'

f_EIA_price = 'EIA Dataset-energy_price-.csv'

f_GREET_efs = 'GREET_EF_EERE.csv'

# declare correspondence files
f_corr_replaced_replacing_fuel = 'corr_replaced_replacing_fuel.csv'
f_corr_fuel_replaced_GREET_pathway = 'corr_fuel_replaced_GREET_pathway.csv'
f_corr_fuel_replacing_GREET_pathway = 'corr_fuel_replacing_GREET_pathway.csv'
f_corr_GGE_GREET_fuel_replaced = 'corr_GGE_GREET_fuel_replaced.csv'
f_corr_GGE_GREET_fuel_replacing = 'corr_GGE_GREET_fuel_replacing.csv'
f_corr_itemized_LCI = 'corr_LCI_GREET_pathway.xlsx'

sheet_corr_itemized_LCI = 'GREET_mappings'

save_interim_files = True
 

#%%
# import packages
import pandas as pd
import numpy as np
import os
from datetime import datetime
import cpi

# Import user defined modules
os.chdir(code_path_prefix)

from unit_conversions import model_units

#%%
# Step: Load data file and select columns for computation

df_econ = pd.read_excel(input_path_TEA + '/' + f_TEA, sheet_name = sheet_TEA, header = 3, index_col=None)

df_econ = df_econ[['Case/Scenario', 'Parameter',
       'Item', 'Stream Description', 'Flow Name', 'Flow: Units (numerator)',
       'Flow: Units (denominator)', 'Flow', 'Cost Item',
       'Cost: Units (numerator)', 'Cost: Units (denominator)', 'Unit Cost',
       'Operating Time: Units', 'Operating Time', 'Operating Time (%)',
       'Total Cost: Units (numerator)', 'Total Cost: Units (denominator)',
       'Total Cost', 'Total Flow: Units (numerator)',
       'Total Flow: Units (denominator)', 'Total Flow', 'Cost Year']]

EIA_price = pd.read_csv(input_path_EIA_price + '/' + f_EIA_price, index_col=None)

ef = pd.read_csv(input_path_GREET + '/' + f_GREET_efs, header = 3, index_col=None).drop_duplicates()

# Unit conversion class object
ob_units = model_units(input_path_units, input_path_GREET, input_path_corr)

# load correspondence files
corr_replaced_replacing_fuel = pd.read_csv(input_path_corr + '/' + f_corr_replaced_replacing_fuel, header=3, index_col=None)
corr_fuel_replaced_GREET_pathway = pd.read_csv(input_path_corr + '/' + f_corr_fuel_replaced_GREET_pathway, header=3, index_col=None)
corr_fuel_replacing_GREET_pathway = pd.read_csv(input_path_corr + '/' + f_corr_fuel_replacing_GREET_pathway, header=3, index_col=None)
corr_GGE_GREET_fuel_replaced = pd.read_csv(input_path_corr + '/' + f_corr_GGE_GREET_fuel_replaced, header=3, index_col=None)
corr_GGE_GREET_fuel_replacing = pd.read_csv(input_path_corr + '/' + f_corr_GGE_GREET_fuel_replacing, header=3, index_col=None)
corr_itemized_LCA = pd.read_excel(input_path_corr + '/' + f_corr_itemized_LCI, sheet_corr_itemized_LCI, header=3, index_col=None)


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
# Create unique list of Items: {Purchased Inputs, Coproducts, and Waste Disposal}

lci_items = cost_items.loc[cost_items['Item'].isin(['Coproducts', 'Purchased Inputs', 'Waste Disposal']), 
                           ['Item', 'Stream Description', 'Flow Name']].drop_duplicates()

#%%

# Step: Create Biofuel Yield table and merge with Cost Item table

# Separate biofuel yield flows
biofuel_yield = df_econ.loc[df_econ['Item'] == 'Final Product',
                            ['Case/Scenario', 'Flow Name', 'Total Flow: Units (numerator)', 
                             'Total Flow: Units (denominator)', 'Total Flow']].reset_index(drop=True).copy()
biofuel_yield.rename(columns={'Flow Name' : 'Biofuel Flow Name',
                              'Total Flow: Units (numerator)' : 'Biofuel Flow: Units (numerator)', 
                              'Total Flow: Units (denominator)' : 'Biofuel Flow: Units (denominator)',
                              'Total Flow' : 'Biofuel Flow'}, inplace=True)

# For co-produced flows, summarize the flow data to one output
biofuel_yield2 = biofuel_yield.groupby(['Case/Scenario', 'Biofuel Flow: Units (numerator)', 
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
MFSP_agg.rename(columns={'Itemized MFSP' : 'MFSP_replacing fuel',
                         'Itemized MFSP: Units (numerator)' : 'MFSP_replacing fuel_Units (numerator)',
                         'Itemized MFSP: Units (denominator)' : 'MFSP_replacing fuel_Units (denominator)'}, inplace=True)

# Getting back the Final Product column
MFSP_agg = pd.merge(biofuel_yield[['Case/Scenario', 'Biofuel Flow Name']].drop_duplicates(), 
                    MFSP_agg, how='left', on='Case/Scenario').reset_index(drop=True)

# Save interim data tables
if save_interim_files == True:
    cost_items.to_csv(output_path_prefix + '/' + f_out_itemized_mfsp)
    MFSP_agg.to_csv(output_path_prefix + '/' + f_out_agg_mfsp)

#%%

# Step: Merge correspondence tables and itemized GREET emission factors

# Unit check for LCA data
corr_itemized_LCA[['Unit (numerator)', 'Unit (denominator)']] = corr_itemized_LCA['Unit'].str.split('/', expand=True)

corr_itemized_LCA.rename(columns={'GREET row names_level1' : 'Metric',
                                  'values_level1' : 'Value'}, inplace=True)
corr_itemized_LCA = corr_itemized_LCA[['Item',
                                       'Stream Description',
                                       'Flow Name',
                                       'Metric',
                                       'Value',
                                       'Unit (numerator)',
                                       'Unit (denominator)']]
## identify and filter rows needed for calculations
## harmonize units

LCA_items = df_econ.loc[df_econ['Item'].isin(['Purchased Inputs',
                                               'Waste Disposal',
                                               'Coproducts']), : ].copy()
# temporary value for production year
LCA_items['Production Year'] = 2021

pd.merge(LCA_items, corr_itemized_LCA, how='left', 
         left_on=['Item', 'Stream Description', 'Flow Name', 'Production Year'],
         right_on=['Item', 'Stream Description', 'Flow Name', 'Year']).reset_index(drop=True)

#%%

# Graphs to be created
# LCA (kgCO2e/MJ) vs pathways
# TEA ($/GGE) vs pathways
# MAC ($/kgCO2e) vs pathways
# LCA vs TEA (kg CO2e/ MJ vs $ per GGE)

# Four quard plot: ratio of ghg of alt fuels and the conv fuels vs. ratio of MFSP of alt fuels and conventional fuels