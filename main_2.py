# -*- coding: utf-8 -*-
"""
Created on Wed Jan  4 10:32:14 2023

@author: Saurajyoti Kar, Argonne National Laboratory
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

f_TEA = 'TEA Database_01_21_2023.xlsx'
sheet_TEA = 'Biofuel'

f_out_itemized_mfsp = 'mfsp_itemized.csv'
f_out_agg_mfsp = 'mfsp_agg.csv'
f_out_itemized_LCA = 'lca_itemized.csv'
f_out_agg_LCA = 'lca_agg.csv'
f_out_MAC = 'mac.csv'

f_EIA_price = 'EIA Dataset-energy_price-.csv'

f_GREET_efs = 'GREET_EF_EERE.csv'

# declare correspondence files
f_corr_replaced_replacing_fuel = 'corr_replaced_replacing_fuel.csv'
f_corr_fuel_replaced_GREET_pathway = 'corr_fuel_replaced_GREET_pathway.csv'
f_corr_fuel_replacing_GREET_pathway = 'corr_fuel_replacing_GREET_pathway.csv'
f_corr_GGE_GREET_fuel_replaced = 'corr_GGE_GREET_fuel_replaced.csv'
f_corr_GGE_GREET_fuel_replacing = 'corr_GGE_GREET_fuel_replacing.csv'
f_corr_itemized_LCI = 'corr_LCI_GREET_temporal.csv'

# Year of study, to which inflation will be adjusted
study_year = 2021

# Option to control cost credit for coproducts while calculating aggregrated MFSP
consider_coproduct_cost_credit = False

# Option to control emissions credit for coproducts while calculating aggregrated CIs
consider_coproduct_env_credit = False

save_interim_files = True
#%%
# import packages
import pandas as pd
import numpy as np
import os
from datetime import datetime
import cpi

#cpi.update()

# Import user defined modules
os.chdir(code_path_prefix)

from unit_conversions import model_units

#%%
# Step: Load data file and select columns for computation

init_time = datetime.now()

df_econ = pd.read_excel(input_path_TEA + '/' + f_TEA, sheet_name = sheet_TEA, header = 3, index_col=None)

df_econ = df_econ[['Case/Scenario', 'Parameter',
       'Item', 'Stream Description', 'Flow Name', 'Flow: Unit (numerator)',
       'Flow: Unit (denominator)', 'Flow', 'Cost Item',
       'Cost: Unit (numerator)', 'Cost: Unit (denominator)', 'Unit Cost',
       'Operating Time: Unit', 'Operating Time', 'Operating Time (%)',
       'Total Cost: Unit (numerator)', 'Total Cost: Unit (denominator)',
       'Total Cost', 'Total Flow: Unit (numerator)',
       'Total Flow: Unit (denominator)', 'Total Flow', 'Cost Year']]

# Temporarily filter df_econ for QA

df_econ = df_econ.loc[df_econ['Case/Scenario'].isin([#'2013 Biochemical Design Case: Corn Stover-Derived Sugars to Diesel',
                                                     #'2015 Biochemical Catalysis Design Report',
                                                     '2018 Biochemical Design Case: BDO Pathway',
                                                     '2018 Biochemical Design Case: Organic Acids Pathway',
                                                     #'2018, 2018 SOT High Octane Gasoline from Lignocellulosic Biomass via Syngas and Methanol/Dimethyl Ether Intermediates',
                                                     #'2018, 2022 projection High Octane Gasoline from Lignocellulosic Biomass via Syngas and Methanol/Dimethyl Ether Intermediates',
                                                     #'2020, 2019 SOT High Octane Gasoline from Lignocellulosic Biomass via Syngas and Methanol/Dimethyl Ether Intermediates',
                                                     #'2020, 2022 projection High Octane Gasoline from Lignocellulosic Biomass via Syngas and Methanol/Dimethyl Ether Intermediates',
                                                     'Biochemical 2019 SOT: Acids Pathway (Burn Lignin Case)',
                                                     'Biochemical 2019 SOT: Acids Pathway (Convert Lignin - "Base" Case)',
                                                     'Biochemical 2019 SOT: Acids Pathway (Convert Lignin - High)',
                                                     'Biochemical 2019 SOT: BDO Pathway (Burn Lignin Case)',
                                                     'Biochemical 2019 SOT: BDO Pathway (Convert Lignin - Base)',
                                                     'Biochemical 2019 SOT: BDO Pathway (Convert Lignin - High)',
                                                     #'Biomass to Gasoline and Diesel Using Integrated Hydropyrolysis and Hydroconversion',
                                                     #'Cellulosic Ethanol',
                                                     #'Cellulosic Ethanol with Jet Upgrading',
                                                     #'Fischer-Tropsch SPK',
                                                     #'Gasification to Methanol',
                                                     #'Gasoline from upgraded bio-oil from pyrolysis'
                                                     ])].reset_index(drop=True)


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
corr_itemized_LCA = pd.read_csv(input_path_corr + '/' + f_corr_itemized_LCI, header=0, index_col=0)

#%%

# Step: Create Cost Item table

# Subset cost items to use for itemized MFSP calculation
cost_items = df_econ.loc[df_econ['Item'].isin(['Purchased Inputs',
                                               'Waste Disposal',
                                               'Coproducts',
                                               'Fixed Costs',
                                               'Capital Depreciation',
                                               'Average Income Tax',
                                               'Average Return on Investment',
                                               'Cost by process steps']), : ].copy()
cost_items.drop_duplicates(inplace=True)

# Separate feedstock demand yearly flows
cost_feedstocks = df_econ.loc[df_econ['Item'] == 'Feedstock Cost', 
                             ['Case/Scenario', 'Stream Description', 'Flow Name', 
                              'Flow: Unit (numerator)', 'Flow: Unit (denominator)', 'Flow']].copy()
cost_feedstocks.rename(columns={'Stream Description' : 'Feedstock Stream Description',
                                'Flow Name' : 'Feedstock',
                                'Flow: Unit (numerator)' : 'Feedstock Flow: Unit (numerator)', 
                                'Flow: Unit (denominator)' : 'Feedstock Flow: Unit (denominator)',
                                'Flow' : 'Feedstock Flow'}, inplace=True)

# Merge with the cost items df
cost_items = pd.merge(cost_items, cost_feedstocks, how='left', on='Case/Scenario').reset_index(drop=True)

#%%

# Step: Create Biofuel Yield table and merge with Cost Item table

# Separate biofuel yield flows
biofuel_yield = df_econ.loc[df_econ['Item'] == 'Final Product',
                            ['Case/Scenario', 'Flow Name', 'Total Flow: Unit (numerator)', 
                             'Total Flow: Unit (denominator)', 'Total Flow']].reset_index(drop=True).copy()
biofuel_yield.rename(columns={'Flow Name' : 'Biofuel Flow Name',
                              'Total Flow: Unit (numerator)' : 'Biofuel Flow: Unit (numerator)', 
                              'Total Flow: Unit (denominator)' : 'Biofuel Flow: Unit (denominator)',
                              'Total Flow' : 'Biofuel Flow'}, inplace=True)

# For co-produced flows, summarize the flow data to one output
biofuel_yield2 = biofuel_yield.groupby(['Case/Scenario', 'Biofuel Flow: Unit (numerator)', 
                                        'Biofuel Flow: Unit (denominator)']).agg({'Biofuel Flow' : 'sum'}).reset_index()

# Merge with the cost items df
cost_items = pd.merge(cost_items, biofuel_yield2, how='left', on='Case/Scenario').reset_index(drop=True)

#%%

# Step: Correct for inflation to the year of study

# drop blanks
cost_items = cost_items.loc[~cost_items['Total Cost'].isin(['-']), : ]

cost_items['Adjusted Total Cost'] = cost_items.apply(lambda x: cpi.inflate(x['Total Cost'], x['Cost Year'], to=study_year), axis=1)
cost_items['Adjusted Cost Year'] = study_year

#%%

# Step: Calculate itemized and aggregrated Marginal Fuel Selling Price (MFSP)

cost_items['Itemized MFSP'] = cost_items['Adjusted Total Cost'].astype(float) / cost_items['Biofuel Flow'].astype(float)
cost_items['Itemized MFSP: Unit (numerator)'] = cost_items['Total Cost: Unit (numerator)']
cost_items['Itemized MFSP: Unit (denominator)'] = cost_items['Biofuel Flow: Unit (numerator)']

# For co-products we consider their cost as credit to the MFSP [co-product credit by displacement]
cost_items.loc[cost_items['Item'] == 'Coproducts', 'Itemized MFSP'] = \
    cost_items.loc[cost_items['Item'] == 'Coproducts', 'Itemized MFSP'] * -1

MFSP_agg = cost_items.copy()

if consider_coproduct_cost_credit == False:
    MFSP_agg = MFSP_agg.loc[~MFSP_agg['Item'].isin(['Coproducts']), :]

MFSP_agg = MFSP_agg[['Case/Scenario',
                       'Feedstock',
                       'Itemized MFSP: Unit (numerator)', 
                       'Itemized MFSP: Unit (denominator)',
                       'Adjusted Cost Year',
                       'Itemized MFSP']]
MFSP_agg = MFSP_agg[MFSP_agg['Itemized MFSP'].notna()]
    
MFSP_agg = MFSP_agg.groupby(['Case/Scenario',
                             'Feedstock',
                             'Itemized MFSP: Unit (numerator)', 
                             'Itemized MFSP: Unit (denominator)',
                             'Adjusted Cost Year']).agg({'Itemized MFSP' : 'sum'}).reset_index()
MFSP_agg.rename(columns={'Itemized MFSP' : 'MFSP replacing fuel',
                         'Itemized MFSP: Unit (numerator)' : 'MFSP replacing fuel: Unit (numerator)',
                         'Itemized MFSP: Unit (denominator)' : 'MFSP replacing fuel: Unit (denominator)'}, inplace=True)

# Getting back the Final Product column
MFSP_agg = pd.merge(biofuel_yield[['Case/Scenario', 'Biofuel Flow Name']].drop_duplicates(), 
                    MFSP_agg, how='left', on='Case/Scenario').reset_index(drop=True)

# Save interim data tables
if save_interim_files == True:
    cost_items.to_csv(output_path_prefix + '/' + f_out_itemized_mfsp)
    MFSP_agg.to_csv(output_path_prefix + '/' + f_out_agg_mfsp)

#%%

# Step: Merge Itemized LCAs to TEA-pathway LCIs

corr_itemized_LCA[['Unit (numerator)', 'Unit (denominator)']] = corr_itemized_LCA['Unit'].str.split('/', expand=True)

corr_itemized_LCA.rename(columns={'GREET row names_level1' : 'LCA_metric',
                                  'values_level1' : 'LCA_value',
                                  'Unit (numerator)' : 'LCA: Unit (numerator)',
                                  'Unit (denominator)' : 'LCA: Unit (denominator)'}, inplace=True)
corr_itemized_LCA = corr_itemized_LCA[['Item',
                                       'Stream Description',
                                       'Flow Name',
                                       'LCA_metric',
                                       'LCA_value',
                                       'LCA: Unit (numerator)',
                                       'LCA: Unit (denominator)',
                                       'Year']]

LCA_items = df_econ.loc[df_econ['Item'].isin(['Purchased Inputs',
                                               'Waste Disposal',
                                               'Coproducts']), : ].copy()
# temporary value for production year
LCA_items['Production Year'] = study_year

# Merge itemized LCAs to LCIs
LCA_items = pd.merge(LCA_items, corr_itemized_LCA, how='left', 
                     left_on=['Item', 'Stream Description', 'Flow Name', 'Production Year'],
                     right_on=['Item', 'Stream Description', 'Flow Name', 'Year']).reset_index(drop=True)

# remove trailing spaces for the Metric column
LCA_items['LCA_metric'] = LCA_items['LCA_metric'].str.strip()

# Considering CO2e metric only at the moment, 'CO2 (w/ C in VOC & CO)' is not considered now
LCA_items = LCA_items.loc[LCA_items['LCA_metric'].isin(['CO2']), :]

# harmonize units
# GREET tonnes represent Short Ton, convert to metric ton
LCA_items['LCA: Unit (denominator)'] = ['Short Tons' if val == 'ton' else val for val in LCA_items['LCA: Unit (denominator)'] ]
LCA_items['LCA_value'] = pd.to_numeric(LCA_items['LCA_value'])

# convert LCA unit of flow to model standard unit
LCA_items.loc[:, ['LCA: Unit (denominator)', 'LCA_value']] = \
    ob_units.unit_convert_df(LCA_items.loc[:, ['LCA: Unit (denominator)', 'LCA_value']],
     Unit = 'LCA: Unit (denominator)', Value = 'LCA_value',
     if_unit_numerator = False, if_given_category=False)

# converting material flow units to model standard units
LCA_items['Total Flow'] = ['0' if val == '-' else val for val in LCA_items['Total Flow'] ]
LCA_items['Total Flow'] = pd.to_numeric(LCA_items['Total Flow'])
LCA_items.loc[~(LCA_items['Total Flow: Unit (numerator)'].isin(['-'])), 
              ['Total Flow: Unit (numerator)', 'Total Flow']] = \
    ob_units.unit_convert_df(LCA_items.loc[~(LCA_items['Total Flow: Unit (numerator)'].isin(['-'])), ['Total Flow: Unit (numerator)', 'Total Flow']],
     Unit = 'Total Flow: Unit (numerator)', Value = 'Total Flow',
     if_unit_numerator = True, if_given_category=False)    

# Identify non-harmonized units if any
ignored_LCA_items = LCA_items.loc[LCA_items['Total Flow: Unit (numerator)'] != LCA_items['LCA: Unit (denominator)'], : ]
if ignored_LCA_items.shape[0] > 0:
    print("Warning: The following need attention as the units are not harmonized ..")
    print(ignored_LCA_items)

#%%

# Step: Itemized and aggregrated LCA metric per pathway

# Some LCA mappings are probably buggy, omitting it until QA
LCA_items = LCA_items.loc[~LCA_items['Flow Name'].isin(['Makeup Water',
                                                        'Makeup water',
                                                        'Cooling Tower Makeup',
                                                        'Cooling tower water makeup',
                                                        'Cooling tower chemicals',
                                                        'Cooling Water Makeup']), :].reset_index(drop=True)

# Calculate itemized LCA metric per year
LCA_items['Total LCA'] = LCA_items['LCA_value'] * LCA_items['Total Flow']
LCA_items['Total LCA: Unit (numerator)'] = LCA_items['LCA: Unit (numerator)']
LCA_items['Total LCA: Unit (denominator)'] = LCA_items['Total Flow: Unit (denominator)']

# If co-product, LCA is credited by displacement
LCA_items.loc[LCA_items['Item'].isin(['Coproducts']), 'Total LCA'] = \
    LCA_items.loc[LCA_items['Item'].isin(['Coproducts']), 'Total LCA'] * -1

# Merge biofuel yield data by Case/Scenario
LCA_items = pd.merge(LCA_items, biofuel_yield2, how='left', on='Case/Scenario').reset_index(drop=True)

# Calculate LCA metric per unit biofuel yield
LCA_items['Total LCA'] = LCA_items['Total LCA'] / LCA_items['Biofuel Flow']
LCA_items['Total LCA: Unit (denominator)'] = LCA_items['Biofuel Flow: Unit (numerator)']

#LCA_items['LCA_metric'] = ['CO2' if val in ['CO2', 'CO2 (w/ C in VOC & CO)'] else val for val in LCA_items['LCA_metric']]

LCA_items_agg = LCA_items.copy()

if consider_coproduct_env_credit == False:
    LCA_items_agg = LCA_items_agg.loc[~LCA_items_agg['Item'].isin(['Coproducts']), : ]

# Calculate net LCA metric per pathway
LCA_items_agg = LCA_items_agg.groupby(['Case/Scenario', 'LCA_metric', 
                                   'Total LCA: Unit (numerator)', 
                                   'Total LCA: Unit (denominator)',
                                   'Production Year'], as_index=False).\
    agg({'Total LCA' : 'sum'})

LCA_items_agg['Total LCA (g per MJ)'] = LCA_items_agg['Total LCA'] / 121.3 # Unit: g CO2e/MJ

# Save interim data tables
if save_interim_files == True:
    LCA_items.to_csv(output_path_prefix + '/' + f_out_itemized_LCA)
    LCA_items_agg.to_csv(output_path_prefix + '/' + f_out_agg_LCA)
    
#%%

# Step: Merge correspondence tables and GREET emission factors

# Merge aggregrated LCA metric to MFSP tables
MAC_df = pd.merge(MFSP_agg, LCA_items_agg, on=['Case/Scenario']).reset_index(drop=True)

# Aggregrate carbon intensities for various GHGs to CO2e
#ef = ef.groupby(['GREET Pathway', 'Unit (Numerator)', 'Unit (Denominator)', 'Case', 'Scope', 'Year'], as_index=False).\
#        agg({'Reference case' : 'sum', 'Elec0' : 'sum'}).reset_index(drop=True)
#ef['Formula'] = 'Carbon dioxide equivalent'
#ef['Formula'] = 'CO2e'
 
# map replaced fuels with replacing fuels
MAC_df = pd.merge(MAC_df, corr_replaced_replacing_fuel, how = 'left', 
               on=['Case/Scenario', 'Biofuel Flow Name', 'Feedstock']).reset_index(drop=True) 

# map replacing fuels with GREET pathways
#MAC_df = pd.merge(MAC_df, corr_fuel_replacing_GREET_pathway, how='left',
#               on=['Case/Scenario', 'Biofuel Flow Name', 'Feedstock']).reset_index(drop=True)
#MAC_df.rename(columns={'GREET Pathway' : 'GREET Pathway for replacing fuel'}, inplace=True)

# map replaced fuels with GREET pathways
MAC_df = pd.merge(MAC_df, corr_fuel_replaced_GREET_pathway, how='left', on=['Replaced Fuel']).reset_index(drop=True)
MAC_df.rename(columns={'GREET Pathway' : 'GREET Pathway for replaced fuel'}, inplace=True)

# map GREET carbon intensities for replaced fuels, considering Decarb Model reference case carbon intensities only
MAC_df = pd.merge(MAC_df, ef.loc[ef['Case'] == 'Reference case', : ], how='left', 
                  left_on=['GREET Pathway for replaced fuel', 'LCA_metric', 'Production Year'],
                  right_on=['GREET Pathway', 'Formula', 'Year']).reset_index(drop=True)
MAC_df.rename(columns={'Flow Name' : 'Flow Name_replaced fuel',
                       'Formula' :'Formula_replaced fuel',
                       'Unit (Numerator)' : 'CI replaced fuel: Unit (Numerator)',
                       'Unit (Denominator)' : 'CI replaced fuel: Unit (Denominator)',
                       'Case' : 'Case_replaced fuel',
                       'Scope' : 'Scope_replaced fuel',
                       'Reference case' : 'CI replaced fuel',
                       'Elec0' : 'CI Elec0_replaced fuel'}, inplace=True)
MAC_df.drop(['GREET Pathway'], axis=1, inplace=True)

# map GREET carbon intensities for replacing fuels
#MAC_df = pd.merge(MAC_df, ef, how='left', 
#                  left_on=['GREET Pathway for replacing fuel', 'Year', 'Formula_replaced fuel'],
#                  right_on=['GREET Pathway', 'Year', 'Formula']).reset_index(drop=True)
#MAC_df.rename(columns={'Flow Name' : 'Flow Name_replacing fuel',
#                       'Formula' :'Formula_replacing fuel',
#                       'Unit (Numerator)' : 'Unit (Numerator)_CI replacing fuel',
#                       'Unit (Denominator)' : 'Unit (Denominator)_CI replacing fuel',
#                       'Case' : 'Case_replacing fuel',
#                       'Scope' : 'Scope_replacing fuel',
#                       'Reference case' : 'CI_replacing fuel',
#                       'Elec0' : 'CI_Elec0_replacing fuel'}, inplace=True)
#MAC_df.drop(['GREET Pathway'], axis=1, inplace=True)

# Map MFSP of replaced fuels
MAC_df = pd.merge(MAC_df, 
                  EIA_price[['Year', 'Value', 'Energy carrier', 'Cost basis', 'Unit']],
                  how='left', 
                  left_on=['Year', 'Replaced Fuel'], 
                  right_on=['Year', 'Energy carrier']).reset_index(drop=True)
MAC_df.rename(columns={'Value' : 'Cost_replaced fuel',
                       'Cost basis' : 'Cost basis_replaced fuel'}, inplace=True)
MAC_df[['Year_Cost_replaced fuel', 'Unit Cost_replaced fuel (Numerator)']] = MAC_df['Unit'].str.split(' ', 1, expand = True)
MAC_df[['Unit Cost_replaced fuel (Numerator)', 
        'Unit Cost_replaced fuel (Denominator)']] = \
      MAC_df['Unit Cost_replaced fuel (Numerator)'].str.split('/', 1, expand = True)
MAC_df.rename(columns={'Unit Cost_replaced fuel (Numerator)' : 'Cost replaced fuel: Unit (Numerator)', 
                       'Unit Cost_replaced fuel (Denominator)' : 'Cost replaced fuel: Unit (Denominator)'}, inplace=True)

MAC_df.drop(['Energy carrier', 'Unit'], axis=1, inplace=True)

# Drop off data for which GREET pathways are not mapped until now
missing_GREET_mappings = MAC_df.loc[MAC_df['GREET Pathway for replaced fuel'].isna(), ['Case/Scenario', 'Biofuel Flow Name', 'Feedstock', 'Replaced Fuel']].drop_duplicates()
if missing_GREET_mappings.shape[0] > 0:
    print("Warning: The following pathways are currently dropped as their mappings to GREET CIs are not available as input ..")
    print(missing_GREET_mappings)
MAC_df = MAC_df.loc[~ MAC_df['GREET Pathway for replaced fuel'].isna(), :].copy()

# Assumption: non-liquid final products are skipped and not credited at the moment
# If classified as 'Coproducts', displacement based credit is considered for LCA and price credit is considered for MFSP 
MAC_df = MAC_df.loc[~ MAC_df['MFSP replacing fuel: Unit (denominator)'].isin(['lb']), : ].copy()

# dropping rows with no data on cost replaced fuel
MAC_df = MAC_df.loc[~MAC_df['Cost_replaced fuel'].isna(), :]

#%%

# Step: Correct for inflation to the year of study

MAC_df['Year_Cost_replaced fuel'] = pd.to_numeric(MAC_df['Year_Cost_replaced fuel'])
MAC_df['Adjusted Cost_replaced fuel'] = \
    MAC_df.apply(lambda x: cpi.inflate(x['Cost_replaced fuel'], x['Year_Cost_replaced fuel'], to=study_year), axis=1)
MAC_df['Adjusted Cost Year'] = study_year
    
#%%

# Step: Unit check and conversions

# Unit check for Replaced Fuel

# barrel to gallon
MAC_df[['Cost replaced fuel: Unit (Denominator)', 'Adjusted Cost_replaced fuel']] = \
    ob_units.unit_convert_df (
        MAC_df[['Cost replaced fuel: Unit (Denominator)', 'Adjusted Cost_replaced fuel']], 
        Unit='Cost replaced fuel: Unit (Denominator)', 
        Value='Adjusted Cost_replaced fuel', 
        if_unit_numerator = False,
        if_given_unit = True, 
        given_unit = 'gal').copy()
    
# Convert fuel cost USD per gallon to $ per GGE
# This conversion is done especially if certain calculations in future is required in GGE

# Map Replaced fuel to 'GREET_Fuel', 'GREET_Fuel type' for GGE conversion
MAC_df = pd.merge(MAC_df, corr_GGE_GREET_fuel_replaced, how='left', 
                  left_on=['Replaced Fuel'], 
                  right_on=['B2B fuel name']).reset_index(drop=True)

MAC_df = pd.merge(MAC_df, ob_units.hv_EIA[['GREET_Fuel', 'GREET_Fuel type', 'GGE']].drop_duplicates(), 
                  how='left', 
                  on=['GREET_Fuel', 'GREET_Fuel type']).reset_index(drop=True)
MAC_df['Adjusted Cost_replaced fuel'] = MAC_df['Adjusted Cost_replaced fuel'] / MAC_df['GGE']
MAC_df['Cost replaced fuel: Unit (Denominator)'] = 'GGE'
MAC_df['Cost replaced fuel: Unit (Numerator)'] = 'USD'
MAC_df.drop(columns=['GREET_Fuel', 'GREET_Fuel type', 'B2B fuel name', 'GGE'], inplace=True)

# Convert fuel cost from $ per GGE to $ per MMBtu
# extract CI of gasoline
tempdf = ob_units.hv_EIA.loc[(ob_units.hv_EIA['Energy carrier'] == 'Gasoline') &
                             (ob_units.hv_EIA['Energy carrier type'] == 'Petroleum Gasoline'), ['LHV', 'Unit']]
# convert unit
tempdf[['unit_numerator', 'unit_denominator']] = tempdf['Unit'].str.split('/', 1, expand=True)
tempdf.drop(columns=['Unit'], inplace=True)
tempdf[['unit_numerator', 'LHV']] = \
    ob_units.unit_convert_df (
        tempdf[['unit_numerator', 'LHV']],
        Unit='unit_numerator',
        Value='LHV',
        if_unit_numerator=True,
        if_given_unit=True,
        given_unit='mmbtu').copy()
tempdf['unit_denominator'] = 'GGE'
# merge with MAC df for unit conversion
MAC_df = pd.merge(MAC_df, tempdf, how='left', 
                  left_on='Cost replaced fuel: Unit (Denominator)', 
                  right_on='unit_denominator').reset_index(drop=True)
MAC_df['Adjusted Cost_replaced fuel'] = MAC_df['Adjusted Cost_replaced fuel']/MAC_df['LHV'] # unit: $/MMBTU
MAC_df['Cost replaced fuel: Unit (Denominator)'] = 'MMBtu'
MAC_df['Cost replaced fuel: Unit (Numerator)'] = 'USD'
MAC_df.drop(columns=['LHV', 'unit_numerator', 'unit_denominator'], inplace=True)


# Unit check for Replacing Fuel

# $/gal to $/GGE
# Map Replacing fuel to 'GREET_Fuel', 'GREET_Fuel type' for GGE conversion
MAC_df = pd.merge(MAC_df, corr_GGE_GREET_fuel_replacing, how='left', 
                  left_on=['Biofuel Flow Name'], 
                  right_on=['B2B fuel name']).reset_index(drop=True)
    
MAC_df = pd.merge(MAC_df, ob_units.hv_EIA[['GREET_Fuel', 'GREET_Fuel type', 'GGE']].drop_duplicates(), 
                  how='left', 
                  on=['GREET_Fuel', 'GREET_Fuel type']).reset_index(drop=True)
MAC_df.loc[MAC_df['MFSP replacing fuel: Unit (denominator)'] == 'gal', 'MFSP replacing fuel'] = \
    MAC_df.loc[MAC_df['MFSP replacing fuel: Unit (denominator)'] == 'gal', 'MFSP replacing fuel'] / \
    MAC_df.loc[MAC_df['MFSP replacing fuel: Unit (denominator)'] == 'gal', 'GGE']
MAC_df.loc[MAC_df['MFSP replacing fuel: Unit (denominator)'] == 'gal', 'MFSP replacing fuel: Unit (denominator)'] = 'GGE'
MAC_df.drop(columns=['GREET_Fuel', 'GREET_Fuel type', 'B2B fuel name', 'GGE'], inplace=True)

# Convert fuel cost from $ per GGE to $ per MMBtu
# extract CI of gasoline
tempdf = ob_units.hv_EIA.loc[(ob_units.hv_EIA['Energy carrier'] == 'Gasoline') &
                             (ob_units.hv_EIA['Energy carrier type'] == 'Petroleum Gasoline'), ['LHV', 'Unit']]
# convert unit
tempdf[['unit_numerator', 'unit_denominator']] = tempdf['Unit'].str.split('/', 1, expand=True)
tempdf.drop(columns=['Unit'], inplace=True)
tempdf[['unit_numerator', 'LHV']] = \
    ob_units.unit_convert_df (
        tempdf[['unit_numerator', 'LHV']],
        Unit='unit_numerator',
        Value='LHV',
        if_unit_numerator=True,
        if_given_unit=True,
        given_unit='mmbtu').copy()
tempdf['unit_denominator'] = 'GGE'
# merge with MAC df for unit conversion
MAC_df = pd.merge(MAC_df, tempdf, how='left', 
                  left_on='MFSP replacing fuel: Unit (denominator)', 
                  right_on='unit_denominator').reset_index(drop=True)
MAC_df['MFSP replacing fuel'] = MAC_df['MFSP replacing fuel']/MAC_df['LHV'] # unit: $/MMBTU
MAC_df.loc[MAC_df['MFSP replacing fuel: Unit (denominator)'] == 'GGE', 'MFSP replacing fuel: Unit (denominator)'] = 'MMBtu'
MAC_df.loc[MAC_df['Total LCA: Unit (denominator)'] == 'GGE', 'Total LCA'] = \
    MAC_df.loc[MAC_df['Total LCA: Unit (denominator)'] == 'GGE', 'Total LCA']/MAC_df['LHV'] # unit: g/MMBtu
MAC_df.loc[MAC_df['Total LCA: Unit (denominator)'] == 'GGE', 'Total LCA: Unit (denominator)'] = 'MMBtu'
MAC_df.drop(columns=['LHV', 'unit_numerator', 'unit_denominator'], inplace=True)

#%%
# Step: Calculate MAC by Cost Items

# MAC = (MFSP_biofuel - MFSP_ref) / (CI_ref - CI_biofuel)
# Unit: ($/MMBtu - $/MMBtu) / (g/MMBtu - g/MMBtu) = $/g
MAC_df['MAC_calculated'] = (MAC_df['MFSP replacing fuel'] - MAC_df['Adjusted Cost_replaced fuel']) / \
                           (MAC_df['CI replaced fuel'] - MAC_df['Total LCA'])
MAC_df['MAC_calculated: Unit (numerator)'] = MAC_df['MFSP replacing fuel: Unit (numerator)']
MAC_df['MAC_calculated: Unit (denominator)'] = MAC_df['Total LCA: Unit (numerator)'] 
MAC_df['MAC_calculated'] = MAC_df['MAC_calculated']*1E6 # unit: $/MT CO2 avoided
MAC_df['MAC_calculated: Unit (denominator)']  = 'MT'

MAC_df['CI of replaced fuel higher'] = MAC_df['CI replaced fuel'] > MAC_df['Total LCA']
MAC_df['Cost of replaced fuel higher'] = MAC_df['Adjusted Cost_replaced fuel'] > MAC_df['MFSP replacing fuel']

# Save interim data tables
if save_interim_files == True:
    MAC_df.to_csv(output_path_prefix + '/' + f_out_MAC)
    
    
print( '    Elapsed time: ' + str(datetime.now() - init_time)) 

"""
numeric_cols = ['Flow', 'Unit Cost', 'Feedstock Flow', 'Operating Time', 'Biofuel Flow']
for col_name in numeric_cols:
    cost_items.loc[cost_items[col_name] == '-', col_name] = '0'

cost_items[numeric_cols] = cost_items[numeric_cols].apply(pd.to_numeric)

# (lb/hr) * (usd/lb) / (US dry ton/yr) * (hr/yr) / (GGE/US dry ton)
cost_items['MAC Value'] = cost_items['Flow'] * cost_items['Unit Cost'] / \
                            cost_items['Feedstock Flow'] * cost_items['Operating Time'] /  cost_items['Biofuel Flow']

# Aggregrate MAC for each feedstock-biofuel conversion pathways
cost_items_agg = cost_items.groupby(['Case/Scenario', 'Feedstock', 'Biofuel Flow Name']).agg({'MAC Value' : 'sum'}).reset_index()
"""

    
#%%
