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

code_path_prefix = 'C:/Users/skar/repos/BestBiomassUse' # path to the Github local repository

input_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - Best use of biomass/data'
output_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - Best use of biomass/data/interim'

input_path_TEA = input_path_prefix + '/TEA'
input_path_LCA = input_path_prefix + '/LCA'
input_path_GREET = input_path_prefix + '/GREET'
input_path_EIA_price = input_path_prefix + '/EIA'
input_path_corr = input_path_prefix + '/correspondence_files'
input_path_units = input_path_prefix + '/Units'

f_TEA = 'MCCAM_04_12_2023_working.xlsx'
sheet_TEA = 'Db'
sheet_param_variability = 'var_p'

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
f_corr_itemized_LCI = 'corr_LCI_GREET_temporal_03_24_2023.csv'
f_corr_replaced_EIA_mfsp = 'corr_replaced_EIA_mfsp.csv'

#f_corr_params_variability = 'corr_params_variability.xlsx'
#sheet_corr_params_variability = 'input_table'

# Year of pathway production 
production_year = 2022

# cost adjust year to which inflation will be adjusted
cost_year = 2016

# Toggle cost credit for coproducts while calculating aggregrated MFSP
consider_coproduct_cost_credit = True

# Toggle to control emissions credit for coproducts while calculating aggregrated CIs
consider_coproduct_env_credit = True

# Toggle variability study
consider_variability_study = False

# Toggle write output to the dashboard workbook
write_to_dashboard = True


save_interim_files = True

dict_gco2e = { # Table 2, AR6/GWP100, GREET1 2022
    'CO2' : 1,
    'CO2 (w/ C in VOC & CO)' : 1,
    'N2O' : 273,
    'CH4' : 29.8,
    'Biogenic CH4' : 29.8}

# List of Stream_Flows that have biogenic C and their CO2 is not excluded
biogenic_lci = ['Biochar',
                'Flue gas',                
                ]
biogenic_emissions = ['Biogenic CO2',
                      'Biogenic CH4'  # biogenic CH4 are not considered to have environmental effect
                      ]

#%%
# import packages
import pandas as pd
import numpy as np
import os
from datetime import datetime
import cpi
import xlwings as xw
from xlwings.constants import DeleteShiftDirection
from xlwings.constants import AutoFillType

#cpi.update()

# Import user defined modules
os.chdir(code_path_prefix)

from unit_conversions import model_units

#%%
# User defined function definitions

# Function to expand on the input variability table 
def variability_table(var_params):
    var_tbl = pd.DataFrame(columns=var_params.columns.to_list() + ['param_value'] )
    for r in range(0, var_params.shape[0]):
        if var_params.loc[r, 'param_dist'] == 'linear':
            val = var_params.loc[r,'param_min']
            while (var_params.loc[r,'param_max'] - val) > 0.0001:  
               #var_params.loc[r,'param_value'] = val                              
               var_tbl = pd.concat([var_tbl, 
                                    pd.DataFrame({
                                     'col_param' : [var_params.loc[r,'col_param']],
                                     'col_val' : [var_params.loc[r,'col_val']],
                                     'param_name' : [var_params.loc[r,'param_name']],
                                     'param_min' : [var_params.loc[r,'param_min']],
                                     'param_max' : [var_params.loc[r,'param_max']],
                                     'param_dist' : [var_params.loc[r,'param_dist']],
                                     'dist_option' : [var_params.loc[r,'dist_option']],
                                     'param_value' : [val]})
                                    ])
               val = val + var_params.loc[r,'dist_option']
    return var_tbl

def mult_numeric(a,b,c):
    if ((type(a) is int) | (type(a) is float)) &\
       ((type(b) is int) | (type(b) is float)) &\
       ((type(c) is int) | (type(c) is float)):
           return a*b*c
    else:
         return 

# Function to add header rows to LCA metric rows, select subset of LCA metrices, and calculate CO2e
def fmt_GREET_LCI(df):
    df = corr_itemized_LCA.copy()   
    #df = df.loc[(df['Item'] == 'Coproducts') & (df['Stream_Flow'] == 'Renewable Gasoline'), : ].reset_index(drop=True)
    
    df = df.drop_duplicates().reset_index(drop=True)
    df.fillna('', inplace=True)    
    
    # add header to LCI metrices rows    
    df[['Unit (numerator)', 'Unit (denominator)']] = df['Unit'].str.split('/', expand=True)

    df.rename(columns={'GREET row names_level1' : 'LCA_metric',
                       'values_level1' : 'LCA_value',
                       'Unit (numerator)' : 'LCA: Unit (numerator)',
                       'Unit (denominator)' : 'LCA: Unit (denominator)'}, inplace=True)
    df = df[['Item',
             'Stream_Flow',
             'Stream_LCA',
             'GREET1 sheet',
             'Coproduct allocation method',
             'GREET classification of coproduct',
             'LCA_metric',
             'LCA_value',
             'LCA: Unit (numerator)',
             'LCA: Unit (denominator)',
             'Year']]
    for row in range(df.shape[0]):
        if df.loc[row,'LCA_metric'] == df.loc[row, 'LCA_metric'].strip():
            header_val = df.loc[row,'LCA_metric']
        else:
            # concatenatig with header 
            df.loc[row,'LCA_metric'] = header_val + '__' + df.loc[row, 'LCA_metric'].strip()
    
    # remove the header rows without values
    df = df.loc[~(df['LCA: Unit (numerator)'].isna()), : ]
    
    # select a subset of LCA metrices
    df = df.loc[(df['LCA_metric'].str.contains('CO2', regex=False)) | 
                (df['LCA_metric'].str.contains('N2O', regex=False)) |
                (df['LCA_metric'].str.contains('CH4', regex=False)) |
                (df['LCA_metric'].str.contains('CO2 (w/ C in VOC & CO)', regex=False)) |
                (df['LCA_metric'].str.contains('GHGs (grams/ton)', regex=False)) |
                (df['LCA_metric'].str.contains('GHGs', regex=False)), : ]    
    
    # If CO2, CH4, N2O are available, ignore GHG or CO2 w/ VOC mertics
    df['count_m'] = (((df['LCA_metric'].str.contains('CO2', regex=False)) &\
                     ~(df['LCA_metric'].str.contains('CO2 (w/ C', regex=False))).replace({True:1, False:0})) +\
                    (((df['LCA_metric'].str.contains('CO2', regex=False)) &\
                    (df['LCA_metric'].str.contains('CO2 (w/ C', regex=False))).replace({True:1, False:0})) +\
                    (df['LCA_metric'].str.contains('N2O', regex=False).replace({True:1, False:0})) +\
                    (((df['LCA_metric'].str.contains('CH4', regex=False)) &\
                     ~(df['LCA_metric'].str.contains('Biogenic CH4', regex=False))).replace({True:1, False:0}))
    df_sub = df.groupby(['Item', 'Stream_Flow', 'Stream_LCA', 
                         'Year','GREET1 sheet','Coproduct allocation method',
                         'GREET classification of coproduct'], dropna=False).agg({'count_m' : 'sum'}).reset_index()
    df = pd.merge(df, df_sub, how='left', on=['Item', 'Stream_Flow', 'Stream_LCA',
                                              'Year', 'GREET1 sheet', 'Coproduct allocation method',
                                              'GREET classification of coproduct']).reset_index(drop=True)
    
    # if CO2 and CO2 (w/C ..) both are present the count becomes 4
    df_sub = df.loc[(df['count_m_y'] == 4), :]    
    df_sub = df_sub.loc[((df_sub['LCA_metric'].str.contains('CO2', regex=False)) &\
                        ~(df['LCA_metric'].str.contains('CO2 (w/ C', regex=False))) |                         
                        (df_sub['LCA_metric'].str.contains('N2O', regex=False)) |
                        (df_sub['LCA_metric'].str.contains('CH4', regex=False)), :]     
    df = df.loc[df['count_m_y']!=4, : ].copy()
    df = pd.concat([df,df_sub]).reset_index(drop=True)
    
    # if CO2 or CO2 (w/C ..) one is present the count becomes 3
    df_sub = df.loc[(df['count_m_y'] == 3), :]
    df_sub = df_sub.loc[((df_sub['LCA_metric'].str.contains('CO2', regex=False)) &\
                        ~(df['LCA_metric'].str.contains('CO2 (w/ C', regex=False))) | 
                        ((df_sub['LCA_metric'].str.contains('CO2', regex=False)) &\
                        (df['LCA_metric'].str.contains('CO2 (w/ C', regex=False))) |
                        (df_sub['LCA_metric'].str.contains('N2O', regex=False)) |
                        (df_sub['LCA_metric'].str.contains('CH4', regex=False)), :]     
    df = df.loc[df['count_m_y']!=3, : ].copy()
    df = pd.concat([df,df_sub]).reset_index(drop=True)
    
    df_sub = df['LCA_metric'].str.split('__', expand=True)
    if df_sub.shape[1] == 2:        
        df[['dummy_metric', 'LCA_metric']] = df['LCA_metric'].str.split('__', expand=True)
        df.loc[df['LCA_metric'].isna(), 'LCA_metric'] = df.loc[df['LCA_metric'].isna(), 'dummy_metric'] 
        df.drop(columns=['count_m_x', 'count_m_y', 'dummy_metric'], inplace=True)
    else:
        df.drop(columns=['count_m_x', 'count_m_y'], inplace=True)
    
    # Avoid biogenic stream flow
    df = df.loc[~((df['Stream_Flow'].isin(biogenic_lci)) &
                (df['LCA_metric'].isin(['CO2']))), : ]
    # Avoid biogenic emissions
    df = df.loc[~(df['LCA_metric'].isin(biogenic_emissions)), : ]
    
    # calculate CO2e
    df['mult'] = df['LCA_metric'].map(dict_gco2e)
    df['LCA_value'] = pd.to_numeric(df['LCA_value'])
    df['LCA_value'] = df['LCA_value'] * df['mult']    
    df = df.groupby(['Item', 'Stream_Flow', 'Stream_LCA', 
                     'Year','GREET1 sheet','Coproduct allocation method',
                     'GREET classification of coproduct',
                     'LCA: Unit (numerator)', 'LCA: Unit (denominator)']).agg({'LCA_value' : 'sum'}).reset_index()
    df['LCA_metric'] = 'CO2e'
    
    # harmonize units
    # GREET tonnes represent Short Ton
    df['LCA: Unit (denominator)'] = ['Short Tons' if val == 'ton' else val for val in df['LCA: Unit (denominator)'] ]
        
    # convert LCA unit of flow to model standard unit
    df.loc[:, ['LCA: Unit (denominator)', 'LCA_value']] = \
        ob_units.unit_convert_df(df.loc[:, ['LCA: Unit (denominator)', 'LCA_value']],
         Unit = 'LCA: Unit (denominator)', Value = 'LCA_value',
         if_unit_numerator = False, if_given_category=False)   
    
    return df


def fmt_GREET_LCI_2(df):
    
    harmonize_headers = {
        
        'Energy demand' : [
            'Energy demand',
            'Energy: mmBtu/ton',
            'Energy Use: mmBtu/ton of product',
            'Energy Use: mmBtu/ton',
            'Energy: Btu/g of material throughput, except as noted',
            'Energy Use: mmBtu per ton',
            'Energy Consumption: Btu/mmBtu of fuel transported',
            'Energy use: Btu/gal treated',
            'Total energy, Btu'
            ],
        
        'Water consumption' : [
            'Water consumption',
            'Water consumption, gallons/ton',
            'Water Consumption',
            'Water consumption: gallons',
            'Water consumption: gallons per ton',
            'Water consumption',
            'Water consumption: gallon/ton',
            'Water consumption (gal/g)',
            'Water consumption: gallons/mmBtu of fuel throughput',
            'Water consumption: gallons/ton',
            'Water consumption: gallons/mmBtu of fuel transported',
            'Water consumption, gallons/gal treated',
            'Water consumption, gallons/mmBtu of fuel throughput'
                               ],
        'Total emissions' : [
            'Total emissions',
            'Total Emissions: grams/ton',
            'Total emissions: grams/mmBtu of fuel throughput, except as noted',
            'Total emissions: grams/g of material throughput, except as noted',
            'Total Emissions: grams per ton',
            'Total Emissions: grams/mmBtu fuel transported',
            'Total emissions: grams/gal treated',
            'Total Emissions: grams/mmBtu of fuel throughput, except as noted',
            'Total emissions: grams'
            ],
        
        'Urban emissions' : [
            'Urban emissions',
            'Urban emissions: grams/ton',
            'Urban Emissions: grams/ton',
            'Urban emissions: grams/g of material throughput, except as noted',
            '5.2) Urban Emissions: Grams per mmBtu of Fuel Throughput at Each Stage',
            '4.2) Urban Emissions: Grams per mmBtu of Fuel Throughput at Each Stage',
            'Urban Emissions: grams/mmBtu of fuel transported',
            'Urban emissions: grams/gal treated',
            'Urban emissions: grams',
            'Urban emissions: grams/mmBtu of fuel throughput, except as noted',
            'Urban Emissions: grams per ton'
            ],                
        }
    
    harmonize_metric = {
        
        'Total energy' : [
            'Total Energy',
            'Total energy',
            ],
        
        'Fossil fuels' : [
            'Fossil fuels',
            'Fossil Fuels',
            'Fossil energy',
            'Fossil fuels, Btu',
            'Fossi lfuels',
            ],
        
        'Coal' : [
            'Coal',
            'Coal, Btu',
            ],
        
        'Natural gas' : [
            'Natural gas',
            'Natural Gas',
            'Natural gas, Btu',
            ],
       
        'Petroleum' : [
            'Petroleum',
            'Petroleum, Btu',
            ],
        
        'VOC' : [
            'VOC',
            'Urban VOC',
            'VOC from bulk terminal',
            'VOC from ref. Station',
            'VOC from refueling station',
            'VOC: Total',
            'VOC: Urban',
            ],
       
        'CO' : [
            'CO',
            'Urban CO',
            'CO: Total',
            'CO: Urban',
            ],
        
        'NOx' : [
            'NOx',
            'Urban NOx',
            'NOx: Total',
            'NOx: Urban',
            ],
        
        'PM10' : [
            'PM10',
            'Urban PM10',
            'PM10: Total',
            'PM10: Urban',
            ],
        
        'PM2.5' : [
            'PM2.5',
            'Urban PM2.5',
            'PM2.5: Total',
            'PM2.5: Urban',
            ],
        
        'SOx' : [
            'SOx',
            'Urban SOx',
            'SOx: Total',
            'SOx: Urban',
                ],
        
        'BC' : [
            'BC',
            'Urban BC',
            'BC Total',
            'BC: Urban',
            'BC, Total'
            ],
        
        'OC' : [
            'OC',
            'Urban OC',
            'OC Total',
            'OC: Urban',
            'OC, Total',
            ],
        
        'CH4' : [
            'CH4',
            'CH4: combustion',
            ],
       
        'N2O' : [
            'N2O',
            ],
        
        'CO2' : [
            'CO2',
            'Misc. CO2',
            ],
        
        'CO2 (w/ C in VOC & CO)' : [
            'CO2 (w/ C in VOC & CO)',
            ],
        
        'GHGs' : [
            'GHGs (grams/ton)',
            'GHGs',
            ],
        
        'Other GHG Emissions' : [
            'Other GHG Emissions',
            ],
        
        'Biogenic CH4' : [
            'Biogenic CH4',
            ],
        
        'Biogenic CO2' : [
            'Biogenic CO2',
            ],

        }
    
    
    harmonize_headers_long = {}
    for k, v in harmonize_headers.items():
        for v1 in v:
            harmonize_headers_long[v1] = k
    harmonize_metric_long = {}
    for k, v in harmonize_metric.items():
        for v1 in v:
            harmonize_metric_long[v1] = k
    
    
    df = corr_itemized_LCA.copy()   
    #df = df.loc[(df['Item'] == 'Coproducts') & (df['Stream_Flow'] == 'Acetone'), : ].reset_index(drop=True)
    
    
    #df = df.drop_duplicates().reset_index(drop=True)
    df.fillna('', inplace=True)    
         
    df.rename(columns={'GREET row names_level1' : 'LCA_metric_GREET',
                       'values_level1' : 'LCA_value'}, inplace=True)
    
    # drop rows with loss factor
    df = df.loc[~(df['LCA_metric_GREET'] == 'Loss factor'), :].reset_index(drop=True)
    
    df[['LCA: Unit (numerator)', 'LCA: Unit (denominator)']] = df['Unit'].str.split('/', expand=True)
    
    df = df[['Item',
             'Stream_Flow',
             'Stream_LCA',
             'GREET1 sheet',
             'Coproduct allocation method',
             'GREET classification of coproduct',
             'LCA_metric_GREET',
             'LCA_value',
             'LCA: Unit (numerator)',
             'LCA: Unit (denominator)',
             'Year']]              
    
    # strip white spaces before and after metric names
    for row in range(df.shape[0]):
        df.loc[row,'LCA_metric_GREET'] = df.loc[row, 'LCA_metric_GREET'].strip()
    
    # identify unique headers
    #tmp_lci_metric = df['LCA_metric'].unique()
    
    # replace with harmonized headers and metrices names
    df.loc[df['LCA_metric_GREET'].isin(harmonize_headers_long.keys()), 'LCA_metric'] =\
        df.loc[df['LCA_metric_GREET'].isin(harmonize_headers_long.keys()), 'LCA_metric_GREET'].map(harmonize_headers_long, na_action='ignore')
    
    df.loc[df['LCA_metric_GREET'].isin(harmonize_metric_long.keys()), 'LCA_metric'] =\
        df.loc[df['LCA_metric_GREET'].isin(harmonize_metric_long.keys()), 'LCA_metric_GREET'].map(harmonize_metric_long, na_action='ignore')
    df['LCA_metric'].fillna('-', inplace=True) 
    
    # Join header to LCI metric    
    for row in range(df.shape[0]):
        if df.loc[row,'LCA_metric'] in harmonize_headers.keys():
            header_val = df.loc[row,'LCA_metric']
        else:
            # concatenating with header 
            df.loc[row,'LCA_metric'] = header_val + '__' + df.loc[row, 'LCA_metric'].strip()
    
    # remove the header rows except water consumption
    df = df.loc[~(df['LCA_metric'].isin(harmonize_headers.keys()-['Water consumption'])), : ]
    
    # select a subset of LCA metrices
    df = df.loc[(df['LCA_metric'].str.contains('CO2', regex=False)) | 
                (df['LCA_metric'].str.contains('N2O', regex=False)) |
                (df['LCA_metric'].str.contains('CH4', regex=False)) |
                (df['LCA_metric'].str.contains('CO2 (w/ C in VOC & CO)', regex=False)) |
                (df['LCA_metric'].str.contains('GHGs (grams/ton)', regex=False)) |
                (df['LCA_metric'].str.contains('GHGs', regex=False)), : ]    
    
    # If CO2, CH4, N2O are available, ignore GHG or CO2 w/ VOC mertics
    df['count_m'] = (((df['LCA_metric'].str.contains('CO2', regex=False)) &\
                     ~(df['LCA_metric'].str.contains('CO2 (w/ C', regex=False))).replace({True:1, False:0})) +\
                    (((df['LCA_metric'].str.contains('CO2', regex=False)) &\
                    (df['LCA_metric'].str.contains('CO2 (w/ C', regex=False))).replace({True:1, False:0})) +\
                    (df['LCA_metric'].str.contains('N2O', regex=False).replace({True:1, False:0})) +\
                    (((df['LCA_metric'].str.contains('CH4', regex=False)) &\
                     ~(df['LCA_metric'].str.contains('Biogenic CH4', regex=False))).replace({True:1, False:0}))
    df_sub = df.groupby(['Item', 'Stream_Flow', 'Stream_LCA', 
                         'Year','GREET1 sheet','Coproduct allocation method',
                         'GREET classification of coproduct'], dropna=False).agg({'count_m' : 'sum'}).reset_index()
    df = pd.merge(df, df_sub, how='left', on=['Item', 'Stream_Flow', 'Stream_LCA',
                                              'Year', 'GREET1 sheet', 'Coproduct allocation method',
                                              'GREET classification of coproduct']).reset_index(drop=True)
    
    # if CO2 and CO2 (w/C ..) both are present the count becomes 4
    df_sub = df.loc[(df['count_m_y'] == 4), :]    
    df_sub = df_sub.loc[((df_sub['LCA_metric'].str.contains('CO2', regex=False)) &\
                        ~(df['LCA_metric'].str.contains('CO2 (w/ C', regex=False))) |                         
                        (df_sub['LCA_metric'].str.contains('N2O', regex=False)) |
                        (df_sub['LCA_metric'].str.contains('CH4', regex=False)), :]     
    df = df.loc[df['count_m_y']!=4, : ].copy()
    df = pd.concat([df,df_sub]).reset_index(drop=True)
    
    # if CO2 or CO2 (w/C ..) one is present the count becomes 3
    df_sub = df.loc[(df['count_m_y'] == 3), :]
    df_sub = df_sub.loc[((df_sub['LCA_metric'].str.contains('CO2', regex=False)) &\
                        ~(df['LCA_metric'].str.contains('CO2 (w/ C', regex=False))) | 
                        ((df_sub['LCA_metric'].str.contains('CO2', regex=False)) &\
                        (df['LCA_metric'].str.contains('CO2 (w/ C', regex=False))) |
                        (df_sub['LCA_metric'].str.contains('N2O', regex=False)) |
                        (df_sub['LCA_metric'].str.contains('CH4', regex=False)), :]     
    df = df.loc[df['count_m_y']!=3, : ].copy()
    df = pd.concat([df,df_sub]).reset_index(drop=True)
    
    df_sub = df['LCA_metric'].str.split('__', expand=True)
    if df_sub.shape[1] == 2:        
        df[['dummy_metric', 'LCA_metric']] = df['LCA_metric'].str.split('__', expand=True)
        df.loc[df['LCA_metric'].isna(), 'LCA_metric'] = df.loc[df['LCA_metric'].isna(), 'dummy_metric'] 
        df.drop(columns=['count_m_x', 'count_m_y', 'dummy_metric'], inplace=True)
    else:
        df.drop(columns=['count_m_x', 'count_m_y'], inplace=True)
    
    # Avoid biogenic stream flow
    df = df.loc[~((df['Stream_Flow'].isin(biogenic_lci)) &
                (df['LCA_metric'].isin(['CO2']))), : ]
    # Avoid biogenic emissions
    df = df.loc[~(df['LCA_metric'].isin(biogenic_emissions)), : ]
    
    # calculate CO2e
    df['mult'] = df['LCA_metric'].map(dict_gco2e)
    df['LCA_value'] = pd.to_numeric(df['LCA_value'])
    df['LCA_value'] = df['LCA_value'] * df['mult']    
    df = df.groupby(['Item', 'Stream_Flow', 'Stream_LCA', 
                     'Year','GREET1 sheet','Coproduct allocation method',
                     'GREET classification of coproduct',
                     'LCA: Unit (numerator)', 'LCA: Unit (denominator)']).agg({'LCA_value' : 'sum'}).reset_index()
    df['LCA_metric'] = 'CO2e'
    
    # harmonize units
    # GREET tonnes represent Short Ton
    df['LCA: Unit (denominator)'] = ['Short Tons' if val == 'ton' else val for val in df['LCA: Unit (denominator)'] ]
        
    # convert LCA unit of flow to model standard unit
    df.loc[:, ['LCA: Unit (denominator)', 'LCA_value']] = \
        ob_units.unit_convert_df(df.loc[:, ['LCA: Unit (denominator)', 'LCA_value']],
         Unit = 'LCA: Unit (denominator)', Value = 'LCA_value',
         if_unit_numerator = False, if_given_category=False)   
    
    return df




def ef_calc_co2e(df):
    # calculate CO2e
    df['mult'] = df['Formula'].map(dict_gco2e)
    df['Reference case'] = pd.to_numeric(df['Reference case'])
    df['Elec0'] = pd.to_numeric(df['Elec0'] )
    df['Reference case'] = df['Reference case'] * df['mult']   
    df['Elec0'] = df['Elec0'] * df['mult'] 
    df = df.groupby(['GREET Pathway', 'Unit (Numerator)',
           'Unit (Denominator)', 'Case', 'Scope', 'Year'
           ]).agg({'Reference case' : 'sum',
                   'Elec0' : 'sum'}).reset_index()
    df['Formula'] = 'CO2e'
    df['Stream_LCA'] = 'Carbon dioxide equivalent'
    return df
    

#%%
# Step: Load data file and select columns for computation

init_time = datetime.now()

df_econ = pd.read_excel(input_path_TEA + '/' + f_TEA, sheet_name = sheet_TEA, header = 3, index_col=None)

df_econ = df_econ[['Case/Scenario', 'Parameter',
       'Item', 'Stream_Flow', 'Stream_LCA', 'Flow: Unit (numerator)',
       'Flow: Unit (denominator)', 'Flow', 'Cost Item',
       'Cost: Unit (numerator)', 'Cost: Unit (denominator)', 'Unit Cost',
       'Operating Time: Unit', 'Operating Time', 'Operating Time (%)',
       'Total Cost: Unit (numerator)', 'Total Cost: Unit (denominator)',
       'Total Cost', 'Total Flow: Unit (numerator)',
       'Total Flow: Unit (denominator)', 'Total Flow', 'Cost Year']]

# Select pathways to consider
pathways_to_consider=[
        
        ###
        '2020, 2019 SOT High Octane Gasoline from Lignocellulosic Biomass via Syngas and Methanol/Dimethyl Ether Intermediates',
        ###
        
        # Tan et al., 2016 pathways
        ###
        # 'Pathway 1A: Syngas to molybdenum disulfide (MoS2)-catalyzed alcohols followed by fuel production via alcohol condensation (Guerbet reaction), dehydration, oligomerization, and hydrogenation',
        # 'Pathway 1B: Syngas fermentation to ethanol followed by fuel production via alcohol condensation (Guerbet reaction), dehydration, oligomerization, and hydrogenation',
        # 'Pathway 2A: Syngas to rhodium (Rh)-catalyzed mixed oxygenates followed by fuel production via carbon coupling/deoxygenation (to isobutene), oligomerization, and hydrogenation',
        # 'Pathway 2B: Syngas fermentation to ethanol followed by fuel production via carbon coupling/deoxygenation (to isobutene), oligomerization, and hydrogenation',
        # 'Pathway FT: Syngas to liquid fuels via Fischer-Tropsch technology as a commercial benchmark for comparisons',
        ###
        
        # Decarb 2b pathways
        # 'Thermochemical Research Pathway to High-Octane Gasoline Blendstock Through Methanol/Dimethyl Ether Intermediates',
        # 'Cellulosic Ethanol',
        ###
         'Decarb 2b: Cellulosic Ethanol to renewable gasoline and jet fuels',
         'Decarb 2b: Cellulosic Ethanol to renewable gasoline and jet fuels with CCS of fermentation offgas CO2',
         'Decarb 2b: Cellulosic Ethanol to renewable gasoline and jet fuels with CCS of fermentation offgas and boiler vent streams CO2',
         
         'Decarb 2b: Cellulosic Ethanol to renewable gasoline and jet fuels_jet',
         'Decarb 2b: Cellulosic Ethanol to renewable gasoline and jet fuels with CCS of fermentation offgas CO2_jet',
         'Decarb 2b: Cellulosic Ethanol to renewable gasoline and jet fuels with CCS of fermentation offgas and boiler vent streams CO2_jet',
         
         # 'Decarb 2b: Fischer-Tropsch SPK',
         # 'Decarb 2b: Fischer-Tropsch SPK with CCS of FT flue gas CO2',
         # 'Decarb 2b: Fischer-Tropsch SPK with CCS of all flue gases CO2',
         # 'Decarb 2b: Ex-Situ CFP',
         # 'Decarb 2b: Ex-Situ CFP with CCS of all flue gases CO2',
        ###
        # 'Gasification to Methanol',
        # 'Gasoline from upgraded bio-oil from pyrolysis'
        
        # 2021 SOT pathways
        ###
        # '2021 SOT: Biochemical design case, Acids pathway with burn lignin',
        # '2021 SOT: Biochemical design case, Acids pathway with convert lignin to BKA',
        # '2021 SOT: Biochemical design case, BDO pathway with burn lignin',
        # '2021 SOT: Biochemical design case, BDO pathway with convert lignin to BKA',
        # '2021 SOT: High octane gasoline from lignocellulosic biomass via syngas and methanol/dimethyl ether intermediates',
        
        # '2020 SOT: Ex-Situ CFP of lignocellulosic biomass to hydrocarbon fuels',
        ###
        
        # Marine pathways
        ###
        # '2022, Marine biocrude via HTL from sludge with NH3 removal for 1000 MTPD sludge',
        # '2022, Marine biocrude via HTL from Manure with NH3 removal for 1000 MTPD Manure',
        # '2022, Partially upgraded marine fuel via HTL from sludge with NH3 removal for 1000 MTPD sludge',
        # '2022, Partially upgraded marine fuel via HTL from Manure with NH3 removal for 1000 MTPD Manure',
        # '2022, Fully upgraded marine fuel via HTL from sludge with NH3 removal for 1000 MTPD sludge',
        # '2022, Fully upgraded marine fuel via HTL from Manure with NH3 removal for 1000 MTPD Manure',
        # '2022, Marine fuel through Catalytic Fast Pyrolysis with ZSM5 of blended woody biomass',
        # '2022, Marine fuel through Catalytic Fast Pyrolysis with Pt/TiO2 of blended woody biomass',
        ###
        
        
        
        # '2013 Biochemical Design Case: Corn Stover-Derived Sugars to Diesel',
        # '2015 Biochemical Catalysis Design Report',
        # '2018 Biochemical Design Case: BDO Pathway',
        # '2018 Biochemical Design Case: Organic Acids Pathway',
        # '2018, 2018 SOT High Octane Gasoline from Lignocellulosic Biomass via Syngas and Methanol/Dimethyl Ether Intermediates',
        # '2018, 2022 projection High Octane Gasoline from Lignocellulosic Biomass via Syngas and Methanol/Dimethyl Ether Intermediates',
                
        # '2020, 2022 projection High Octane Gasoline from Lignocellulosic Biomass via Syngas and Methanol/Dimethyl Ether Intermediates',
        # 'Biochemical 2019 SOT: Acids Pathway (Burn Lignin Case)',
        # 'Biochemical 2019 SOT: Acids Pathway (Convert Lignin - "Base" Case)',
        # 'Biochemical 2019 SOT: Acids Pathway (Convert Lignin - High)',
        # 'Biochemical 2019 SOT: BDO Pathway (Burn Lignin Case)',
        # 'Biochemical 2019 SOT: BDO Pathway (Convert Lignin - Base)',
        # 'Biochemical 2019 SOT: BDO Pathway (Convert Lignin - High)',
        # 'Biomass to Gasoline and Diesel Using Integrated Hydropyrolysis and Hydroconversion',
        # 'Corn stover ETJ', 
        # 'Dry Mill (Corn) ETJ',
        # 'Ex Situ CFP 2022 Target Case', 
        # 'Ex-Situ CFP 2019 SOT',
        # 'Ex-Situ Fixed Bed 2018 SOT (0.5 wt% Pt/TiO2 Catalyst)',
        # 'Ex-Situ Fixed Bed 2022 Projection',
        # 'In-Situ CFP 2022 Target Case',      
        
        ]
df_econ = df_econ.loc[df_econ['Case/Scenario'].isin(pathways_to_consider)].reset_index(drop=True)

# When studying variability of unit cost on MFSP and MAC,
# following pathways are avoided because detailed LCI are not available yet
cases_to_avoid = ['Cellulosic Ethanol',
                  'Cellulosic Ethanol with Jet Upgrading',
                  'Fischer-Tropsch SPK',
                  'Gasification to Methanol',
                  'Gasoline from upgraded bio-oil from pyrolysis']

# Exclude cases to avoid if performing variability analysis
if consider_variability_study:
    df_econ = df_econ.loc[~df_econ['Case/Scenario'].isin(cases_to_avoid)].reset_index(drop=True)

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
corr_replaced_EIA_mfsp = pd.read_csv(input_path_corr + '/' + f_corr_replaced_EIA_mfsp, header=3, index_col=None) 
if consider_variability_study:
    corr_params_variability =  pd.read_excel(input_path_TEA + '/' + f_TEA, sheet_name=sheet_param_variability, header=3, index_col=None)

#%%

# Step: Create Cost Item table

df_econ.loc[df_econ['Stream_Flow'].isna(), 'Stream_Flow'] = ''

# Subset cost items to use for itemized MFSP calculation
cost_items = df_econ.loc[df_econ['Item'].isin(['Feedstock',
                                               'Purchased Inputs',
                                               'Waste Disposal',
                                               'Coproducts',
                                               'Fixed Costs',
                                               
                                               'Capital Depreciation',
                                               'Average Income Tax',
                                               'Average Return on Investment',
                                               
                                               'Cost by process steps']), : ].copy()
cost_items.drop_duplicates(inplace=True)

# Separate feedstock demand yearly flows
cost_feedstocks = df_econ.loc[df_econ['Item'] == 'Feedstock', 
                             ['Case/Scenario', 'Stream_Flow', 'Stream_LCA', 
                              'Flow: Unit (numerator)', 'Flow: Unit (denominator)', 'Flow']].copy()
cost_feedstocks.rename(columns={'Stream_Flow' : 'Feedstock Stream_Flow',
                                'Stream_LCA' : 'Feedstock',
                                'Flow: Unit (numerator)' : 'Feedstock Flow: Unit (numerator)', 
                                'Flow: Unit (denominator)' : 'Feedstock Flow: Unit (denominator)',
                                'Flow' : 'Feedstock Flow'}, inplace=True)

# Merge with the cost items df
cost_items = pd.merge(cost_items, cost_feedstocks, how='left', on='Case/Scenario').reset_index(drop=True)

#%%

# Step: Create Biofuel Yield table and merge with Cost Item table

# Separate biofuel yield flows
biofuel_yield = df_econ.loc[df_econ['Item'] == 'Final Product',
                            ['Case/Scenario', 'Stream_LCA', 'Total Flow: Unit (numerator)', 
                             'Total Flow: Unit (denominator)', 'Total Flow']].reset_index(drop=True).copy()
biofuel_yield.rename(columns={'Stream_LCA' : 'Biofuel Stream_LCA',
                              'Total Flow: Unit (numerator)' : 'Biofuel Flow: Unit (numerator)', 
                              'Total Flow: Unit (denominator)' : 'Biofuel Flow: Unit (denominator)',
                              'Total Flow' : 'Biofuel Flow'}, inplace=True)

# For co-produced flows, summarize the flow data to one output
biofuel_yield2 = biofuel_yield.groupby(['Case/Scenario', 'Biofuel Flow: Unit (numerator)', 
                                        'Biofuel Flow: Unit (denominator)']).agg({'Biofuel Flow' : 'sum'}).reset_index()

# Merge with the cost items df
cost_items = pd.merge(cost_items, biofuel_yield2, how='left', on='Case/Scenario').reset_index(drop=True)

#%%
# Step: calculate cost per variability of parameters

# drop blanks
cost_items = cost_items.loc[~(cost_items['Total Cost'].isin(['-']) | cost_items['Total Cost'].isna()), : ]

if consider_variability_study:
    

    # unit check
    check_units = (cost_items['Flow: Unit (numerator)'] != cost_items['Cost: Unit (denominator)']) |\
        (cost_items['Flow: Unit (denominator)'] != cost_items['Operating Time: Unit'])
    cost_items = cost_items.loc[~check_units]
    check_units = cost_items.loc[check_units, : ]
    if check_units.shape[0] > 0:
        print("Warning: The following cost items need attention as the units are not harmonized ..")
        print(check_units)
    
    var_params = corr_params_variability.loc[corr_params_variability['col_param'].isin(['Cost Item']), : ]
            
    var_params_tbl = variability_table(var_params).reset_index(drop=True)
    var_params_tbl['variability_id'] = var_params_tbl.index
    
    
    cost_items_temp = cost_items.copy()

    cost_items = pd.DataFrame(columns=cost_items_temp.columns.to_list() + ['variability_id'])
    for r in range(0, var_params_tbl.shape[0]):
        cost_items_temp.loc[
            cost_items_temp[var_params_tbl.loc[r, 'col_param']].isin([var_params_tbl.loc[r, 'param_name']]), 
            var_params_tbl.loc[r, 'col_val']] = var_params_tbl.loc[r, 'param_value']
        cost_items_temp['variability_id'] = var_params_tbl.loc[r, 'variability_id']
        cost_items = pd.concat([cost_items, cost_items_temp])
    
    cost_items = cost_items.merge(var_params_tbl, how='left', on='variability_id').reset_index(drop=True)
    
    # Calculate itemized MFSP
    
    # re-calculate total cost
    cost_items.loc[cost_items['Flow'] != '-', 'Flow'] =\
        pd.to_numeric(cost_items.loc[cost_items['Flow'] != '-', 'Flow']).copy()
    cost_items.loc[cost_items['Operating Time'] != '-', 'Operating Time'] =\
        pd.to_numeric(cost_items.loc[cost_items['Operating Time'] != '-', 'Operating Time']).copy()
    cost_items.loc[cost_items['Unit Cost'] != '-', 'Unit Cost'] =\
        pd.to_numeric(cost_items.loc[cost_items['Unit Cost'] != '-', 'Unit Cost']).copy()
    
    cost_items.loc[((cost_items['Flow'] != '-') &
                       (cost_items['Operating Time'] != '-') &
                       (cost_items['Unit Cost'] != '-'))
        ,'Total Cost'] = \
        cost_items.loc[((cost_items['Flow'] != '-') &
                           (cost_items['Operating Time'] != '-') &
                           (cost_items['Unit Cost'] != '-')), 'Flow'] *\
        cost_items.loc[((cost_items['Flow'] != '-') &
                            (cost_items['Operating Time'] != '-') &
                            (cost_items['Unit Cost'] != '-')), 'Operating Time'] *\
        cost_items.loc[((cost_items['Flow'] != '-') &
                           (cost_items['Operating Time'] != '-') &
                           (cost_items['Unit Cost'] != '-')), 'Unit Cost']
    
# Correct for inflation to the year of study
cost_items['Cost Year'] = cost_items['Cost Year'].astype(int)
#cost_items['Total Cost'] = pd.to_numeric(cost_items['Total Cost'])
cost_items['Adjusted Total Cost'] = cost_items.apply(lambda x: cpi.inflate(x['Total Cost'], x['Cost Year'], to=cost_year), axis=1)
cost_items['Adjusted Cost Year'] = cost_year

# accounting calculations for one production year only for now
cost_items['Production Year'] = production_year
   
# Calculate itemized MFSP
cost_items['Itemized MFSP'] = cost_items['Adjusted Total Cost'].astype(float) / cost_items['Biofuel Flow'].astype(float)
cost_items['Itemized MFSP: Unit (numerator)'] = cost_items['Total Cost: Unit (numerator)']
cost_items['Itemized MFSP: Unit (denominator)'] = cost_items['Biofuel Flow: Unit (numerator)']

# Identify non-harmonized units if any
ignored_cost_items = cost_items.loc[cost_items['Total Flow: Unit (numerator)'] != cost_items['Cost: Unit (denominator)'], : ]
if ignored_cost_items.shape[0] > 0:
    print("Warning: The following cost items need attention as the units are not harmonized ..")
    print(ignored_cost_items)
ignored_cost_items = ignored_cost_items[['Case/Scenario', 'Parameter', 'Item',
                                         'Stream_Flow', 'Stream_LCA',
                                         'Total Flow: Unit (numerator)', 'Cost: Unit (denominator)']]


# For co-products we consider their cost as credit to the MFSP [co-product credit by displacement]
cost_items.loc[cost_items['Item'] == 'Coproducts', 'Itemized MFSP'] = \
  cost_items.loc[cost_items['Item'] == 'Coproducts', 'Itemized MFSP'] * -1

#%%
# Step: Calculate aggregrated Marginal Fuel Selling Price (MFSP)

MFSP_agg = cost_items.copy()

if consider_coproduct_cost_credit == False:
    MFSP_agg = MFSP_agg.loc[~MFSP_agg['Item'].isin(['Coproducts']), :]

if consider_variability_study:
    
    MFSP_agg = MFSP_agg[['Case/Scenario',
                         'Feedstock',
                         'Production Year',
                         'Itemized MFSP: Unit (numerator)', 
                         'Itemized MFSP: Unit (denominator)',
                         'Adjusted Cost Year',
                         'Itemized MFSP',
                         'variability_id',
                         'col_param',
                         'col_val',
                         'param_name',
                         'param_min',
                         'param_max',
                         'param_dist',
                         'dist_option',
                         'param_value']]
    MFSP_agg = MFSP_agg[MFSP_agg['Itemized MFSP'].notna()]
        
    MFSP_agg = MFSP_agg.groupby(['Case/Scenario',
                                 'Feedstock',
                                 'Production Year',
                                 'Itemized MFSP: Unit (numerator)', 
                                 'Itemized MFSP: Unit (denominator)',
                                 'Adjusted Cost Year',
                                 'variability_id',
                                 'col_param',
                                 'col_val',
                                 'param_name',
                                 'param_min',
                                 'param_max',
                                 'param_dist',
                                 'dist_option',
                                 'param_value']).agg({'Itemized MFSP' : 'sum'}).reset_index()
else:
    MFSP_agg = MFSP_agg[['Case/Scenario',
                         'Feedstock',
                         'Production Year',
                         'Itemized MFSP: Unit (numerator)', 
                         'Itemized MFSP: Unit (denominator)',
                         'Adjusted Cost Year',
                         'Itemized MFSP']]
    MFSP_agg = MFSP_agg[MFSP_agg['Itemized MFSP'].notna()]
        
    MFSP_agg = MFSP_agg.groupby(['Case/Scenario',
                                 'Feedstock',
                                 'Production Year',
                                 'Itemized MFSP: Unit (numerator)', 
                                 'Itemized MFSP: Unit (denominator)',
                                 'Adjusted Cost Year'
                                 ]).agg({'Itemized MFSP' : 'sum'}).reset_index()
    
MFSP_agg.rename(columns={'Itemized MFSP' : 'MFSP replacing fuel',
                         'Itemized MFSP: Unit (numerator)' : 'MFSP replacing fuel: Unit (numerator)',
                         'Itemized MFSP: Unit (denominator)' : 'MFSP replacing fuel: Unit (denominator)'}, inplace=True)

# Getting back the Final Product column
MFSP_agg = pd.merge(biofuel_yield[['Case/Scenario', 'Biofuel Stream_LCA']].drop_duplicates(), 
                    MFSP_agg, how='left', on='Case/Scenario').reset_index(drop=True)

# Save interim data tables
if save_interim_files == True:
    cost_items.to_csv(output_path_prefix + '/' + f_out_itemized_mfsp)
    MFSP_agg.to_csv(output_path_prefix + '/' + f_out_agg_mfsp)

#%%

# Step: Merge Itemized LCAs to TEA-pathway LCIs    
    
LCA_items = df_econ.loc[df_econ['Item'].isin(['Feedstock',
                                              'Purchased Inputs',
                                              'Coproducts',
                                              'Final Product',
                                              'CCS Stream',
                                              'Waste Disposal',
                                              ]), : ].copy()
LCA_items = LCA_items[['Case/Scenario', 
                       'Parameter', 
                       'Item', 
                       'Stream_Flow', 
                       'Stream_LCA',
                       'Flow: Unit (numerator)',
                       'Flow: Unit (denominator)',
                       'Flow',
                       'Operating Time: Unit',
                       'Operating Time',
                       'Operating Time (%)',
                       'Total Flow: Unit (numerator)',
                       'Total Flow: Unit (denominator)',
                       'Total Flow']]



# temporary value for production year
LCA_items['Production Year'] = production_year

# format LCI
corr_itemized_LCA = fmt_GREET_LCI(corr_itemized_LCA)

# Merge itemized LCAs to LCIs
LCA_items = pd.merge(LCA_items, corr_itemized_LCA, how='left', 
                     left_on=['Item', 'Stream_Flow', 'Stream_LCA', 'Production Year'],
                     right_on=['Item', 'Stream_Flow', 'Stream_LCA', 'Year']).reset_index(drop=True)

# harmonize units
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
    print("Warning: The following LCA items need attention as the units are not harmonized ..")
    print(ignored_LCA_items)
LCA_items = LCA_items.loc[~(LCA_items['Total Flow: Unit (numerator)'] != LCA_items['LCA: Unit (denominator)']), : ]

#%%

# Step: Itemized and aggregrated LCA metric per pathway

"""
# Some LCA mappings are unconfirmed or have marginal impact on CI estimation, hence ignored
LCA_items = LCA_items.loc[~LCA_items['Stream_LCA'].isin(['Makeup Water',
                                                        'Makeup water',
                                                        'Water make-up',
                                                        'Cooling Tower Makeup',
                                                        'Cooling Tower Make-up',
                                                        'Cooling tower water makeup',
                                                        'Cooling tower chemicals',
                                                        'Cooling Water Makeup']), :].reset_index(drop=True)
"""

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

# Save interim data tables
if save_interim_files == True:
    LCA_items.to_csv(output_path_prefix + '/' + f_out_itemized_LCA)
    LCA_items_agg.to_csv(output_path_prefix + '/' + f_out_agg_LCA)
    
#%%

# Step: Merge correspondence tables and GREET emission factors

# Merge aggregrated LCA metric to MFSP tables
MAC_df = pd.merge(MFSP_agg, LCA_items_agg, on=['Case/Scenario', 'Production Year']).reset_index(drop=True)

# map replaced fuels with replacing fuels
MAC_df = pd.merge(MAC_df, corr_replaced_replacing_fuel, how = 'left', 
               on=['Case/Scenario', 'Biofuel Stream_LCA', 'Feedstock']).reset_index(drop=True) 

# map replaced fuels with GREET pathways
MAC_df = pd.merge(MAC_df, corr_fuel_replaced_GREET_pathway, how='left', on=['Replaced Fuel']).reset_index(drop=True)
MAC_df.rename(columns={'GREET Pathway' : 'GREET Pathway for replaced fuel'}, inplace=True)

# map replaced fuels to their CIs
MAC_df = pd.merge(MAC_df,
                  corr_itemized_LCA,
                  left_on=['Item', 'Stream_Flow', 'Stream_LCA', 'Production Year'],
                  right_on=['Item', 'Stream_Flow', 'Stream_LCA', 'Year']).reset_index(drop=True)
MAC_df.drop(['Year', 'GREET1 sheet', 'Coproduct allocation method', 
             'GREET classification of coproduct'], axis=1, inplace=True)
MAC_df.rename(columns={'LCA: Unit (numerator)' : 'CI replaced fuel: Unit (Numerator)',
                       'LCA: Unit (denominator)' : 'CI replaced fuel: Unit (Denominator)',
                       'LCA_value' : 'CI replaced fuel',
                       'LCA_metric_x' : 'Metric_replacing fuel',
                       'LCA_metric_y' : 'Metric_replaced fuel'}, inplace=True)

"""
# harmonize emission factors of conventional fuels to CO2e unit
ef = ef_calc_co2e(ef)

# Merge aggregrated LCA metric to MFSP tables
MAC_df = pd.merge(MFSP_agg, LCA_items_agg, on=['Case/Scenario']).reset_index(drop=True)

# map replaced fuels with replacing fuels
MAC_df = pd.merge(MAC_df, corr_replaced_replacing_fuel, how = 'left', 
               on=['Case/Scenario', 'Biofuel Stream_LCA', 'Feedstock']).reset_index(drop=True) 

# map replaced fuels with GREET pathways
MAC_df = pd.merge(MAC_df, corr_fuel_replaced_GREET_pathway, how='left', on=['Replaced Fuel']).reset_index(drop=True)
MAC_df.rename(columns={'GREET Pathway' : 'GREET Pathway for replaced fuel'}, inplace=True)

# map GREET carbon intensities for replaced fuels, considering Decarb Model reference case carbon intensities only
MAC_df = pd.merge(MAC_df, ef.loc[ef['Case'] == 'Reference case', : ], how='left', 
                  left_on=['GREET Pathway for replaced fuel', 'LCA_metric', 'Production Year'],
                  right_on=['GREET Pathway', 'Formula', 'Year']).reset_index(drop=True)
MAC_df.rename(columns={'Stream_LCA' : 'Stream_LCA_replaced fuel',
                       'Formula' :'Formula_replaced fuel',
                       'Unit (Numerator)' : 'CI replaced fuel: Unit (Numerator)',
                       'Unit (Denominator)' : 'CI replaced fuel: Unit (Denominator)',
                       'Case' : 'Case_replaced fuel',
                       'Scope' : 'Scope_replaced fuel',
                       'Reference case' : 'CI replaced fuel',
                       'Elec0' : 'CI Elec0_replaced fuel'}, inplace=True)
MAC_df.drop(['GREET Pathway'], axis=1, inplace=True)
"""

# Map MFSP of replaced fuels
MAC_df = pd.merge(MAC_df,
                  corr_replaced_EIA_mfsp,
                  how='left',
                  on=['Replaced Fuel']).reset_index(drop=True)
MAC_df = pd.merge(MAC_df, 
                  EIA_price[['Year', 'Value', 'Energy carrier', 'Cost basis', 'Unit']],
                  how='left', 
                  left_on=['Production Year', 'EIA_fuel_mapping_for_price'], 
                  right_on=['Year', 'Energy carrier']).reset_index(drop=True)
MAC_df.rename(columns={'Value' : 'Cost_replaced fuel',
                       'Cost basis' : 'Cost basis_replaced fuel'}, inplace=True)

MAC_df[['Year_Cost_replaced fuel', 'Unit Cost_replaced fuel (Numerator)']] = MAC_df['Unit'].str.split(' ', n=1, expand = True)

MAC_df[['Unit Cost_replaced fuel (Numerator)', 
        'Unit Cost_replaced fuel (Denominator)']] = \
      MAC_df['Unit Cost_replaced fuel (Numerator)'].str.split('/', n=1, expand = True)

MAC_df.rename(columns={'Unit Cost_replaced fuel (Numerator)' : 'Cost replaced fuel: Unit (Numerator)', 
                       'Unit Cost_replaced fuel (Denominator)' : 'Cost replaced fuel: Unit (Denominator)'}, inplace=True)

MAC_df.drop(['Energy carrier', 'Unit'], axis=1, inplace=True)

"""
# Drop off data for which GREET pathways are not mapped until now
missing_GREET_mappings = MAC_df.loc[MAC_df['GREET Pathway for replaced fuel'].isna(), ['Case/Scenario', 'Biofuel Stream_LCA', 'Feedstock', 'Replaced Fuel']].drop_duplicates()
if missing_GREET_mappings.shape[0] > 0:
    print("Warning: The following pathways are currently dropped as their mappings to GREET CIs are not available as input ..")
    print(missing_GREET_mappings)
MAC_df = MAC_df.loc[~ MAC_df['GREET Pathway for replaced fuel'].isna(), :].copy()


# Assumption: non-liquid final products are skipped and not credited at the moment
# If classified as 'Coproducts', displacement based credit is considered for LCA and price credit is considered for MFSP 
MAC_df = MAC_df.loc[~ MAC_df['MFSP replacing fuel: Unit (denominator)'].isin(['lb']), : ].copy()

# dropping rows with no data on cost replaced fuel
MAC_df = MAC_df.loc[~MAC_df['Cost_replaced fuel'].isna(), :]
"""
#%%

# Step: Correct for inflation to the year of study

MAC_df['Year_Cost_replaced fuel'] = pd.to_numeric(MAC_df['Year_Cost_replaced fuel'])
MAC_df['Adjusted Cost_replaced fuel'] = \
    MAC_df.apply(lambda x: cpi.inflate(x['Cost_replaced fuel'], x['Year_Cost_replaced fuel'], to=production_year), axis=1)
MAC_df['Adjusted Cost Year'] = production_year
    
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
    
# # Convert fuel cost $/gallon to $/GGE
# # This conversion is done especially if certain calculations in future is required in GGE

# # Map Replaced fuel to 'GREET_Fuel', 'GREET_Fuel type' for GGE conversion
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

# Convert fuel cost from $ per GGE to $ per MJ
# extract CI of gasoline
tempdf = ob_units.hv_EIA.loc[(ob_units.hv_EIA['Energy carrier'] == 'Gasoline') &
                              (ob_units.hv_EIA['Energy carrier type'] == 'Petroleum Gasoline'), ['LHV', 'Unit']]
# convert unit
tempdf[['unit_numerator', 'unit_denominator']] = tempdf['Unit'].str.split('/', n=1, expand=True)
tempdf.drop(columns=['Unit'], inplace=True)
tempdf[['unit_numerator', 'LHV']] = \
    ob_units.unit_convert_df (
        tempdf[['unit_numerator', 'LHV']],
        Unit='unit_numerator',
        Value='LHV',
        if_unit_numerator=True,
        if_given_unit=True,
        given_unit='MJ').copy()
tempdf['unit_denominator'] = 'GGE'
# merge with MAC df for unit conversion
MAC_df = pd.merge(MAC_df, tempdf, how='left', 
                  left_on='Cost replaced fuel: Unit (Denominator)', 
                  right_on='unit_denominator').reset_index(drop=True)
MAC_df['Adjusted Cost_replaced fuel'] = MAC_df['Adjusted Cost_replaced fuel']/MAC_df['LHV'] # unit: $/MJ
MAC_df['Adjusted Cost replaced fuel: Unit (Denominator)'] = MAC_df['unit_numerator']
MAC_df['Adjusted Cost replaced fuel: Unit (Numerator)'] = 'USD'
MAC_df.drop(columns=['LHV', 'unit_numerator', 'unit_denominator'], inplace=True)


# Unit check for Replacing Fuel

# $/gal to $/GGE
# Map Replacing fuel to 'GREET_Fuel', 'GREET_Fuel type' for GGE conversion
# =============================================================================
# MAC_df = pd.merge(MAC_df, corr_GGE_GREET_fuel_replacing, how='left', 
#                   left_on=['Biofuel Stream_LCA'], 
#                   right_on=['B2B fuel name']).reset_index(drop=True)
#     
# MAC_df = pd.merge(MAC_df, ob_units.hv_EIA[['GREET_Fuel', 'GREET_Fuel type', 'GGE']].drop_duplicates(), 
#                   how='left', 
#                   on=['GREET_Fuel', 'GREET_Fuel type']).reset_index(drop=True)
# MAC_df.loc[MAC_df['MFSP replacing fuel: Unit (denominator)'] == 'gal', 'MFSP replacing fuel'] = \
#     MAC_df.loc[MAC_df['MFSP replacing fuel: Unit (denominator)'] == 'gal', 'MFSP replacing fuel'] / \
#     MAC_df.loc[MAC_df['MFSP replacing fuel: Unit (denominator)'] == 'gal', 'GGE']
# MAC_df.loc[MAC_df['MFSP replacing fuel: Unit (denominator)'] == 'gal', 'MFSP replacing fuel: Unit (denominator)'] = 'GGE'
# MAC_df.drop(columns=['GREET_Fuel', 'GREET_Fuel type', 'B2B fuel name', 'GGE'], inplace=True)
# 
# # Convert fuel cost from $ per GGE to $ per MJ
# # extract CI of gasoline
# tempdf = ob_units.hv_EIA.loc[(ob_units.hv_EIA['Energy carrier'] == 'Gasoline') &
#                              (ob_units.hv_EIA['Energy carrier type'] == 'Petroleum Gasoline'), ['LHV', 'Unit']]
# # convert unit
# tempdf[['unit_numerator', 'unit_denominator']] = tempdf['Unit'].str.split('/', n=1, expand=True)
# tempdf.drop(columns=['Unit'], inplace=True)
# tempdf[['unit_numerator', 'LHV']] = \
#     ob_units.unit_convert_df (
#         tempdf[['unit_numerator', 'LHV']],
#         Unit='unit_numerator',
#         Value='LHV',
#         if_unit_numerator=True,
#         if_given_unit=True,
#         given_unit='MJ').copy()
# tempdf['unit_denominator'] = 'GGE'
# # merge with MAC df for unit conversion
# MAC_df = pd.merge(MAC_df, tempdf, how='left', 
#                   left_on='MFSP replacing fuel: Unit (denominator)', 
#                   right_on='unit_denominator').reset_index(drop=True)
# MAC_df['MFSP replacing fuel'] = MAC_df['MFSP replacing fuel']/MAC_df['LHV'] # unit: $/MJ
# MAC_df.loc[MAC_df['MFSP replacing fuel: Unit (denominator)'] == 'GGE', 'MFSP replacing fuel: Unit (denominator)'] = \
#     MAC_df.loc[MAC_df['Total LCA: Unit (denominator)'] == 'GGE', 'unit_numerator']
# 
# MAC_df.loc[MAC_df['Total LCA: Unit (denominator)'] == 'GGE', 'Total LCA'] = \
#     MAC_df.loc[MAC_df['Total LCA: Unit (denominator)'] == 'GGE', 'Total LCA']/MAC_df['LHV'] # unit: g/MJ
# MAC_df.loc[MAC_df['Total LCA: Unit (denominator)'] == 'GGE', 'Total LCA: Unit (denominator)'] = \
#     MAC_df.loc[MAC_df['Total LCA: Unit (denominator)'] == 'GGE', 'unit_numerator']
#     
# MAC_df.drop(columns=['LHV', 'unit_numerator', 'unit_denominator'], inplace=True)
# =============================================================================

# unit check for CI replaced fuel
MAC_df[['CI replaced fuel: Unit (Denominator)', 'CI replaced fuel']] = \
    ob_units.unit_convert_df (
        MAC_df[['CI replaced fuel: Unit (Denominator)', 'CI replaced fuel']],
        Unit='CI replaced fuel: Unit (Denominator)',
        Value='CI replaced fuel',
        if_unit_numerator=False,
        if_given_unit=True,
        given_unit='MJ').copy()

#%%
# Step: Calculate MAC by Cost Items

# MAC = (MFSP_biofuel - MFSP_ref) / (CI_ref - CI_biofuel)
# Unit: ($/MJ - $/MJ) / (g/MJ - g/MJ) = $/g
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
    
#%%
# write data to the model dashboard tabs

if write_to_dashboard:
    
    with xw.App(visible=False) as app: 
        
        wb = xw.Book(input_path_TEA + '/' + f_TEA)
        
        if consider_variability_study:
            
            sheet_1 = wb.sheets['lca']
            sheet_1.range(str(4) + ':1048576').clear_contents()
            sheet_1['A4'].options(index=False, chunksize=10000).value =\
                LCA_items_agg[['Case/Scenario',
                               'LCA_metric',
                               'Total LCA: Unit (numerator)',
                               'Total LCA: Unit (denominator)',
                               'Production Year',
                               'Total LCA']]
            
            sheet_1 = wb.sheets['mfsp_var']
            sheet_1.range(str(4) + ':1048576').clear_contents()
            sheet_1['A4'].options(index=False, chunksize=10000).value =\
            MFSP_agg[['variability_id',
                      'col_param',
                      'col_val',
                      'param_name',
                      'param_min',
                      'param_max',
                      'param_dist',
                      'dist_option',
                      'param_value',
                      'Case/Scenario',	
                      'Production Year',                      
                      'MFSP replacing fuel: Unit (numerator)',
                      'MFSP replacing fuel: Unit (denominator)',
                      'MFSP replacing fuel',
                      'Adjusted Cost Year'                    
                      ]]
            
            sheet_1 = wb.sheets['mac_var']
            sheet_1.range(str(4) + ':1048576').clear_contents()
            sheet_1['A4'].options(index=False, chunksize=10000).value =\
            MAC_df[['variability_id',
                    'col_param',
                    'col_val',
                    'param_name',
                    'param_min',
                    'param_max',
                    'param_dist',
                    'dist_option',
                    'param_value',
                    'Case/Scenario',
                    'Biofuel Stream_LCA',
                    'Feedstock',
                    'Production Year',
                    'MFSP replacing fuel: Unit (numerator)',
                    'MFSP replacing fuel: Unit (denominator)',
                    'Adjusted Cost Year',
                    'MFSP replacing fuel',
                    'Metric_replacing fuel',	
                    'Total LCA: Unit (numerator)',
                    'Total LCA: Unit (denominator)',
                    'Total LCA',
                    'Replaced Fuel',
                    'Stream_Flow',
                    'Stream_LCA'	,
                    'CI replaced fuel: Unit (Numerator)',
                    'CI replaced fuel: Unit (Denominator)',
                    'CI replaced fuel',
                    'Metric_replaced fuel',
                    'EIA_fuel_mapping_for_price',
                    'Year_Cost_replaced fuel',
                    'Cost basis_replaced fuel',
                    'Year_Cost_replaced fuel',
                    'Cost replaced fuel: Unit (Numerator)',
                    'Cost replaced fuel: Unit (Denominator)',
                    'Adjusted Cost_replaced fuel',
                    'Adjusted Cost replaced fuel: Unit (Denominator)',
                    'Adjusted Cost replaced fuel: Unit (Numerator)',
                    'MAC_calculated',
                    'MAC_calculated: Unit (numerator)',
                    'MAC_calculated: Unit (denominator)'
                    
                    ]]
            
            sheet_1 = wb.sheets['lca_itm']
            sheet_1.range(str(4) + ':1048576').clear_contents()
            sheet_1['A4'].options(index=False, chunksize=10000).value =\
                LCA_items[['Case/Scenario', 
                           'Parameter', 
                           'Item', 
                           'Stream_Flow', 
                           'Stream_LCA',
                           'Flow: Unit (numerator)', 
                           'Flow: Unit (denominator)', 
                           'Flow',
                           'Operating Time: Unit', 
                           'Operating Time', 
                           'Operating Time (%)',
                           'Total Flow: Unit (numerator)', 
                           'Total Flow: Unit (denominator)',
                           'Total Flow', 
                           'Production Year',
                           #'Year', 
                           #'GREET1 sheet',
                           #'Coproduct allocation method', 
                           #'GREET classification of coproduct',
                           'LCA: Unit (numerator)',
                           'LCA: Unit (denominator)',
                           'LCA_value',
                           'LCA_metric',
                           'Total LCA',
                           'Total LCA: Unit (numerator)',
                           'Total LCA: Unit (denominator)',
                           #'Biofuel Flow: Unit (numerator)',
                           #'Biofuel Flow: Unit (denominator)',
                           #'Biofuel Flow'
                           ]]
                
            sheet_1 = wb.sheets['mfsp_itm_var']
            sheet_1.range(str(4) + ':1048576').clear_contents()
            sheet_1['A4'].options(index=False, chunksize=10000).value =\
                 cost_items[['variability_id',
                             'col_param',
                             'col_val',
                             'param_name',
                             'param_min',
                             'param_max',
                             'param_dist',
                             'dist_option',
                             'param_value',
                             'Case/Scenario', 
                             'Parameter', 
                             'Item', 
                             'Stream_Flow',
                             'Stream_LCA',
                             'Flow: Unit (numerator)',
                             'Flow: Unit (denominator)',
                             'Flow',
                             'Cost Item',
                             'Cost: Unit (numerator)', 
                             'Cost: Unit (denominator)',
                             'Unit Cost',
                             'Operating Time: Unit',
                             'Operating Time',
                             'Operating Time (%)', 
                             'Total Cost: Unit (numerator)',
                             'Total Cost: Unit (denominator)',
                             'Total Cost',
                             'Total Flow: Unit (numerator)',
                             'Total Flow: Unit (denominator)',
                             'Total Flow', 
                             'Cost Year',
                             # 'Feedstock Stream_Flow',
                             # 'Feedstock',
                             # 'Feedstock Flow: Unit (numerator)',
                             # 'Feedstock Flow: Unit (denominator)', 
                             # 'Feedstock Flow',
                             # 'Biofuel Flow: Unit (numerator)',
                             # 'Biofuel Flow: Unit (denominator)',
                             # 'Biofuel Flow',
                             'Adjusted Total Cost', 
                             'Adjusted Cost Year',
                             'Production Year', 
                             'Itemized MFSP', 
                             'Itemized MFSP: Unit (numerator)',
                             'Itemized MFSP: Unit (denominator)'
                             ]]
                
        else:
            
            sheet_1 = wb.sheets['lca']            
            sheet_1.range(str(4) + ':1048576').clear_contents()
            sheet_1['A4'].options(index=False, chunksize=10000).value =\
                LCA_items_agg[['Case/Scenario',
                               'LCA_metric',
                               'Total LCA: Unit (numerator)',
                               'Total LCA: Unit (denominator)',
                               'Production Year',
                               'Total LCA']]
            
            sheet_1 = wb.sheets['mfsp']
            sheet_1.range(str(4) + ':1048576').clear_contents()
            sheet_1['A4'].options(index=False, chunksize=10000).value =\
            MFSP_agg[['Case/Scenario',	
                      'Production Year', 
                      'MFSP replacing fuel: Unit (numerator)',
                      'MFSP replacing fuel: Unit (denominator)',
                      'MFSP replacing fuel',
                      'Adjusted Cost Year'
                      ]]
            
            sheet_1 = wb.sheets['mac']
            sheet_1.range(str(4) + ':1048576').clear_contents()
            sheet_1['A4'].options(index=False, chunksize=10000).value =\
            MAC_df[['Case/Scenario',
                    'Biofuel Stream_LCA',
                    'Feedstock',
                    'Production Year',
                    'MFSP replacing fuel: Unit (numerator)',
                    'MFSP replacing fuel: Unit (denominator)',
                    'Adjusted Cost Year',
                    'MFSP replacing fuel',
                    'Metric_replacing fuel',	
                    'Total LCA: Unit (numerator)',
                    'Total LCA: Unit (denominator)',
                    'Total LCA',
                    'Replaced Fuel',
                    'Stream_Flow',
                    'Stream_LCA'	,
                    'CI replaced fuel: Unit (Numerator)',
                    'CI replaced fuel: Unit (Denominator)',
                    'CI replaced fuel',
                    'Metric_replaced fuel',
                    'EIA_fuel_mapping_for_price',
                    'Year_Cost_replaced fuel',
                    'Cost basis_replaced fuel',
                    'Year_Cost_replaced fuel',
                    'Cost replaced fuel: Unit (Numerator)',
                    'Cost replaced fuel: Unit (Denominator)',
                    'Adjusted Cost_replaced fuel',
                    'Adjusted Cost replaced fuel: Unit (Denominator)',
                    'Adjusted Cost replaced fuel: Unit (Numerator)',
                    'MAC_calculated',
                    'MAC_calculated: Unit (numerator)',
                    'MAC_calculated: Unit (denominator)'
                    ]]
            
            sheet_1 = wb.sheets['lca_itm']
            sheet_1.range(str(4) + ':1048576').clear_contents()
            sheet_1['A4'].options(index=False, chunksize=10000).value =\
                LCA_items[['Case/Scenario', 
                           'Parameter', 
                           'Item', 
                           'Stream_Flow', 
                           'Stream_LCA',
                           'Flow: Unit (numerator)', 
                           'Flow: Unit (denominator)', 
                           'Flow',
                           'Operating Time: Unit', 
                           'Operating Time', 
                           'Operating Time (%)',
                           'Total Flow: Unit (numerator)', 
                           'Total Flow: Unit (denominator)',
                           'Total Flow', 
                           'Production Year',
                           #'Year', 
                           #'GREET1 sheet',
                           #'Coproduct allocation method', 
                           #'GREET classification of coproduct',
                           'LCA: Unit (numerator)',
                           'LCA: Unit (denominator)',
                           'LCA_value',
                           'LCA_metric',
                           'Total LCA',
                           'Total LCA: Unit (numerator)',
                           'Total LCA: Unit (denominator)',
                           #'Biofuel Flow: Unit (numerator)',
                           #'Biofuel Flow: Unit (denominator)',
                           #'Biofuel Flow'
                           ]]
                
            sheet_1 = wb.sheets['mfsp_itm']
            sheet_1.range(str(4) + ':1048576').clear_contents()
            sheet_1['A4'].options(index=False, chunksize=10000).value =\
                 cost_items[['Case/Scenario', 
                             'Parameter', 
                             'Item', 
                             'Stream_Flow',
                             'Stream_LCA',
                             'Flow: Unit (numerator)',
                             'Flow: Unit (denominator)',
                             'Flow',
                             'Cost Item',
                             'Cost: Unit (numerator)', 
                             'Cost: Unit (denominator)',
                             'Unit Cost',
                             'Operating Time: Unit',
                             'Operating Time',
                             'Operating Time (%)', 
                             'Total Cost: Unit (numerator)',
                             'Total Cost: Unit (denominator)',
                             'Total Cost',
                             'Total Flow: Unit (numerator)',
                             'Total Flow: Unit (denominator)',
                             'Total Flow', 
                             'Cost Year',
                             # 'Feedstock Stream_Flow',
                             # 'Feedstock',
                             # 'Feedstock Flow: Unit (numerator)',
                             # 'Feedstock Flow: Unit (denominator)', 
                             # 'Feedstock Flow',
                             # 'Biofuel Flow: Unit (numerator)',
                             # 'Biofuel Flow: Unit (denominator)',
                             # 'Biofuel Flow',
                             'Adjusted Total Cost', 
                             'Adjusted Cost Year',
                             'Production Year', 
                             'Itemized MFSP', 
                             'Itemized MFSP: Unit (numerator)',
                             'Itemized MFSP: Unit (denominator)'
                             ]]
        wb.save()
        wb.close()
        
#%%
