# -*- coding: utf-8 -*-
"""
Copyright © 2025, UChicago Argonne, LLC
The full description is available in the LICENSE file at location:
    https://github.com/Saurajyoti/BestBiomassUse/blob/master/LICENSE

@Project: Best Use of Biomass
@Title: Main script to call data processing scripts, process data, perform calculations, and save output files
@Authors: Saurajyoti Kar
@Contact: skar@anl.gov
@Affiliation: Argonne National Laboratory

Created on Wed Jan  4 10:32:14 2023

"""

'''
This is the main script to call data processing scripts, process data, perform calculations, and 
save output files. This version of main implements itemized LCA assessment of biofuel pathways.
'''

# %%
# Declare data input and other parameters

import numbers
from collections import Counter
from xlwings.constants import AutoFillType
from xlwings.constants import DeleteShiftDirection
import xlwings as xw
import cpi
from datetime import datetime
import os
import numpy as np
import pandas as pd

code_path_prefix = 'C:/Users/skar/repos/BestBiomassUse'

input_path_prefix = 'C:/Users/skar/repos/BestBiomassUse/data'
input_path_decarb_model = 'C:/Users/skar/Box/EERE SA Decarbonization/1. Tool/EERE Tool/Dashboard'

output_path_prefix = 'C:/Users/skar/repos/BestBiomassUse/data/interim'

# Excel worksheet row limit
EXCEL_MAX_ROWS = 1_048_576

input_path_model = input_path_prefix + '/model'
input_path_GREET = input_path_prefix + '/GREET'
input_path_EIA_price = input_path_prefix + '/EIA'
input_path_corr = input_path_prefix + '/correspondence_files'
input_path_units = input_path_prefix + '/Units'
input_path_BT16 = input_path_prefix + '/BT16'

# path to the Github local repository
os.chdir(code_path_prefix)
from unit_conversions import model_units

# f_model = 'MCCAM_09_11_2023_working.xlsx'

f_model = 'MCCAM_04_10_2025_working.xlsx'

sheet_TEA = 'Db'
sheet_param_variability = 'var_p'
sheet_name_lists = 'lists'

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
# f_corr_fuel_replacing_GREET_pathway = 'corr_fuel_replacing_GREET_pathway.csv'
f_corr_GGE_GREET_fuel_replaced = 'corr_GGE_GREET_fuel_replaced.csv'
f_corr_GGE_GREET_fuel_replacing = 'corr_GGE_GREET_fuel_replacing.csv'
f_corr_itemized_LCI = 'corr_LCI_GREET_temporal_11_15_2023.csv'
f_corr_replaced_EIA_mfsp = 'corr_replaced_mfsp.csv'

f_Decarb_Model = 'US Decarbonization Model - Dashboard.xlsx'

f_BT16_availability = 'B2B resource availability.xlsx'
sheet_BT16_availability = 'Sheet1'

# Year(s) of production defined as a list:
# If single year: [year]
# If multiple year: [first year, last year], inclusive of both the bounds
# production_year = [2022]
production_year = [2022, 2050]

# cost adjustment year, to which inflation will be adjusted
cost_year = 2020

# cost year of BT16 feedstock availability data set
BT16_cost_year = 2014

# Toggle cost credit for coproducts while calculating aggregrated MFSP
consider_coproduct_cost_credit = True

# Toggle to control emissions credit for coproducts while calculating aggregrated CIs
consider_coproduct_env_credit = True

# Toggle variability study
consider_variability_study = True # When true toggle scale-up study below to False 
# Selection to either run 'Cost_Item' or 'Stream_LCA' variabilities so that MAC is calculated with one parameters type varied at a time.
consider_which_variabilities = 'Stream_LCA' 

# Toggle writing output to interim files
save_interim_files = True

# Toggle write output to the dashboard workbook
write_to_dashboard = True

# Toggle scale-up analysis,
# Only run scale-up analysis with consider_variability_study = False
consider_scale_up_study = False # True when doing scale-up study, otherwise False

# Toggle implementing Decarb Model electric grid carbon intensity
decarb_electric_grid = False

# Scenario of decarb electric grid CI, when different from Decarb Model
decarb_grid_scenario1 = False
decarb_grid_scenario1_values = [1E-20, 140000]  # [min, max], g per mmBtu

# Toggle True to calibrate biopower scenarios baseline (CI and Marginal Cost) for MAC calculations
# to the baseline scenario from the data source report in place of the grid [True for QA purpose only]
adjust_biopower_baseline = False

# This controls if CO2 w/ C from VOC and CO gets calculated even if such value
# is present in emission factor table. Some instances of EF table may not have
# it already calculated accurate, so please keep it always True, unless for QA
always_calc_CO2_w_VOC_CO = True

# Toggle on/off to harmonize CCS, fossil and combustion emissions, Fossil
harmonize_CCS_fossil = True

# Define the type of allocation performed:
# Pathway: all energy products 'Fuel Use' are summed up as energy product and share the same pathway carbon intensity value
# Energy: all energy products are considered, primary fuel energy fraction is considerd for allocation of GHG emissions and costs
"""
Hybrid: One primary fuel is declared. Energy allocation of inputs flow rates, 
coproduct flow rates are done accross all liquid fuel products. Non-primary fuel types contribute to cr"""

allocation_type = 'Energy'  # Pathway, Hybrid, Energy

dict_gco2e = {  # Table 2, AR6/GWP100, GREET1 2022
    'CO2': 1,
    'CO2 (w/ C in VOC & CO)': 1,
    'N2O': 273,
    'CH4': 29.8,
    # 'Biogenic CH4' : 29.8,
    'VOC': 0,
    'CO': 0,
    'NOx': 0,
    'BC': 0,
    'OC': 0
}

dict_frac_C = {
    'Carbon ratio of VOC': 0.85,
    'Carbon ratio of CO': 0.43,
    'Carbon ratio of CH4': 0.75,
    'Carbon ratio of CO2': 0.27,
    'Sulfur ratio of SO2': 0.50
}

# List of Stream_Flows that have biogenic C and their CO2 is not excluded
biogenic_lci = ['Biochar',
                'Flue gas',
                ]
biogenic_emissions = ['Biogenic CO2',
                      'Biogenic CH4'  # biogenic CH4 are not considered to have environmental effect
                      ]

# %%
# import packages

# cpi.update()

# Import user defined modules
os.chdir(code_path_prefix)

# %% Customize the Excel Instance

class ExcelApp(xw.App):
    """override xw.App default properties"""
    calculation = 'manual'
    display_alerts = False
    enable_events = False
    screen_updating = False
    visible = False

# %%
# User defined function definitions

# Function to expand on the input variability table
def variability_table(var_params):
    # Create an empty list to collect all rows
    all_rows = []

    for r in range(var_params.shape[0]):
        param_min = var_params.loc[r, 'param_min']
        param_max = var_params.loc[r, 'param_max']
        dist_option = var_params.loc[r, 'dist_option']
        param_dist = var_params.loc[r, 'param_dist']
        
        # Check for linear distribution
        if param_dist == 'linear':
            val = param_min
            while val <= param_max:
                all_rows.append({
                    'col_param': var_params.loc[r, 'col_param'],
                    'col_val': var_params.loc[r, 'col_val'],
                    'param_name': var_params.loc[r, 'param_name'],
                    'param_min': param_min,
                    'param_max': param_max,
                    'param_dist': param_dist,
                    'dist_option': dist_option,
                    'param_value': val
                })
                val += dist_option

    # Convert list of rows into a DataFrame
    var_tbl = pd.DataFrame(all_rows, columns=var_params.columns.to_list() + ['param_value'])
    
    return var_tbl


def mult_numeric(a, b, c):
    if ((type(a) is int) | (type(a) is float)) &\
       ((type(b) is int) | (type(b) is float)) &\
       ((type(c) is int) | (type(c) is float)):
        return a*b*c
    else:
        return

# Function to add header rows to LCA metric rows, select subset of LCA metrices, and calculate CO2e
def fmt_GREET_LCI(df):

    harmonize_headers = {

        'Energy demand': [
            'Energy demand',
            'Energy: mmBtu/ton',
            'Energy Use: mmBtu/ton of product',
            'Energy Use: mmBtu/ton',
            'Energy: Btu/g of material throughput, except as noted',
            'Energy use: Btu/mmBtu of fuel throughput (except as noted)',
            'Energy Use: mmBtu per ton',
            'Energy Consumption: Btu/mmBtu of fuel transported',
            'Energy use: Btu/gal treated',
            'Total energy, Btu',
            'Energy Use: MJ per MJ'
        ],

        'Water consumption': [
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
            'Water consumption, gallons/mmBtu of fuel throughput',
            'Water consumption: gallons per MJ',
            
        ],
        'Total emissions': [
            'Total emissions',
            'Total Emissions: grams/ton',
            'Total emissions: grams/mmBtu of fuel throughput, except as noted',
            'Total emissions: grams/g of material throughput, except as noted',
            'Total Emissions: grams per ton',
            'Total Emissions: grams/mmBtu fuel transported',
            'Total Emissions: grams/mmBtu of fuel transported',
            'Total emissions: grams/gal treated',
            'Total Emissions: grams/mmBtu of fuel throughput, except as noted',
            'Total emissions: grams',
            'Total emissions: grams/mmBtu of fuel throughput',
            'Total Emissions: grams per MJ',
        ],

        'Urban emissions': [
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
            'Urban Emissions: grams per ton',
            'Urban emissions: grams/mmBtu of fuel throughput',
            'Urban Emissions: grams per MJ'
        ],
    }

    harmonize_metric = {

        'Total energy': [
            'Total Energy',
            'Total energy',
        ],

        'Fossil fuels': [
            'Fossil fuels',
            'Fossil Fuels',
            'Fossil energy',
            'Fossil fuels, Btu',
            'Fossi lfuels',
        ],

        'Coal': [
            'Coal',
            'Coal, Btu',
        ],

        'Natural gas': [
            'Natural gas',
            'Natural Gas',
            'Natural gas, Btu',
        ],

        'Petroleum': [
            'Petroleum',
            'Petroleum, Btu',
        ],

        'VOC': [
            'VOC',
            'Urban VOC',
            'VOC from bulk terminal',
            'VOC from ref. Station',
            'VOC from refueling station',
            'VOC: Total',
            'VOC: Urban',
        ],

        'CO': [
            'CO',
            'Urban CO',
            'CO: Total',
            'CO: Urban',
        ],

        'NOx': [
            'NOx',
            'Urban NOx',
            'NOx: Total',
            'NOx: Urban',
        ],

        'PM10': [
            'PM10',
            'Urban PM10',
            'PM10: Total',
            'PM10: Urban',
        ],

        'PM2.5': [
            'PM2.5',
            'Urban PM2.5',
            'PM2.5: Total',
            'PM2.5: Urban',
        ],

        'SOx': [
            'SOx',
            'Urban SOx',
            'SOx: Total',
            'SOx: Urban',
        ],

        'BC': [
            'BC',
            'Urban BC',
            'BC Total',
            'BC: Urban',
            'BC, Total'
        ],

        'OC': [
            'OC',
            'Urban OC',
            'OC Total',
            'OC: Urban',
            'OC, Total',
        ],

        'CH4': [
            'CH4',
            'CH4: combustion',
        ],

        'N2O': [
            'N2O',
        ],

        'CO2': [
            'CO2',
            'Misc. CO2',
        ],

        'CO2 (w/ C in VOC & CO)': [
            'CO2 (w/ C in VOC & CO)',
        ],

        'GHGs': [
            'GHGs (grams/ton)',
            'GHGs',
        ],

        'Other GHG Emissions': [
            'Other GHG Emissions',
        ],

        'Biogenic CH4': [
            'Biogenic CH4',
        ],

        'Biogenic CO2': [
            'Biogenic CO2',
        ],

    }

    harmonize_headers_long = {v1: k for k, v_list in harmonize_headers.items() for v1 in v_list}
    harmonize_metric_long = {v1: k for k, v_list in harmonize_metric.items() for v1 in v_list}

    # testing only
    # df = corr_itemized_LCA.copy()
    # df = df.loc[(df['Parameter_B'] == 'Avoided Ems Credits') & (df['Stream_Flow'] == 'Sludge (dry basis), Counterfactual'), : ].reset_index(drop=True)
    # df = df.loc[(df['Stream_Flow'] == 'Jet Range') & (df['Stream_LCA'] == 'SAF: jet fuel T&D, combustion emissions'), : ].reset_index(drop=True)

    df = df.drop_duplicates().reset_index(drop=True)

    df.rename(columns={'GREET row names_level1': 'LCA_metric_GREET', 'values_level1': 'LCA_value'}, inplace=True)

    # drop rows with loss factor
    df = df.loc[df['LCA_metric_GREET'].ne('Loss factor')].reset_index(drop=True)

    df[['LCA: Unit (numerator)', 'LCA: Unit (denominator)']] = df['Unit'].str.split('/', expand=True)

    df = df[[
        'Parameter_B',
        'Stream_Flow',
        'Stream_LCA',
        'GREET1 sheet',
        'Coproduct allocation method',
        'GREET classification of coproduct',
        'LCA_metric_GREET',
        'LCA_value',
        'LCA: Unit (numerator)',
        'LCA: Unit (denominator)',
        'Year'
        ]]

    # strip white spaces before and after metric names
    df['LCA_metric_GREET'] = df['LCA_metric_GREET'].str.strip()

    # replace with harmonized headers and metrices
    full_mapping = {**harmonize_headers_long, **harmonize_metric_long}
    df['LCA_metric'] = df['LCA_metric_GREET'].map(full_mapping).fillna('-')
            
    # Join header to LCI metric   
    df['header_val'] = df['LCA_metric'].where(df['LCA_metric'].isin(harmonize_headers)).ffill()
    df['LCA_metric'] = df['header_val'] + '__' + df['LCA_metric']

    # remove the header rows except water consumption
    df = df.loc[~(df['LCA_metric'].isin(
    set(harmonize_headers.keys()) - {'Water consumption'})), :].reset_index(drop=True)

    # select a subset of LCA metrices for GHG calculation
    select_GHG_metrices = [
        'Total emissions__VOC',
        'Total emissions__CO',
        'Total emissions__NOx',
        # 'Total emissions__PM10',
        # 'Total emissions__PM2.5',
        # 'Total emissions__SOx',
        'Total emissions__BC',
        'Total emissions__OC',
        'Total emissions__CH4',
        'Total emissions__N2O',
        'Total emissions__CO2',
        'Total emissions__CO2 (w/ C in VOC & CO)',
        # 'Total emissions__GHGs',
        'Total emissions__Biogenic CH4',
        # 'Total emissions__Biogenic CO2'
    ]

    col_indices = ['Parameter_B', 'Stream_Flow', 'Stream_LCA', 'GREET1 sheet',
                   'Coproduct allocation method', 'GREET classification of coproduct',
                   'LCA_metric', 'LCA: Unit (numerator)', 'LCA: Unit (denominator)', 'Year']
    #cols_duplicate_check = ['Case/Scenario', 'Parameter_B', 'Stream_Flow', 'Stream_LCA', 'Year']

    df = df.loc[df['LCA_metric'].isin(select_GHG_metrices), :]

    df.drop(columns=['LCA_metric_GREET'], inplace=True)

    df = df.loc[~df['LCA_value'].isna(), :]

    df['LCA_value'] = df['LCA_value'].astype('float')

    df['LCA: Unit (denominator)'] = df['LCA: Unit (denominator)'].fillna(
        '-', inplace=False)

    # harmonize units
    # GREET tonnes represent Short Ton
    df['LCA: Unit (denominator)'] = ['Short Tons' if val ==
                                     'ton' else val for val in df['LCA: Unit (denominator)']]

    # convert LCA unit of flow to model standard unit
    df.loc[:, ['LCA: Unit (denominator)', 'LCA_value']] = \
        ob_units.unit_convert_df(df.loc[:, ['LCA: Unit (denominator)', 'LCA_value']],
                                 Unit='LCA: Unit (denominator)', Value='LCA_value',
                                 if_unit_numerator=False, if_given_category=False)

    # identify duplicate rows
    df_duplicates = df.loc[df[col_indices].duplicated(), :].copy()
    if df_duplicates.shape[0] > 0:
        print("Warning: Note certain LCA metrices are duplicates. The duplicates arrise due to harmonizing LCA metrices. Duplicates with all data considered are already removed after file is read.")
        print(df_duplicates)
    df = df.groupby(col_indices, dropna=False, sort=False).agg(
        {'LCA_value': 'sum'}).reset_index()

    df = df.pivot(index=['Parameter_B', 'Stream_Flow', 'Stream_LCA', 'GREET1 sheet',
                         'Coproduct allocation method', 'GREET classification of coproduct',
                         'LCA: Unit (numerator)', 'LCA: Unit (denominator)', 'Year'],
                  columns='LCA_metric',
                  values='LCA_value').reset_index()
    
    # GHG columns selection and NA filling to 0
    df[list(set(select_GHG_metrices) - set(df.columns.values))] = 0
    df[select_GHG_metrices] = df[select_GHG_metrices].fillna(0)
   
    # Check and calculate CO2 (w/C in VOC & CO)
    # if co2 == 0: don't re-calculate CO2 w/C.
    mask = (df['Total emissions__CO2'] != 0) & (always_calc_CO2_w_VOC_CO)
    df.loc[mask, 'Total emissions__CO2 (w/ C in VOC & CO)'] = (
        df.loc[mask, 'Total emissions__CO2'] +
        df.loc[mask, 'Total emissions__VOC'] * dict_frac_C['Carbon ratio of VOC'] / dict_frac_C['Carbon ratio of CO2'] +
        df.loc[mask, 'Total emissions__CO'] * dict_frac_C['Carbon ratio of CO'] / dict_frac_C['Carbon ratio of CO2']
    )


    # Calculate GHG metric
    df['GHG'] = df['Total emissions__CO2 (w/ C in VOC & CO)'] +\
        df['Total emissions__CH4'] * dict_gco2e['CH4'] +\
        df['Total emissions__N2O'] * dict_gco2e['N2O'] +\
        df['Total emissions__VOC'] * dict_gco2e['VOC'] +\
        df['Total emissions__CO'] * dict_gco2e['CO'] +\
        df['Total emissions__NOx'] * dict_gco2e['NOx'] +\
        df['Total emissions__BC'] * dict_gco2e['BC'] +\
        df['Total emissions__OC'] * dict_gco2e['OC'] -\
        df['Total emissions__Biogenic CH4'] * \
        dict_frac_C['Carbon ratio of CH4'] / dict_frac_C['Carbon ratio of CO2']

    df = df[['Parameter_B',
             'Stream_Flow',
             'Stream_LCA',
             'GREET1 sheet',
             'Coproduct allocation method',
             'GREET classification of coproduct',
             'LCA: Unit (numerator)',
             'LCA: Unit (denominator)',
             'Year',
             # 'Total emissions__BC',
             # 'Total emissions__Biogenic CH4',
             # 'Total emissions__CH4',
             # 'Total emissions__CO',
             # 'Total emissions__CO2',
             # 'Total emissions__CO2 (w/ C in VOC & CO)',
             # 'Total emissions__N2O',
             # 'Total emissions__NOx',
             # 'Total emissions__OC',
             # 'Total emissions__VOC',
             'GHG',
             ]]

    df.rename(columns={'GHG': 'LCA_value'}, inplace=True)
    df['LCA_metric'] = 'CO2e'

    return df_duplicates, df


def ef_calc_co2e(df):

    df['Reference case'] = pd.to_numeric(df['Reference case'], errors='coerce')
    df['Elec0'] = pd.to_numeric(df['Elec0'], errors='coerce')

    # Multiply by the corresponding factor from dict_gco2e
    df['mult'] = df['Formula'].map(dict_gco2e)
    
    # Calculate CO2e for 'Reference case' and 'Elec0'
    df['Reference case'] *= df['mult']
    df['Elec0'] *= df['mult']

    df = df.groupby(['GREET Pathway', 'Unit (Numerator)',
                     'Unit (Denominator)', 'Case', 'Scope', 'Year'], as_index=False).agg({
                         'Reference case': 'sum',
                         'Elec0': 'sum'
                     })
    df['Formula'] = 'CO2e'
    df['Stream_LCA'] = 'Carbon dioxide equivalent'
    
    return df


# %%
# Step: Load data file and select columns for computation

init_time = datetime.now()

df_econ = pd.read_excel(input_path_model + '/' + f_model, sheet_name=sheet_TEA, header=3, index_col=None,
                        dtype={'Case/Scenario': str,
                               'Parameter_A': str,
                               'Parameter_B': str,
                               'Stream_Flow': str,
                               'Stream_LCA': str,
                               'Energy_alloc_primary_fuel': str,
                               'Flow: Unit (numerator)': str,
                               'Flow: Unit (denominator)': str,
                               'Flow': float,
                               'Cost Item': str,
                               'Cost: Unit (numerator)': str,
                               'Cost: Unit (denominator)': str,
                               'Unit Cost': float,
                               'Operating Time: Unit': str,
                               'Operating Time': float,
                               'Operating Time (%)': float,
                               'Total Cost: Unit (numerator)': str,
                               'Total Cost: Unit (denominator)': str,
                               'Total Cost': float,
                               'Total Flow: Unit (numerator)': str,
                               'Total Flow: Unit (denominator)': str,
                               'Total Flow': float,
                               'Cost Year': float},
                        na_values=['-'])

pathway_names = pd.read_excel(
    input_path_model + '/' + f_model, sheet_name=sheet_name_lists, header=3, usecols='B:H')

df_econ = df_econ[['Case/Scenario',
                   'Parameter_A',
                   'Parameter_B',
                   'Stream_Flow',
                   'Stream_LCA',
                   'Energy_alloc_primary_fuel',
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
                   'Cost Year']]

# Biopower pathways
biopower_scenarios = [

    ###
    'Baseline for Biopower, 100% coal, w/o CCS, 650 MWe',
    'Baseline for Biopower, 100% coal, w/ CCS, 650 MWe',

    'Biopower: 80% coal, w/o BECCS, 650 MWe',
    'Biopower: 80% coal, w/ BECCS, 650 MWe',

    'Biopower: 100% biomass, w/o BECCS, 130 MWe',
    'Biopower: 100% biomass, w/ BECCS, 130 MWe',
    ###
    # 'Biopower: 51% coal, w/ BECCS, 650 MWe',
    # 'Biopower: 100% biomass, w/o BECCS, 650 MWe',
    # 'Biopower: 100% biomass, w/ BECCS, 650 MWe',
    # 'Biopower: 51% coal, w/o BECCS, 650 MWe',

]

# Select pathways to consider
pathways_to_consider = [

    ###
    '2020, 2019 SOT High Octane Gasoline from Lignocellulosic Biomass via Syngas and Methanol/Dimethyl Ether Intermediates',
    ###

    # Tan et al., 2016 pathways
    ###
    'Pathway 1A: Syngas to molybdenum disulfide (MoS2)-catalyzed alcohols followed by fuel production via alcohol condensation (Guerbet reaction), dehydration, oligomerization, and hydrogenation',
    'Pathway 1B: Syngas fermentation to ethanol followed by fuel production via alcohol condensation (Guerbet reaction), dehydration, oligomerization, and hydrogenation',
    'Pathway 2A: Syngas to rhodium (Rh)-catalyzed mixed oxygenates followed by fuel production via carbon coupling/deoxygenation (to isobutene), oligomerization, and hydrogenation',
    'Pathway 2B: Syngas fermentation to ethanol followed by fuel production via carbon coupling/deoxygenation (to isobutene), oligomerization, and hydrogenation',
    'Pathway FT: Syngas to liquid fuels via Fischer-Tropsch technology as a commercial benchmark for comparisons',
    ###

    # Decarb 2b pathways
    # 'Thermochemical Research Pathway to High-Octane Gasoline Blendstock Through Methanol/Dimethyl Ether Intermediates',
    # 'Cellulosic Ethanol',

    ###
    'Decarb 2b: Cellulosic Ethanol to renewable gasoline and jet fuels',
    'Decarb 2b: Cellulosic Ethanol to renewable gasoline and jet fuels with CCS of fermentation offgas CO2',
    'Decarb 2b: Cellulosic Ethanol to renewable gasoline and jet fuels with CCS of fermentation offgas and boiler vent streams CO2',
    ###

    # 'Decarb 2b: Cellulosic Ethanol to renewable gasoline and jet fuels_jet',
    # 'Decarb 2b: Cellulosic Ethanol to renewable gasoline and jet fuels with CCS of fermentation offgas CO2_jet',
    # 'Decarb 2b: Cellulosic Ethanol to renewable gasoline and jet fuels with CCS of fermentation offgas and boiler vent streams CO2_jet',

    ###
    'Decarb 2b: Fischer-Tropsch SPK',
    'Decarb 2b: Fischer-Tropsch SPK with CCS of FT flue gas CO2',
    'Decarb 2b: Fischer-Tropsch SPK with CCS of all flue gases CO2',
    'Decarb 2b: Ex-Situ CFP',
    'Decarb 2b: Ex-Situ CFP with CCS of all flue gases CO2',
    ###

    # 'Gasification to Methanol',
    # 'Gasoline from upgraded bio-oil from pyrolysis'

    # 2021 SOT pathways
    ###
    '2021 SOT: Biochemical design case, Acids pathway with burn lignin',
    '2021 SOT: Biochemical design case, Acids pathway with convert lignin to BKA',
    '2021 SOT: Biochemical design case, BDO pathway with burn lignin',
    '2021 SOT: Biochemical design case, BDO pathway with convert lignin to BKA',
    '2021 SOT: High octane gasoline from lignocellulosic biomass via syngas and methanol/dimethyl ether intermediates',

    '2020 SOT: Ex-Situ CFP of lignocellulosic biomass to hydrocarbon fuels',

    ###


    # Marine pathways
    ###
    '2022, Marine biocrude via HTL from sludge with NH3 removal for 1000 MTPD sludge',
    '2022, Marine biocrude via HTL from Manure with NH3 removal for 1000 MTPD Manure',
    '2022, Partially upgraded marine fuel via HTL from sludge with NH3 removal for 1000 MTPD sludge',
    '2022, Partially upgraded marine fuel via HTL from Manure with NH3 removal for 1000 MTPD Manure',
    '2022, Fully upgraded marine fuel via HTL from sludge with NH3 removal for 1000 MTPD sludge',
    '2022, Fully upgraded marine fuel via HTL from Manure with NH3 removal for 1000 MTPD Manure',
    '2022, Marine fuel through Fast Pyrolysis of blended woody biomass',
    '2022, Marine fuel through Catalytic Fast Pyrolysis with ZSM5 of blended woody biomass',
    '2022, Marine fuel through Catalytic Fast Pyrolysis with Pt/TiO2 of blended woody biomass',
    ###

    # SAF, RD, RG via HTL of Sludge and Manure, and CFP of biomass (marine fuel comparison pathways)
    ###
    '2022 SOT: Sludge HTL to Biocrude upgraded to Hydrocarbons',

    ###

    # SAF, RD, RG via CFP of biomass (marine fuel comparison pathways)
    ###
    '2023: Catalytic Fast Pyrolysis of woody biomass to SAF, renewable diesel, and renewable gasoline, 17% oxygen content',
    '2023: Catalytic Fast Pyrolysis of woody biomass to SAF, renewable diesel, and renewable gasoline, 20% oxygen content',
    '2023: Catalytic Fast Pyrolysis of woody biomass to SAF, renewable diesel, and renewable gasoline, 22% oxygen content',
    ###


    # Biomass to Hydrogen
    ###
    'Biomass to Hydrogen',
    'Biomass to Hydrogen with CCS, design case',
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

    ###
    'Ex-Situ Fixed Bed 2018 SOT (0.5 wt% Pt/TiO2 Catalyst)',
    ###

    # 'Ex-Situ Fixed Bed 2022 Projection',
    # 'In-Situ CFP 2022 Target Case',

]
pathways_to_consider = pathways_to_consider + biopower_scenarios

df_econ = df_econ.loc[df_econ['Case/Scenario'].isin(
    pathways_to_consider)].reset_index(drop=True)

# When studying variability of unit cost on MFSP and MAC,
# following pathways are avoided because detailed LCI are not available yet
cases_to_avoid = [
    # 'Cellulosic Ethanol',
    # 'Cellulosic Ethanol with Jet Upgrading',
    # 'Fischer-Tropsch SPK',
    # 'Gasification to Methanol',
    # 'Gasoline from upgraded bio-oil from pyrolysis'
]

# Exclude cases to avoid if performing variability analysis
if consider_variability_study:
    df_econ = df_econ.loc[~df_econ['Case/Scenario']
                          .isin(cases_to_avoid)].reset_index(drop=True)

EIA_price = pd.read_csv(input_path_EIA_price + '/' +
                        f_EIA_price, index_col=None)

ef = pd.read_csv(input_path_GREET + '/' + f_GREET_efs,
                 header=3, index_col=None).drop_duplicates()

# Unit conversion class object
ob_units = model_units(input_path_units, input_path_GREET, input_path_corr)

# load correspondence files
corr_replaced_replacing_fuel = pd.read_csv(
    input_path_corr + '/' + f_corr_replaced_replacing_fuel, header=3, index_col=None)
corr_fuel_replaced_GREET_pathway = pd.read_csv(
    input_path_corr + '/' + f_corr_fuel_replaced_GREET_pathway, header=3, index_col=None)
# corr_fuel_replacing_GREET_pathway = pd.read_csv(input_path_corr + '/' + f_corr_fuel_replacing_GREET_pathway, header=3, index_col=None)
corr_GGE_GREET_fuel_replaced = pd.read_csv(
    input_path_corr + '/' + f_corr_GGE_GREET_fuel_replaced, header=3, index_col=None)
corr_GGE_GREET_fuel_replacing = pd.read_csv(
    input_path_corr + '/' + f_corr_GGE_GREET_fuel_replacing, header=3, index_col=None)

corr_itemized_LCA = pd.read_csv(
    input_path_corr + '/' + f_corr_itemized_LCI, dtype={8: 'str'}, header=0, index_col=0)
corr_itemized_LCA.drop_duplicates(inplace=True)

corr_replaced_mfsp = pd.read_csv(
    input_path_corr + '/' + f_corr_replaced_EIA_mfsp, header=3, index_col=None)

if consider_variability_study:
    corr_params_variability = pd.read_excel(input_path_model + '/' + f_model,
                                            sheet_name=sheet_param_variability,
                                            header=3, index_col=None,
                                            usecols="A:G")

# %%

# Step: Create Cost Item table

df_econ.loc[df_econ['Stream_Flow'].isna(), 'Stream_Flow'] = ''

# Subset cost items to use for itemized MFSP calculation
cost_items = df_econ.loc[df_econ['Parameter_B'].isin([
    'Conversion: Input Supply Chains',
    'Coproduct Credits',
    'Avoided Ems Credits',

    'Fuel Use',  # Co-produced fuels marked as Fuel Use are used to calculate displacement credit in Hybrid approach

    'Fixed Costs',
    'Capital Depreciation',
    'Average Income Tax',
    'Average Return on Investment',

    'Cost by process steps']), :].copy()

# Check if cost_items have duplicates
tmpdf = cost_items.duplicated()
if sum(tmpdf):
    print("Warning: Following duplicate rows in cost_items table. Investigate the cause of duplication ..")
    tmpdf = cost_items[tmpdf]
    print(tmpdf)

    # Validated duplicates
    # Natural gas for '2021 SOT BDO and Acids pathways'

# %%

# Step: Create Biofuel Yield table

# Separate biofuel yield flows
biofuel_yield = df_econ.loc[df_econ['Parameter_B'] == 'Fuel Use',
                            ['Case/Scenario', 'Stream_LCA', 'Total Flow: Unit (numerator)',
                             'Total Flow: Unit (denominator)', 'Total Flow', 'Energy_alloc_primary_fuel']].reset_index(drop=True).copy()
biofuel_yield.rename(columns={'Stream_LCA': 'Biofuel Stream_LCA',
                              'Total Flow: Unit (numerator)': 'Biofuel Flow: Unit (numerator)',
                              'Total Flow: Unit (denominator)': 'Biofuel Flow: Unit (denominator)',
                              'Total Flow': 'Biofuel Flow'}, inplace=True)

biofuel_yield_primary_out = biofuel_yield.loc[biofuel_yield['Energy_alloc_primary_fuel'].isin(['Y']), : ].reset_index(drop=True)

# If energy allocation is to be implemented, filter by primary fuel flag
# Pathway allocation is helpful for assessing CI of the production pathway

if allocation_type == 'Pathway':  # to be checked, may not be required as it is similar to energy allocation
    # For co-produced fuels, summarize the flow data to net hydrocarbon flow
    biofuel_yield2 = biofuel_yield.groupby(['Case/Scenario', 'Biofuel Flow: Unit (numerator)',
                                            'Biofuel Flow: Unit (denominator)']).agg({'Biofuel Flow': 'sum'}).reset_index()
    biofuel_yield2['biofuel_yield_energy_alloc'] = 1

elif allocation_type == 'Energy':  # to be checked
    # For co-produced fuels, summarize the flow data to net hydrocarbon flow
    biofuel_yield2 = biofuel_yield.groupby(['Case/Scenario', 'Biofuel Flow: Unit (numerator)',
                                            'Biofuel Flow: Unit (denominator)']).agg({'Biofuel Flow': 'sum'}).reset_index()
    biofuel_yield2['biofuel_yield_energy_alloc'] = 1

elif allocation_type == 'Hybrid':
    # Calculate energy allocation fraction per 'Fuel Use' product
    biofuel_yield['biofuel_yield_energy_alloc'] = biofuel_yield['Biofuel Flow']/biofuel_yield.groupby(['Case/Scenario', 'Biofuel Flow: Unit (numerator)',
                                                                                                       'Biofuel Flow: Unit (denominator)'])['Biofuel Flow'].transform('sum')

    # filter and select primary fuel
    biofuel_yield2 = biofuel_yield.loc[biofuel_yield['Energy_alloc_primary_fuel'].isin([
                                                                                       'Y']), :]
    # If two 'Fuel Use' are identified as primary product, they are aggregrated at this stage
    biofuel_yield2 = biofuel_yield2.groupby(['Case/Scenario', 'Biofuel Flow: Unit (numerator)',
                                          'Biofuel Flow: Unit (denominator)']).agg({'Biofuel Flow': 'sum',
                                                                                       'biofuel_yield_energy_alloc': 'sum'}).reset_index()

else:
    print("Warning: unrecognized energy allocation type, please check parameter declaration.")

# %%

# Step: Merge biofuel flows with Cost Item table
# For primary fuel products, their biofuel flow are mapped to singular flows rather than aggregrated 'fuel use' energy products
# This separation helps to add T&D costs for primary product
tmpdf = cost_items.loc[cost_items['Energy_alloc_primary_fuel'].isin(['Y']), : ]
cost_items = cost_items.loc[~ (cost_items['Energy_alloc_primary_fuel'].isin(['Y']) ), : ]

cost_items = pd.merge(cost_items, biofuel_yield2, how='left', on='Case/Scenario').reset_index(drop=True)
tmpdf = pd.merge(tmpdf, biofuel_yield_primary_out[['Case/Scenario', 'Biofuel Stream_LCA', 'Biofuel Flow: Unit (numerator)',
       'Biofuel Flow: Unit (denominator)', 'Biofuel Flow']], how='left', 
                  left_on=['Case/Scenario', 'Stream_LCA'],
                  right_on = ['Case/Scenario', 'Biofuel Stream_LCA']).reset_index(drop=True)
cost_items =  pd.concat([cost_items, tmpdf], ignore_index=True).copy()
cost_items.sort_values(by=['Case/Scenario', 'Parameter_B']).reset_index(drop=True, inplace=True)
#cost_items.drop(columns=['Energy_alloc_primary_fuel_x', 'Biofuel Stream_LCA', 'Energy_alloc_primary_fuel_y'], inplace=True)

if allocation_type == 'Energy':
    pass

# based on energy allocation choice, all components are allocated per 'Fuel Use' product
elif allocation_type == 'Hybrid':
    for colm in ['Flow', 'Total Cost', 'Total Flow']:
        cost_items.loc[[isinstance(x, numbers.Number) for x in cost_items[colm]], colm] =\
            cost_items.loc[[isinstance(x, numbers.Number) for x in cost_items[colm]], colm].multiply(
            cost_items.loc[[isinstance(x, numbers.Number) for x in cost_items[colm]], 'biofuel_yield_energy_alloc'], axis='index')

# %%
# Step: calculate cost per variability of parameters

# drop blanks and zeros
cost_items = cost_items.query('`Total Cost` not in ["-", None, 0]').copy()


if consider_variability_study & (consider_which_variabilities == 'Cost_Item'):

    # unit check
    check_units = (cost_items['Flow: Unit (numerator)'] != cost_items['Cost: Unit (denominator)']) |\
        (cost_items['Flow: Unit (denominator)']
         != cost_items['Operating Time: Unit'])
    cost_items = cost_items.loc[~check_units]
    check_units = cost_items.loc[check_units, :]
    if check_units.shape[0] > 0:
        print("Warning: The following cost items need attention as the units are not harmonized ..")
        print(check_units)

    var_params = corr_params_variability.loc[corr_params_variability['col_param'].isin(['Cost Item']), :].reset_index(drop=True)

    var_params_tbl = variability_table(var_params).reset_index(drop=True)
    var_params_tbl['variability_id'] = var_params_tbl.index

    cost_items_temp = cost_items.copy()

    cost_items_list = []    
    for r in range(0, var_params_tbl.shape[0]):
        cost_items_temp.loc[
            cost_items_temp[var_params_tbl.loc[r, 'col_param']].isin(
                [var_params_tbl.loc[r, 'param_name']]),
            var_params_tbl.loc[r, 'col_val']] = var_params_tbl.loc[r, 'param_value']
        cost_items_temp['variability_id'] = var_params_tbl.loc[r,'variability_id']
        cost_items_list.append(cost_items_temp)
    cost_items = pd.concat(cost_items_list, ignore_index=True)

    cost_items = cost_items.merge(
        var_params_tbl, how='left', on='variability_id').reset_index(drop=True)


if consider_scale_up_study:
    
    # Map biomass types based on feedstock
    map_bm_types = {
        'Coal to Power Plants': '', 
        'Poplar': 'Woody', 
        'Blended woody biomass': 'Woody',
        'Forest Residue': 'Woody', 
        'Switchgrass': 'Herbaceous',
        'Sludge to Biorefinery for HTL': 'Wastewater sludge',
        'Logging Residues for CFP, 2020 SOT': 'Woody',
        'Clean Pine for CFP, 2020 SOT': 'Woody',
        'Manure to Biorefinery for HTL': 'Manure'
    }

    # Read data on biomass availability
    bm = pd.read_excel(input_path_BT16 + '/' + f_BT16_availability,
                    sheet_name=sheet_BT16_availability,
                    header=17, index_col=None,
                    usecols="C:K")
    bm = bm.loc[~bm['Aggregated'].isin(['Total']), : ]
    bm = bm.melt(['Aggregated'])
    bm.rename(columns={'Aggregated': 'bm_types', 'variable': 'bm_cost', 'value': 'qty_dry_bm'}, inplace=True)

    # Add unit cost and year columns
    bm['bm_cost: Unit (numerator)'] = 'USD'
    bm['bm_cost: Unit (denominator)'] = 'dt'
    bm['bm_cost: USD year'] = BT16_cost_year
    bm['qty_dry_bm: Unit'] = 'MM dt'

    # Add cost of T&D and feedstock loss penalty
    bm['bm_cost'] += 40  # Reactor throat price
    bm['qty_dry_bm'] *= 0.7  # 30% penalty for feedstock loss

    # Get unique biomass costs
    unique_bm_costs = bm['bm_cost'].unique()

    # Initialize cost_items DataFrame
    tmp_cost_items = cost_items.copy()
    cost_items = pd.DataFrame(columns=tmp_cost_items.columns.to_list() + ['bm_cost_id'])

    # Map biomass types to feedstocks
    feedstock_mask = tmp_cost_items['Parameter_A'] == 'Feedstock'
    tmp_cost_items.loc[feedstock_mask, 'bm_types'] = tmp_cost_items.loc[feedstock_mask, 'Stream_LCA'].map(map_bm_types)

    # Expand cost_items for every feedstock cost
    for c in unique_bm_costs:
        # Apply feedstock cost where valid
        valid_feedstocks = feedstock_mask & tmp_cost_items['bm_types'].notna() & (tmp_cost_items['bm_types'] != '')
        
        tmp_cost_items.loc[valid_feedstocks, 'Unit Cost'] = c
        tmp_cost_items.loc[valid_feedstocks, 'Cost Year'] = BT16_cost_year
        
        # Harmonize unit (dt to dry lb)
        tmp_cost_items.loc[valid_feedstocks, 'Unit Cost'] /= 2000
        
        # Recalculate total cost
        total_cost_mask = valid_feedstocks & (tmp_cost_items['Flow'] != '-') & (tmp_cost_items['Operating Time'] != '-') & (tmp_cost_items['Unit Cost'] != '-')
        tmp_cost_items.loc[total_cost_mask, 'Total Cost'] = (
            tmp_cost_items.loc[total_cost_mask, 'Flow'] *
            tmp_cost_items.loc[total_cost_mask, 'Operating Time'] *
            tmp_cost_items.loc[total_cost_mask, 'Unit Cost']
        )
        
        # Add identifier column
        tmp_cost_items['bm_cost_id'] = c

        # Accumulate results in a list to avoid multiple concat calls
        if cost_items.empty:
            cost_items = tmp_cost_items.copy()
        else:
            cost_items = pd.concat([cost_items, tmp_cost_items], ignore_index=True, sort=False)

    # Reset index after concat and clean up
    cost_items.reset_index(drop=True, inplace=True)


if consider_variability_study and (consider_which_variabilities == 'Cost_Item'):
    # Calculate itemized MFSP

    # Create a mask for valid rows where 'Flow', 'Operating Time', and 'Unit Cost' are not '-'
    valid_mask = (cost_items['Flow'] != '-') & (cost_items['Operating Time'] != '-') & (cost_items['Unit Cost'] != '-')

    # Convert 'Flow', 'Operating Time', and 'Unit Cost' to numeric where valid
    cost_items.loc[valid_mask, 'Flow'] = pd.to_numeric(cost_items.loc[valid_mask, 'Flow'])
    cost_items.loc[valid_mask, 'Operating Time'] = pd.to_numeric(cost_items.loc[valid_mask, 'Operating Time'])
    cost_items.loc[valid_mask, 'Unit Cost'] = pd.to_numeric(cost_items.loc[valid_mask, 'Unit Cost'])

    # Recalculate total cost for valid rows
    cost_items.loc[valid_mask, 'Total Cost'] = (
        cost_items.loc[valid_mask, 'Flow'] *
        cost_items.loc[valid_mask, 'Operating Time'] *
        cost_items.loc[valid_mask, 'Unit Cost']
    )


print("Correcting inflation to the year of study..")

# Correct for inflation to the year of study

#cost_items['Cost Year'] = cost_items['Cost Year'].astype(int)
# Inflation row-wise
#cost_items['Adjusted Total Cost'] = cost_items.apply(
#    lambda row: cpi.inflate(row['Total Cost'], row['Cost Year'], to=cost_year),
#    axis=1
#)

# revising inflation adjustment code to improve performance
cpi_to = cpi.get(cost_year)
cost_items['Cost Year'] = cost_items['Cost Year'].astype(int)
start_years = cost_items['Cost Year'].unique().astype(int)
cpi_start_dict = {year: cpi.get(year) for year in start_years}
cost_items['cpi_start'] = cost_items['Cost Year'].map(cpi_start_dict)
cost_items['Adjusted Total Cost'] = cost_items['Total Cost'] * (cpi_to / cost_items['cpi_start'])

# Record adjusted year
cost_items['Adjusted Cost Year'] = cost_year

# Set or expand LCI based on production year
if len(production_year) == 1:
    cost_items['Production Year'] = production_year[0]
else:
    # Create a list of all years in the range
    years = list(range(production_year[0], production_year[1] + 1))
    
    # Replicate the cost_items DataFrame for each year and assign the 'Production Year'
    cost_items = pd.DataFrame(np.repeat(cost_items.values, len(years), axis=0), columns=cost_items.columns)
    cost_items['Production Year'] = np.tile(years, len(cost_items) // len(years))
    
cost_items.reset_index(drop=True, inplace=True)

print ("Calculating MFSP ..")

# Calculate itemized MFSP
cost_items['Itemized MFSP'] = cost_items['Adjusted Total Cost'].astype(float) / cost_items['Biofuel Flow'].astype(float)
cost_items['Itemized MFSP: Unit (numerator)'] = cost_items['Total Cost: Unit (numerator)']
cost_items['Itemized MFSP: Unit (denominator)'] = cost_items['Biofuel Flow: Unit (numerator)']

# Harmonize energy units, convert kWh to MJ
tmp_cost_items = cost_items[cost_items['Itemized MFSP: Unit (denominator)'] == 'kWh'].copy()
cost_items = cost_items[cost_items['Itemized MFSP: Unit (denominator)'] != 'kWh']

converted = ob_units.unit_convert_df(
    tmp_cost_items[['Itemized MFSP: Unit (denominator)', 'Itemized MFSP']],
    Unit='Itemized MFSP: Unit (denominator)',
    Value='Itemized MFSP',
    if_unit_numerator=False,
    if_given_unit=True,
    given_unit='MJ'
)
tmp_cost_items['Itemized MFSP: Unit (denominator)'] = converted['Itemized MFSP: Unit (denominator)']
tmp_cost_items['Itemized MFSP'] = converted['Itemized MFSP']
cost_items = pd.concat([cost_items, tmp_cost_items], ignore_index=True)

# Identify non-harmonized units
ignored_cost_items = cost_items[(cost_items['Total Flow: Unit (numerator)'] != cost_items['Cost: Unit (denominator)']) &
                                 ~(cost_items['Parameter_A'].isin(['Fixed Costs', 'Capital Depreciation', 'Average Income Tax', 'Average Return on Investment']))]

if ignored_cost_items.shape[0] > 0:
    print("Warning: The following cost items need attention as the units are not harmonized ..")
    print(ignored_cost_items)

# Concatenate the energy unit-adjusted data and original data
cost_items = pd.concat([tmp_cost_items, cost_items]).reset_index(drop=True)

# For co-products, consider their cost as a credit to the MFSP
cost_items.loc[cost_items['Parameter_B'] == 'Coproduct Credits', 'Itemized MFSP'] *= -1
if allocation_type == 'Energy':
    # In Energy allocation, no cost credit is considered for produced fuel (per MJ of fuels produced)
    # Remove mapping of Fuel Use to T&D costs (already removed in database)
    pass

elif allocation_type == 'Hybrid':
    # For 'Hybrid' allocation, fuel use items that are not primary fuel get a displacement credit
    cost_items.loc[(cost_items['Parameter_B'] == 'Fuel Use') &
                   (cost_items['Energy_alloc_primary_fuel'] != 'Y'), 'Itemized MFSP'] *= -1


# %%
# Step: Calculate aggregrated Marginal Fuel Selling Price (MFSP)

MFSP_agg = cost_items.copy()

if not consider_coproduct_cost_credit:
    MFSP_agg = MFSP_agg.loc[~MFSP_agg['Parameter_B'].isin(
        ['Coproduct Credits']), :]

if consider_variability_study and (consider_which_variabilities == 'Cost_Item'):
    # Filter and group for variability study
    MFSP_agg = MFSP_agg[['Case/Scenario',
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
    MFSP_agg = MFSP_agg.groupby([
        'Case/Scenario', 'Production Year', 'Itemized MFSP: Unit (numerator)', 
        'Itemized MFSP: Unit (denominator)', 'Adjusted Cost Year', 'variability_id', 
        'col_param', 'col_val', 'param_name', 'param_min', 'param_max', 'param_dist', 
        'dist_option', 'param_value'
    ])['Itemized MFSP'].sum().reset_index()

elif consider_scale_up_study:
    # Filter and group for scale-up study
    MFSP_agg = MFSP_agg[['Case/Scenario',
                         'Production Year',
                         'Itemized MFSP: Unit (numerator)',
                         'Itemized MFSP: Unit (denominator)',
                         'Adjusted Cost Year',
                         'Itemized MFSP',
                         'bm_cost_id']]
    MFSP_agg = MFSP_agg[MFSP_agg['Itemized MFSP'].notna()]
    MFSP_agg = MFSP_agg.groupby([
        'Case/Scenario', 'Production Year', 'Itemized MFSP: Unit (numerator)', 
        'Itemized MFSP: Unit (denominator)', 'Adjusted Cost Year', 'bm_cost_id'
    ])['Itemized MFSP'].sum().reset_index()

else:
    # Filter and group for general case (no study considered)
    MFSP_agg = MFSP_agg[['Case/Scenario',
                         'Production Year',
                         'Itemized MFSP: Unit (numerator)',
                         'Itemized MFSP: Unit (denominator)',
                         'Adjusted Cost Year',
                         'Itemized MFSP']]
    MFSP_agg = MFSP_agg[MFSP_agg['Itemized MFSP'].notna()]
    MFSP_agg = MFSP_agg.groupby([
        'Case/Scenario', 'Production Year', 'Itemized MFSP: Unit (numerator)', 
        'Itemized MFSP: Unit (denominator)', 'Adjusted Cost Year'
    ])['Itemized MFSP'].sum().reset_index()

# Rename columns to reflect MFSP replacing fuel
MFSP_agg.rename(columns={
    'Itemized MFSP': 'MFSP replacing fuel',
    'Itemized MFSP: Unit (numerator)': 'MFSP replacing fuel: Unit (numerator)',
    'Itemized MFSP: Unit (denominator)': 'MFSP replacing fuel: Unit (denominator)'
}, inplace=True)

# Merge with biofuel yield data to get Fuel Use column back
MFSP_agg = pd.merge(biofuel_yield[['Case/Scenario', 'Biofuel Stream_LCA', 'Energy_alloc_primary_fuel']].drop_duplicates(),
                    MFSP_agg, how='left', on='Case/Scenario').reset_index(drop=True)

# Save interim data tables
if save_interim_files:
    cost_items.to_csv(output_path_prefix + '/' + f_out_itemized_mfsp)
    MFSP_agg.to_csv(output_path_prefix + '/' + f_out_agg_mfsp)

# %%

# Step: Expand LCIs based on TEA production year

print ('Expand LCI and LCA calculation ..')

# Filter relevant LCA items
LCA_parameters = [
    'Conversion: Input Supply Chains',
    'Avoided Ems Credits',
    'Conversion: Combustion Ems, Fossil',
    'Conversion: Combustion Ems, Biogenic',
    'Conversion: Non-Combustion Ems, Fossil',
    'Conversion: Non-Combustion Ems, Biogenic',
    'Coproduct Credits',
    'Fuel Use',
    'CCS Stream, Fossil',
    'CCS Stream, Biogenic'
]

LCA_columns = [
    'Case/Scenario',
    'Parameter_A',
    'Parameter_B',
    'Stream_Flow',
    'Stream_LCA',
    'Energy_alloc_primary_fuel',
    'Flow: Unit (numerator)',
    'Flow: Unit (denominator)',
    'Flow',
    'Operating Time: Unit',
    'Operating Time',
    'Operating Time (%)',
    'Total Flow: Unit (numerator)',
    'Total Flow: Unit (denominator)',
    'Total Flow'
]

# Apply filter and column selection
LCA_items = df_econ.loc[df_econ['Parameter_B'].isin(LCA_parameters), LCA_columns].reset_index(drop=True)


if len(production_year) == 1:
    LCA_items['Production Year'] = production_year[0]
else:
    LCA_items = pd.concat(
        [LCA_items.assign(**{'Production Year': yr}) for yr in range(production_year[0], production_year[1]+1)],
        ignore_index=True
    )

# format LCI
tempdf, corr_itemized_LCA = fmt_GREET_LCI(corr_itemized_LCA)

# %%
# Update carbon intensities as per scope of study

# implementing carbon intensities of decarbonized electric grid

if decarb_electric_grid:
    decarb_elec_CI = pd.read_excel(input_path_decarb_model + '/' + f_Decarb_Model,
                                   sheet_name='EPS - CI', header=3)

    decarb_elec_CI = decarb_elec_CI.loc[(decarb_elec_CI['Case'] == 'Mitigation') &
                                        (decarb_elec_CI['Mitigation Case'] == 'NREL Electric Power Decarb') &
                                        (decarb_elec_CI['LCIA Method'] == 'AR4') &
                                        (decarb_elec_CI['timeframe_years'] == 100), :]

    decarb_elec_CI = decarb_elec_CI.groupby(['Year', 'Emissions Unit', 'Energy Unit'
                                             ]).agg({'LCIA_estimate': 'sum'}).reset_index()

    decarb_elec_CI = decarb_elec_CI.loc[decarb_elec_CI['Year'].isin(
        LCA_items['Production Year'].unique())].copy()

    # if artificial scaling of CI is enabled for sensitivity analysis
    if decarb_grid_scenario1:
        tmpdf = pd.DataFrame({'Year': np.linspace(max(decarb_elec_CI['Year']), min(decarb_elec_CI['Year']), max(decarb_elec_CI['Year']) - min(decarb_elec_CI['Year'])+1),
                               'LCA_value_replace': np.linspace(decarb_grid_scenario1_values[0], decarb_grid_scenario1_values[1], max(decarb_elec_CI['Year']) - min(decarb_elec_CI['Year'])+1)})
        decarb_elec_CI = pd.merge(decarb_elec_CI, tmpdf, how='left', on=[
                                  'Year']).reset_index(drop=True)
        decarb_elec_CI.drop(columns=['LCIA_estimate'], inplace=True)
        decarb_elec_CI.rename(
            columns={'LCA_value_replace': 'LCIA_estimate'}, inplace=True)

    # Creating mapping columns
    decarb_elec_CI[['Parameter_B', 'Stream_Flow', 'Stream_LCA']] = [
        'Conversion: Input Supply Chains', 'Electricity', 'Stationary Use: U.S. Mix']
    tmpdf = decarb_elec_CI.copy()
    tmpdf['Parameter_B'] = 'Coproduct Credits'
    decarb_elec_CI = pd.concat([decarb_elec_CI, tmpdf], ignore_index=True)

    # replace CIs in LCA data frame
    decarb_elec_CI.rename(columns={
        'Emissions Unit': 'LCA: Unit (numerator)',
        'Energy Unit': 'LCA: Unit (denominator)',
        'LCIA_estimate': 'LCA_value',
    }, inplace=True)
    decarb_elec_CI['LCA_metric'] = 'CO2e'
    decarb_elec_CI[list((Counter(corr_itemized_LCA.columns) -
                        Counter(decarb_elec_CI.columns)).elements())] = '-'
    decarb_elec_CI = decarb_elec_CI.loc[decarb_elec_CI['Year'].isin(
        corr_itemized_LCA['Year'].unique()), :]

    # convert LCA unit of flow to model standard unit
    decarb_elec_CI.loc[:, ['LCA: Unit (denominator)', 'LCA_value']] = \
        ob_units.unit_convert_df(decarb_elec_CI.loc[:, ['LCA: Unit (denominator)', 'LCA_value']],
                                 Unit='LCA: Unit (denominator)', Value='LCA_value',
                                 if_unit_numerator=False, if_given_category=False)

    corr_itemized_LCA = corr_itemized_LCA.loc[~ ((corr_itemized_LCA['Stream_Flow'] == 'Electricity') &
                                                 (corr_itemized_LCA['Stream_LCA'] == 'Stationary Use: U.S. Mix')), :]
    corr_itemized_LCA = pd.concat(
        [corr_itemized_LCA, decarb_elec_CI], ignore_index=True)


# %%
# Merge itemized LCAs to LCIs
LCA_items = pd.merge(LCA_items, corr_itemized_LCA, how='left',
                     left_on=['Parameter_B', 'Stream_Flow',
                              'Stream_LCA', 'Production Year'],
                     right_on=['Parameter_B', 'Stream_Flow', 'Stream_LCA', 'Year']).reset_index(drop=True)

# Variability analysis for LCA parameters
if consider_variability_study and (consider_which_variabilities == 'Stream_LCA'):

    var_params = corr_params_variability.loc[
        corr_params_variability['col_param'] == 'Stream_LCA'
    ].reset_index(drop=True)

    var_params_tbl = variability_table(var_params).reset_index(drop=True)
    var_params_tbl['variability_id'] = var_params_tbl.index

    # Prepare a list to collect all modified versions
    modified_LCA_items = []

    for _, row in var_params_tbl.iterrows():
        temp = LCA_items.copy()
        mask = temp[row['col_param']].isin([row['param_name']])
        temp.loc[mask, row['col_val']] = row['param_value']
        temp['variability_id'] = row['variability_id']
        modified_LCA_items.append(temp)

    # Only one concat after the loop (MUCH faster)
    LCA_items = pd.concat(modified_LCA_items, ignore_index=True)

    LCA_items = LCA_items.merge(
        var_params_tbl, how='left', on='variability_id'
    ).reset_index(drop=True)


# harmonize units
# converting material flow units to model standard units
# Replace '-' with 0 and convert to numeric in one go (vectorized)
LCA_items['Total Flow'] = pd.to_numeric(
    LCA_items['Total Flow'].replace('-', 0)
)

# Fill missing values
LCA_items['Total Flow: Unit (numerator)'] = LCA_items['Total Flow: Unit (numerator)'].fillna('-')

# Create a mask once
mask = LCA_items['Total Flow: Unit (numerator)'] != '-'

# Only call unit_convert_df once with the filtered rows
if mask.any():
    converted = ob_units.unit_convert_df(
        LCA_items.loc[mask, ['Total Flow: Unit (numerator)', 'Total Flow']],
        Unit='Total Flow: Unit (numerator)', 
        Value='Total Flow',
        if_unit_numerator=True,
        if_given_category=False
    )
    LCA_items.loc[mask, ['Total Flow: Unit (numerator)', 'Total Flow']] = converted

# Check for non-harmonized units
non_harmonized_mask = LCA_items['Total Flow: Unit (numerator)'] != LCA_items['LCA: Unit (denominator)']
if non_harmonized_mask.any():
    ignored_LCA_items = LCA_items.loc[non_harmonized_mask]
    print("Warning: The following LCA items need attention as the units are not harmonized ..")
    print(ignored_LCA_items)

# Keep only harmonized items
LCA_items = LCA_items.loc[~non_harmonized_mask]


# %%
# Step: Itemized LCA and CCS implementation

# Calculate itemized LCA metric per year
LCA_items['Total LCA'] = LCA_items['LCA_value'] * LCA_items['Total Flow']
LCA_items['Total LCA: Unit (numerator)'] = LCA_items['LCA: Unit (numerator)']
LCA_items['Total LCA: Unit (denominator)'] = LCA_items['Total Flow: Unit (denominator)']

# If co-product, credit LCA by displacement
coproduct_mask = LCA_items['Parameter_B'] == 'Coproduct Credits'
LCA_items.loc[coproduct_mask, 'Total LCA'] *= -1

# Merge biofuel yield data by 'Case/Scenario'
LCA_items = LCA_items.merge(biofuel_yield2, how='left', on='Case/Scenario').reset_index(drop=True)

# Allocation step
if allocation_type == 'Energy':
    pass  # Nothing to do

elif allocation_type == 'Hybrid':
    # Only for numeric columns
    cols_to_allocate = ['Flow', 'Total Flow', 'Total LCA']
    
    # Prepare mask once for all columns (assuming consistent types)
    numeric_mask = LCA_items[cols_to_allocate[0]].apply(lambda x: isinstance(x, numbers.Number))

    for col in cols_to_allocate:
        LCA_items.loc[numeric_mask, col] *= LCA_items.loc[numeric_mask, 'biofuel_yield_energy_alloc']

# %%

# Step: Checks for CCS, Fossil net calculation

# If CCS_fossil_CO2 > combustion_fossil_CO2 show warning to users, to investigate source of the residual CO2 for CCS.
# [combustion_fossil_CO2 - CCS_CO2] net emission of combustion_fossil_CO2 is accounted.
# CCS_biogenic_CO2 is credited
if harmonize_CCS_fossil:

    # Select Case/Scenario with CCS flow
    CCS_cases = LCA_items.query("Parameter_B == 'CCS Stream, Fossil' and Stream_Flow == 'Carbon Dioxide'")['Case/Scenario'].drop_duplicates()

    # Select rows with selected cases and Parameter_B in ['CCS Stream', 'Combustion, Fossil']
    tmp_LCA_items = LCA_items.query("`Case/Scenario` in @CCS_cases and Parameter_B in ['CCS Stream, Fossil', 'Conversion: Combustion Ems, Fossil']").copy()

    # Remove selected rows from LCA_items
    LCA_items = LCA_items.drop(tmp_LCA_items.index).reset_index(drop=True)

    tmp_CCS = tmp_LCA_items.query("Parameter_B == 'CCS Stream, Fossil'").copy()
    tmp_combust = tmp_LCA_items.query("Parameter_B == 'Conversion: Combustion Ems, Fossil'").copy()

    # Merge combustion emissions if-any to CCS flows
    tmp_merge = pd.merge(
    tmp_CCS[['Case/Scenario', 'Production Year', 'Year', 'Total LCA', 'Total LCA: Unit (numerator)', 'Total LCA: Unit (denominator)']],
    tmp_combust[['Case/Scenario', 'Production Year', 'Year', 'Total LCA', 'Total LCA: Unit (numerator)', 'Total LCA: Unit (denominator)']],
    on=['Case/Scenario', 'Production Year', 'Year'],
    suffixes=('_CCS, fossil', '_combustion, fossil'),
    how='left'
    )
    
    # Check Unit Consistency
    unit_mismatch = (
        (tmp_merge['Total LCA: Unit (numerator)_CCS, fossil'] != tmp_merge['Total LCA: Unit (numerator)_combustion, fossil']) |
        (tmp_merge['Total LCA: Unit (denominator)_CCS, fossil'] != tmp_merge['Total LCA: Unit (denominator)_combustion, fossil'])
    )

    if unit_mismatch.any():
        print("\n⚠️ Warning: Unit mismatch detected between CCS and Combustion LCA values:")
        print(tmp_merge.loc[unit_mismatch].to_string(index=False))

    # Keep only matching units
    tmp_merge = tmp_merge.loc[~unit_mismatch].reset_index(drop=True)


    # Calculate net CCS stream, CCS_net: 'Combustion, Fossil' + 'CCS Stream credit'
    tmp_merge['Total LCA_combustion, fossil_net'] =\
        tmp_merge['Total LCA_combustion, fossil'] + \
        tmp_merge['Total LCA_CCS, fossil']

    # zero CCS, fossil
    tmp_merge['Total LCA_CCS, fossil_net'] = 0

    # Warning for Negative Net Emissions
    negatives = tmp_merge[tmp_merge['Total LCA_combustion, fossil_net'] < 0]

    if not negatives.empty:
        print("\n⚠️ Warning: Net combustion emissions are negative in the following rows:")
        print(negatives.to_string(index=False))

    tmp_combust = tmp_combust.merge(
    tmp_merge[['Case/Scenario', 'Production Year', 'Year', 'Total LCA_combustion, fossil_net']],
    how='left',
    on=['Case/Scenario', 'Production Year', 'Year']
    )
    tmp_combust['Total LCA'] = tmp_combust['Total LCA_combustion, fossil_net']

    # Map net CCS values
    tmp_CCS = tmp_CCS.merge(
        tmp_merge[['Case/Scenario', 'Production Year', 'Year', 'Total LCA_CCS, fossil_net']],
        how='left',
        on=['Case/Scenario', 'Production Year', 'Year']
    )
    tmp_CCS['Total LCA'] = tmp_CCS['Total LCA_CCS, fossil_net']

    # Combine and rebuild tmp_LCA_items
    tmp_LCA_items_updated = pd.concat([tmp_CCS, tmp_combust], ignore_index=True)

    # Keep only necessary columns
    columns_to_keep = [
        'Case/Scenario', 'Parameter_A', 'Parameter_B', 'Stream_Flow',
        'Stream_LCA', 'Flow: Unit (numerator)', 'Flow: Unit (denominator)', 'Flow',
        'Operating Time: Unit', 'Operating Time', 'Operating Time (%)',
        'Total Flow: Unit (numerator)', 'Total Flow: Unit (denominator)', 'Total Flow',
        'Production Year', 'GREET1 sheet', 'Coproduct allocation method',
        'GREET classification of coproduct', 'LCA: Unit (numerator)', 'LCA: Unit (denominator)',
        'Year', 'LCA_value', 'LCA_metric', 'Total LCA',
        'Total LCA: Unit (numerator)', 'Total LCA: Unit (denominator)',
        'Biofuel Flow: Unit (numerator)', 'Biofuel Flow: Unit (denominator)', 'Biofuel Flow'
    ]

    tmp_LCA_items_updated = tmp_LCA_items_updated[columns_to_keep].copy()

    # Final Concatenation
    LCA_items = pd.concat([LCA_items, tmp_LCA_items_updated], ignore_index=True)

# %%

# Step: LCA Metric calculation, itemized and aggregrated

# Calculate LCA metric per unit biofuel yield
LCA_items['Total LCA'] /= LCA_items['Biofuel Flow']
LCA_items['Total LCA: Unit (denominator)'] = LCA_items['Biofuel Flow: Unit (numerator)']

# Harmonize units: convert 'kWh' -> 'MJ'
kWh_mask = LCA_items['Total LCA: Unit (denominator)'] == 'kWh'
if kWh_mask.any():
    tmp_LCA_items = LCA_items.loc[kWh_mask].copy()
    LCA_items = LCA_items.loc[~kWh_mask].copy()

    tmp_LCA_items[['Total LCA: Unit (denominator)', 'Total LCA']] = ob_units.unit_convert_df(
        tmp_LCA_items[['Total LCA: Unit (denominator)', 'Total LCA']],
        Unit='Total LCA: Unit (denominator)',
        Value='Total LCA',
        if_unit_numerator=False,
        if_given_unit=True,
        given_unit='MJ'
    )

    # Re-combine
    LCA_items = pd.concat([LCA_items, tmp_LCA_items], ignore_index=True)

# Prepare for aggregation
LCA_items_agg = LCA_items.copy()

if not consider_coproduct_env_credit:
    LCA_items_agg = LCA_items_agg[LCA_items_agg['Parameter_B'] != 'Coproduct Credits']

# Aggregate Total LCA
if consider_variability_study and consider_which_variabilities == 'Stream_LCA':
    group_cols = [
        'Case/Scenario', 'LCA_metric', 'Total LCA: Unit (numerator)', 'Total LCA: Unit (denominator)',
        'Production Year', 'variability_id', 'col_param', 'col_val',
        'param_name', 'param_min', 'param_max', 'param_dist', 'dist_option', 'param_value'
    ]
else:
    group_cols = [
        'Case/Scenario', 'LCA_metric', 'Total LCA: Unit (numerator)', 'Total LCA: Unit (denominator)',
        'Production Year'
    ]

LCA_items_agg = LCA_items_agg.groupby(group_cols, as_index=False)['Total LCA'].sum()

# Save interim outputs
if save_interim_files:
    LCA_items.to_csv(f"{output_path_prefix}/{f_out_itemized_LCA}", index=False)
    LCA_items_agg.to_csv(f"{output_path_prefix}/{f_out_agg_LCA}", index=False)


# %%

# Step: Merge correspondence tables and GREET emission factors

print ('Calculate MAC ..')

# Merge aggregrated LCA metric to MFSP tables
# Merge MFSP and LCA aggregated data
MAC_df = pd.merge(
    MFSP_agg.loc[MFSP_agg['Energy_alloc_primary_fuel'] == 'Y'],
    LCA_items_agg,
    on=['Case/Scenario', 'Production Year'],
    how='inner'
).reset_index(drop=True)

# Map replaced fuels to replacing fuels
MAC_df = MAC_df.merge(
    corr_replaced_replacing_fuel,
    on=['Case/Scenario', 'Biofuel Stream_LCA'],
    how='left'
)

# Map replaced fuels to GREET pathways
MAC_df = MAC_df.merge(
    corr_fuel_replaced_GREET_pathway,
    on='Replaced Fuel',
    how='left'
).rename(
    columns={'GREET Pathway': 'GREET Pathway for replaced fuel'}
)

# Map replaced fuels to their CIs
MAC_df = MAC_df.merge(
    corr_itemized_LCA,
    left_on=['Parameter_B', 'Stream_Flow', 'Stream_LCA', 'Production Year'],
    right_on=['Parameter_B', 'Stream_Flow', 'Stream_LCA', 'Year'],
    how='left'
)

# Clean up and rename columns
MAC_df.drop(
    columns=['Year', 'GREET1 sheet', 'Coproduct allocation method', 'GREET classification of coproduct'],
    inplace=True
)

MAC_df.rename(
    columns={
        'LCA: Unit (numerator)': 'CI replaced fuel: Unit (Numerator)',
        'LCA: Unit (denominator)': 'CI replaced fuel: Unit (Denominator)',
        'LCA_value': 'CI replaced fuel',
        'LCA_metric_x': 'Metric_replacing fuel',
        'LCA_metric_y': 'Metric_replaced fuel'
    },
    inplace=True
)

MAC_df.reset_index(drop=True, inplace=True)


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
# Map replaced fuels to MFSP
MAC_df = MAC_df.merge(
    corr_replaced_mfsp,
    on='Replaced Fuel',
    how='left'
)

# Map replaced fuels to EIA prices
MAC_df = MAC_df.merge(
    EIA_price[['Year', 'Value', 'Energy carrier', 'Cost basis', 'Unit']],
    how='left',
    left_on=['Production Year', 'Fuel_mapping_for_price'],
    right_on=['Year', 'Energy carrier']
)

# Rename columns
MAC_df.rename(
    columns={
        'Value': 'Cost_replaced fuel',
        'Cost basis': 'Cost basis_replaced fuel'
    },
    inplace=True
)

# Split unit strings into numerator and denominator
MAC_df[['Year_Cost_replaced fuel', 'Unit Cost_replaced fuel (Numerator)']] = (
    MAC_df['Unit'].str.split(' ', n=1, expand=True)
)

MAC_df[['Cost replaced fuel: Unit (Numerator)', 'Cost replaced fuel: Unit (Denominator)']] = (
    MAC_df['Unit Cost_replaced fuel (Numerator)'].str.split('/', n=1, expand=True)
)

# Drop unnecessary columns
MAC_df.drop(
    columns=['Unit Cost_replaced fuel (Numerator)', 'Energy carrier', 'Unit', 'Cost basis_replaced fuel'],
    inplace=True
)

# Final reset index
MAC_df.reset_index(drop=True, inplace=True)


# %%

# Step: Correct inflation of replacing fuel cost
# revising inflation adjustment code to improve performance
cpi_to = cpi.get(cost_year)
MAC_df['Year_Cost_replaced fuel']  = MAC_df['Year_Cost_replaced fuel'] .astype(int)
start_years = MAC_df['Year_Cost_replaced fuel'].unique()
cpi_start_dict = {year: cpi.get(year) for year in start_years}
MAC_df['cpi_start'] = MAC_df['Year_Cost_replaced fuel'].map(cpi_start_dict)
MAC_df['Adjusted Cost_replaced fuel'] = MAC_df['Cost_replaced fuel'] * (cpi_to / MAC_df['cpi_start'])


# %%

# Step: Unit check and conversions

# Unit check for Replaced Fuel

# Unit convert for liquid fuels (fuels except unit of energy kWh)
tmp_MAC_df = MAC_df.loc[(MAC_df['Cost replaced fuel: Unit (Denominator)'].isin(['kWh'])) |
                        (MAC_df['Biofuel Stream_LCA'].isin(['Hydrogen'])), :].copy()
MAC_df = MAC_df.loc[~((MAC_df['Cost replaced fuel: Unit (Denominator)'].isin(['kWh'])) |
                      (MAC_df['Biofuel Stream_LCA'].isin(['Hydrogen']))), :]

# Map Replaced fuel to 'GREET_Fuel', 'GREET_Fuel type' for conversion to unit of energy
MAC_df = pd.merge(MAC_df, corr_GGE_GREET_fuel_replaced, how='left',
                  left_on=['Replaced Fuel'],
                  right_on=['B2B fuel name']).reset_index(drop=True)

MAC_df = pd.merge(MAC_df, ob_units.hv_EIA[['GREET_Fuel', 'GREET_Fuel type', 'LHV', 'Unit']].drop_duplicates(),
                  how='left',
                  on=['GREET_Fuel', 'GREET_Fuel type']).reset_index(drop=True)
MAC_df[['LHV_numerator', 'LHV_denominator']
       ] = MAC_df['Unit'].str.split('/', n=1, expand=True)

MAC_df.loc[MAC_df['Cost replaced fuel: Unit (Denominator)'] == MAC_df['LHV_denominator'], 'Adjusted Cost_replaced fuel'] =\
    MAC_df.loc[MAC_df['Cost replaced fuel: Unit (Denominator)'] == MAC_df['LHV_denominator'], 'Adjusted Cost_replaced fuel'] /\
    MAC_df.loc[MAC_df['Cost replaced fuel: Unit (Denominator)']
               == MAC_df['LHV_denominator'], 'LHV']

MAC_df.loc[MAC_df['Cost replaced fuel: Unit (Denominator)'] == MAC_df['LHV_denominator'], 'Cost replaced fuel: Unit (Denominator)'] =\
    MAC_df.loc[MAC_df['Cost replaced fuel: Unit (Denominator)']
               == MAC_df['LHV_denominator'], 'LHV_numerator']

MAC_df[['Cost replaced fuel: Unit (Denominator)', 'Adjusted Cost_replaced fuel']] = \
    ob_units.unit_convert_df(
        MAC_df[[
            'Cost replaced fuel: Unit (Denominator)', 'Adjusted Cost_replaced fuel']],
        Unit='Cost replaced fuel: Unit (Denominator)',
        Value='Adjusted Cost_replaced fuel',
        if_unit_numerator=False,
        if_given_unit=True,
        given_unit='MJ').copy()
MAC_df['Adjusted Cost replaced fuel: Unit (Denominator)'] = 'MJ'

MAC_df.drop(columns=['B2B fuel name', 'GREET_Fuel', 'GREET_Fuel type', 'Fuel_mapping_for_price', 'LHV', 'Unit',
                     'LHV_numerator', 'LHV_denominator'], inplace=True)

# convert kWh to MJ
tmp_MAC_df['Adjusted Cost replaced fuel: Unit (Denominator)'] = tmp_MAC_df[
    'Cost replaced fuel: Unit (Denominator)']
tmp_MAC_df[['Adjusted Cost replaced fuel: Unit (Denominator)', 'Adjusted Cost_replaced fuel']] = \
    ob_units.unit_convert_df(
    tmp_MAC_df[[
        'Cost replaced fuel: Unit (Denominator)', 'Adjusted Cost_replaced fuel']],
    Unit='Cost replaced fuel: Unit (Denominator)',
    Value='Adjusted Cost_replaced fuel',
    if_unit_numerator=False,
    if_given_unit=True,
    given_unit='MJ').copy()
tmp_MAC_df['Adjusted Cost replaced fuel: Unit (Numerator)'] = 'USD'

# Concatenate the data frames
MAC_df = pd.concat([tmp_MAC_df, MAC_df]).reset_index(drop=True).copy()

MAC_df['Adjusted Cost replaced fuel: Unit (Numerator)'] = 'USD'


# %%
# Step: when decarbonized electric grid is considered, reference CI is updated for biopower pathways

if decarb_electric_grid:
    biopower_sc = MAC_df.loc[MAC_df['Case/Scenario']
                             .isin(biopower_scenarios), :]
    MAC_df = MAC_df.loc[~(MAC_df['Case/Scenario'].isin(biopower_scenarios)), :]
    tmpdf = decarb_elec_CI.loc[decarb_elec_CI['Parameter_B'] == 'Coproduct Credits',
                                ['Year',
                                 'LCA: Unit (numerator)',
                                 'LCA: Unit (denominator)',
                                 'LCA_value']]
    tmpdf.rename(columns={'LCA: Unit (numerator)': 'decarb_grid_CI: Unit (numerator)',
                           'LCA: Unit (denominator)': 'decarb_grid_CI: Unit (denominator)',
                           'LCA_value': 'decarb_grid_CI'}, inplace=True)
    biopower_sc = pd.merge(biopower_sc, tmpdf, how='left',
                           on=['Year'])

    # unit check
    check_units_decarbgrid = (biopower_sc['CI replaced fuel: Unit (Numerator)'] != biopower_sc['decarb_grid_CI: Unit (numerator)']) |\
        (biopower_sc['CI replaced fuel: Unit (Denominator)']
         != biopower_sc['decarb_grid_CI: Unit (denominator)'])
    biopower_sc = biopower_sc.loc[~check_units_decarbgrid]
    check_units_decarbgrid = biopower_sc.loc[check_units_decarbgrid, :]
    if check_units_decarbgrid.shape[0] > 0:
        print("Warning: The following decarb grid CI rows need attention as the units are not harmonized ..")
        print(check_units_decarbgrid)

    biopower_sc['CI replaced fuel'] = biopower_sc['decarb_grid_CI']
    biopower_sc.drop(columns=['decarb_grid_CI: Unit (numerator)',
                              'decarb_grid_CI: Unit (denominator)',
                              'decarb_grid_CI'], inplace=True)

    # Concatenate the data frames
    MAC_df = pd.concat([biopower_sc, MAC_df]).reset_index(drop=True).copy()

# %%
# Step: baseline check for MAC calculations

# For biopower scenarios, the 'Baseline for Biopower, 100% coal, w/o CCS, 650 MWe' Case/Scenario is considered
# for baseline MFSP and LCA
if adjust_biopower_baseline:
    biopower_baseline = MAC_df.loc[MAC_df['Case/Scenario'].isin(['Baseline for Biopower, 100% coal, w/o CCS, 650 MWe',
                                                                 'Baseline for Biopower, 100% coal, w/ CCS, 650 MWe']), :]
    MAC_df = MAC_df.loc[~(MAC_df['Case/Scenario'].isin(['Baseline for Biopower, 100% coal, w/o CCS, 650 MWe',
                                                        'Baseline for Biopower, 100% coal, w/ CCS, 650 MWe'])), :]

    MAC_df['baseline Case/Scenario'] = ''
    MAC_df.loc[MAC_df['Case/Scenario'].isin(['Biopower: 80% coal, w/o BECCS, 650 MWe',
                                             'Biopower: 100% biomass, w/o BECCS, 130 MWe']), 'baseline Case/Scenario'] =\
        'Baseline for Biopower, 100% coal, w/o CCS, 650 MWe'
    MAC_df.loc[MAC_df['Case/Scenario'].isin(['Biopower: 80% coal, w/ BECCS, 650 MWe',
                                             'Biopower: 100% biomass, w/ BECCS, 130 MWe']), 'baseline Case/Scenario'] =\
        'Baseline for Biopower, 100% coal, w/ CCS, 650 MWe'

    biopower_baseline = biopower_baseline[[
        'Case/Scenario', 'Production Year', 'Year', 'MFSP replacing fuel', 'Total LCA']]
    biopower_baseline.rename(columns={'Case/Scenario': 'baseline Case/Scenario',
                                      'MFSP replacing fuel': 'Adjusted Cost_replaced fuel_baseline',
                                      'Total LCA': 'CI replaced fuel_baseline'}, inplace=True)

    MAC_df = MAC_df.merge(biopower_baseline, how='left',
                          on=['baseline Case/Scenario', 'Production Year', 'Year']).reset_index(drop=True)

    MAC_df.loc[~(MAC_df['baseline Case/Scenario'].isin([''])), 'Adjusted Cost_replaced fuel'] =\
        MAC_df.loc[~(MAC_df['baseline Case/Scenario'].isin([''])),
                   'Adjusted Cost_replaced fuel_baseline']

    MAC_df.loc[~(MAC_df['baseline Case/Scenario'].isin([''])), 'CI replaced fuel'] =\
        MAC_df.loc[~(MAC_df['baseline Case/Scenario'].isin([''])),
                   'CI replaced fuel_baseline']

    MAC_df.drop(columns=['baseline Case/Scenario', 'Adjusted Cost_replaced fuel_baseline',
                'CI replaced fuel_baseline'], inplace=True)


# %%
# Step: Calculate MAC by Cost Items

# MAC = (MFSP_biofuel - MFSP_ref) / (CI_ref - CI_biofuel)
# Unit: ($/MJ - $/MJ) / (g/MJ - g/MJ) = $/g
MAC_df['MAC_calculated'] = (MAC_df['MFSP replacing fuel'] - MAC_df['Adjusted Cost_replaced fuel']) / \
                           (MAC_df['CI replaced fuel'] - MAC_df['Total LCA'])
MAC_df['MAC_calculated: Unit (numerator)'] = MAC_df['MFSP replacing fuel: Unit (numerator)']
MAC_df['MAC_calculated: Unit (denominator)'] = MAC_df['Total LCA: Unit (numerator)']
MAC_df['MAC_calculated'] = MAC_df['MAC_calculated'] * \
    1E6  # unit: $/MT CO2 avoided
MAC_df['MAC_calculated: Unit (denominator)'] = 'MT'

MAC_df['CI of replaced fuel higher'] = MAC_df['CI replaced fuel'] > MAC_df['Total LCA']
MAC_df['Cost of replaced fuel higher'] = MAC_df['Adjusted Cost_replaced fuel'] > MAC_df['MFSP replacing fuel']
MAC_df['Percent CI reduciton'] = (
    (MAC_df['CI replaced fuel'] - MAC_df['Total LCA']) / MAC_df['CI replaced fuel']) * 100
MAC_df['Percent MFSP increase'] = (
    (MAC_df['MFSP replacing fuel'] - MAC_df['Adjusted Cost_replaced fuel']) / MAC_df['Adjusted Cost_replaced fuel']) * 100

# Save interim data tables
if save_interim_files == True:
    MAC_df.to_csv(output_path_prefix + '/' + f_out_MAC)

print('    Elapsed time: ' + str(datetime.now() - init_time))

#%%

# Scale-up analysis
# calculate CI reduction (g GHG/MJ) and MFSP increase ($/MJ)
# Map feedstock flow rate (dry tons/MJ)
# Map feedstock availability (dry tons/year)
# Calculate net GHG reduction and net cost increase

if consider_scale_up_study:

    # calculate CI reduction (g GHG/MJ) and MFSP increase ($/MJ)
    scale_up = MAC_df.copy()
    
    scale_up['CI_reduction'] = MAC_df['CI replaced fuel'] - MAC_df['Total LCA']
    scale_up['MFSP_increase'] = MAC_df['MFSP replacing fuel'] - MAC_df['Adjusted Cost_replaced fuel']
    
    # Map and merge feedstock and fuel product flow rates (dry lb/hr, MJ/hr -> dry lb/MJ)
    tmpdf_feedstock = cost_items.loc[cost_items['Parameter_A'].isin(['Feedstock']), 
                   ['Case/Scenario', 'Stream_Flow', 'Stream_LCA', 'bm_cost_id',
                    'Flow: Unit (numerator)', 'Flow: Unit (denominator)', 'Flow']].drop_duplicates().reset_index(drop=True)
    tmpdf_feedstock.rename(columns={
        'Flow' : 'feedstock_Flow',
        'Stream_Flow' : 'feedstock_Stream_Flow',
        'Stream_LCA' : 'feedstock_Stream_LCA',
        'Flow: Unit (numerator)' : 'feedstock_Flow: Unit (numerator)', 
        'Flow: Unit (denominator)' : 'feedstock_Flow: Unit (denominator)'}, inplace=True)
    
    tmpdf_product = cost_items.loc[cost_items['Parameter_A'].isin(['Final Product']), 
                   ['Case/Scenario', 'Stream_Flow', 'Stream_LCA', 'bm_cost_id',
                    'Flow: Unit (numerator)', 'Flow: Unit (denominator)', 'Flow']].drop_duplicates().reset_index(drop=True)
    # aggregating energy products
    tmpdf_product = tmpdf_product.groupby(['Case/Scenario', 'bm_cost_id', 'Flow: Unit (numerator)', 'Flow: Unit (denominator)']).agg({'Flow':'sum'}).reset_index()
    tmpdf_product.rename(columns={
        'Flow' : 'product_Flow',
        'Flow: Unit (numerator)' : 'product_Flow: Unit (numerator)',
        'Flow: Unit (denominator)' : 'product_Flow: Unit (denominator)'}, inplace=True)
    
    # merge feedstock and product flows
    tmpdf = tmpdf_feedstock.merge(tmpdf_product, how='left', on=['Case/Scenario', 'bm_cost_id']).reset_index(drop=True)
    tmpdf['feedstock_per_product'] = tmpdf['feedstock_Flow'] / tmpdf['product_Flow']
    tmpdf['feedstock_per_product: Unit (numerator)'] = tmpdf['feedstock_Flow: Unit (numerator)']
    tmpdf['feedstock_per_product: Unit (denominator)'] = tmpdf['product_Flow: Unit (numerator)']
    
    # Calculate total GHG and USD per feedstock flow rate
    scale_up = scale_up.merge(tmpdf, how = 'left', on=['Case/Scenario', 'bm_cost_id']).reset_index(drop=True)
    scale_up['GHG_reduction_per_feedstock_flow'] = scale_up['CI_reduction'] / scale_up['feedstock_per_product']
    scale_up['cost_increase_per_feedstock_flow'] = scale_up['MFSP_increase'] / scale_up['feedstock_per_product']
    scale_up['GHG_reduction_per_feedstock_flow: Unit (numerator)'] = scale_up['CI replaced fuel: Unit (Numerator)']
    scale_up['GHG_reduction_per_feedstock_flow: Unit (denominator)'] = scale_up['feedstock_per_product: Unit (numerator)']
    scale_up['cost_increase_per_feedstock_flow: Unit (numerator)'] = scale_up['MFSP replacing fuel: Unit (numerator)']
    scale_up['cost_increase_per_feedstock_flow: Unit (denominator)'] = scale_up['feedstock_per_product: Unit (numerator)']
    
    # Map feedstock availability (dry tons/year)
    
    # Map biomass to feedstocks  
    scale_up['bm_types'] = scale_up['feedstock_Stream_LCA'].map(map_bm_types)
    
    # match to scale_up for all availability costs
    scale_up = scale_up.merge(bm, how='left', 
                              left_on = ['bm_types', 'bm_cost_id'],
                              right_on = ['bm_types', 'bm_cost']).reset_index(drop=True)
    
    # Calculate net GHG reduction and net cost increase
    scale_up['net_GHG_reduction'] = scale_up['GHG_reduction_per_feedstock_flow'] * scale_up['qty_dry_bm'] * 1E6 / 1E12 * 2204.6226 # dry ton to dry lb of biomass; grams GHG to million metric ton GHG
    scale_up['net_cost_increase'] = scale_up['cost_increase_per_feedstock_flow'] * scale_up['qty_dry_bm'] * 1E6 / 1E9 * 2204.6226 # dry ton to dry lb; USD to Billion USD
    scale_up['net_GHG_reduction: Unit'] = 'MM mt' # scale_up['GHG_reduction_per_feedstock_flow: Unit (numerator)'] 
    scale_up['net_cost_increase: Unit'] = 'B USD' #scale_up['cost_increase_per_feedstock_flow: Unit (denominator)']
    
    # Total fuel produced
    
    # Convert kWh to MJ
    scale_up.loc[scale_up['feedstock_per_product: Unit (denominator)'].isin(['kWh']), 'feedstock_per_product'] =\
        scale_up.loc[scale_up['feedstock_per_product: Unit (denominator)'].isin(['kWh']), 'feedstock_per_product'] / 3.6 # lb/kWh -> lb/MJ, 1 kWh = 3.6 MJ
    
    scale_up.loc[scale_up['feedstock_per_product: Unit (denominator)'].isin(['kWh']), 'feedstock_per_product: Unit (denominator)'] = 'MJ'
        
    scale_up.loc[scale_up['qty_dry_bm: Unit'].isin(['MM dt']), 'net_primary_fuel_produced'] =\
        1/(scale_up.loc[scale_up['qty_dry_bm: Unit'].isin(['MM dt']), 'feedstock_per_product']*0.0005) * \
        scale_up.loc[scale_up['qty_dry_bm: Unit'].isin(['MM dt']), 'qty_dry_bm'] /1E3 # MJ/dt * MM dt = Tera Joules -> Peta Joules
    
    scale_up.loc[scale_up['qty_dry_bm: Unit'].isin(['MM dt']), 'net_primary_fuel_produced: Unit'] = 'PJ'
    
# %%
# write data to the model dashboard tabs

if write_to_dashboard:
    print('Writing to Dashboard ..')

    pathway_names.rename(
        columns={'process|feedstock|product yield': 'Pathway Short Form'}, inplace=True)
    LCA_items = pd.merge(LCA_items, pathway_names[['Case/Scenario', 'Pathway Short Form']],
                         how='left', on='Case/Scenario').reset_index(drop=True)
    cost_items = pd.merge(cost_items, pathway_names[['Case/Scenario', 'Pathway Short Form']],
                          how='left', on='Case/Scenario').reset_index(drop=True)
    LCA_items_agg = pd.merge(LCA_items_agg, pathway_names[['Case/Scenario', 'Pathway Short Form']],
                             how='left', on='Case/Scenario').reset_index(drop=True)
    MFSP_agg = pd.merge(MFSP_agg, pathway_names[['Case/Scenario', 'Pathway Short Form']],
                        how='left', on='Case/Scenario').reset_index(drop=True)
    MAC_df = pd.merge(MAC_df, pathway_names[['Case/Scenario', 'Pathway Short Form']],
                      how='left', on='Case/Scenario').reset_index(drop=True)
    if consider_scale_up_study:
        scale_up = pd.merge(scale_up, pathway_names[['Case/Scenario', 'Pathway Short Form']],
                          how='left', on='Case/Scenario').reset_index(drop=True)

    # with ExcelApp() as app:
    with xw.App(visible=False) as app:

        wb = xw.Book(input_path_model + '/' + f_model)
        wb.app.calculation = 'manual'
        wb.app.screen_updating = False
        # wb.app.raw_value = True

        if consider_scale_up_study:
            sheet_1 = wb.sheets['scale_up']
            sheet_1.range(str(4) + ':1048576').clear_contents()
            sheet_1['A4'].options(index=False, chunksize=10000).value =\
                scale_up[[
                    'Pathway Short Form',
                    'Case/Scenario', 'Biofuel Stream_LCA', 'Energy_alloc_primary_fuel',
                           'Production Year', 'MFSP replacing fuel: Unit (numerator)',
                           'MFSP replacing fuel: Unit (denominator)', 'Adjusted Cost Year',
                           'bm_cost_id', 'MFSP replacing fuel', 'Metric_replacing fuel',
                           'Total LCA: Unit (numerator)', 'Total LCA: Unit (denominator)',
                           'Total LCA', 'Replaced Fuel', 'Parameter_B', 'Stream_Flow',
                           'Stream_LCA', 'CI replaced fuel: Unit (Numerator)',
                           'CI replaced fuel: Unit (Denominator)', 'CI replaced fuel',
                           'Metric_replaced fuel', 'Fuel_mapping_for_price', 'Source', 'Year',
                           'Cost_replaced fuel', 'Year_Cost_replaced fuel',
                           'Cost replaced fuel: Unit (Numerator)',
                           'Cost replaced fuel: Unit (Denominator)', 'Adjusted Cost_replaced fuel',
                           'Adjusted Cost replaced fuel: Unit (Denominator)',
                           'Adjusted Cost replaced fuel: Unit (Numerator)', 'MAC_calculated',
                           'MAC_calculated: Unit (numerator)',
                           'MAC_calculated: Unit (denominator)', 'CI of replaced fuel higher',
                           'Cost of replaced fuel higher', 'Percent CI reduciton',
                           'Percent MFSP increase', 'CI_reduction', 'MFSP_increase',
                           'feedstock_Stream_Flow', 'feedstock_Stream_LCA',
                           'feedstock_Flow: Unit (numerator)',
                           'feedstock_Flow: Unit (denominator)', 'feedstock_Flow',
                           'product_Flow: Unit (numerator)', 'product_Flow: Unit (denominator)',
                           'product_Flow', 'feedstock_per_product',
                           'feedstock_per_product: Unit (numerator)',
                           'feedstock_per_product: Unit (denominator)',
                           'GHG_reduction_per_feedstock_flow', 'cost_increase_per_feedstock_flow',
                           'GHG_reduction_per_feedstock_flow: Unit (numerator)',
                           'GHG_reduction_per_feedstock_flow: Unit (denominator)',
                           'cost_increase_per_feedstock_flow: Unit (numerator)',
                           'cost_increase_per_feedstock_flow: Unit (denominator)', 'bm_types',
                           'bm_cost', 'qty_dry_bm', 'bm_cost: Unit (numerator)',
                           'bm_cost: Unit (denominator)', 'bm_cost: USD year', 'qty_dry_bm: Unit',
                           'net_GHG_reduction', 'net_cost_increase', 'net_GHG_reduction: Unit',
                           'net_cost_increase: Unit',
                           'net_primary_fuel_produced',
                           'net_primary_fuel_produced: Unit'
                    ]]

        elif consider_variability_study & (consider_which_variabilities == 'Cost_Item'):
            
            sheet_1 = wb.sheets['lca']
            sheet_1.range(str(4) + ':1048576').clear_contents()
            sheet_1['A4'].options(index=False, chunksize=10000).value =\
                LCA_items_agg[[
                               'Pathway Short Form',
                               'Case/Scenario',
                               'LCA_metric',
                               'Total LCA: Unit (numerator)',
                               'Total LCA: Unit (denominator)',
                               'Production Year',
                               'Total LCA']]

            print ('writing to sheet mfsp_var ..')
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
                          'Pathway Short Form',
                          'Case/Scenario',
                          'Production Year',
                          'MFSP replacing fuel: Unit (numerator)',
                          'MFSP replacing fuel: Unit (denominator)',
                          'MFSP replacing fuel',
                          'Adjusted Cost Year'
                          ]]

            print ('writing to sheet mac_var_Cost_Item ..')
            sheet_1 = wb.sheets['mac_var_Cost_Item']
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
                        'Pathway Short Form',
                        'Case/Scenario',
                        'Biofuel Stream_LCA',
                        # 'Feedstock',
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
                        'Stream_LCA',
                        'CI replaced fuel: Unit (Numerator)',
                        'CI replaced fuel: Unit (Denominator)',
                        'CI replaced fuel',
                        'Metric_replaced fuel',
                        'Cost_replaced fuel',
                        'Year_Cost_replaced fuel',
                        'Cost replaced fuel: Unit (Numerator)',
                        'Cost replaced fuel: Unit (Denominator)',
                        'Adjusted Cost_replaced fuel',
                        'Adjusted Cost replaced fuel: Unit (Denominator)',
                        'Adjusted Cost replaced fuel: Unit (Numerator)',
                        'MAC_calculated',
                        'MAC_calculated: Unit (numerator)',
                        'MAC_calculated: Unit (denominator)',
                        'CI of replaced fuel higher',
                        'Cost of replaced fuel higher',
                        'Percent CI reduciton',
                        'Percent MFSP increase'
                        ]]

            sheet_1 = wb.sheets['lca_itm']
            print ('writing to sheet lca_itm ..')
            sheet_1 = wb.sheets['lca_itm']
            sheet_1.range(str(4) + ':1048576').clear_contents()
            sheet_1['A4'].options(index=False, chunksize=10000).value =\
                LCA_items[[
                           'Pathway Short Form',
                           'Case/Scenario',
                           'Parameter_A',
                           'Parameter_B',
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
                           # 'Year',
                           # 'GREET1 sheet',
                           # 'Coproduct allocation method',
                           # 'GREET classification of coproduct',
                           'LCA: Unit (numerator)',
                           'LCA: Unit (denominator)',
                           'LCA_value',
                           'LCA_metric',
                           'Total LCA',
                           'Total LCA: Unit (numerator)',
                           'Total LCA: Unit (denominator)',
                           # 'Biofuel Flow: Unit (numerator)',
                           # 'Biofuel Flow: Unit (denominator)',
                           # 'Biofuel Flow'
                           ]]

            
            print ('writing to sheet mfsp_itm_var')
            
            cost_items = cost_items.fillna('')
            cost_items = cost_items.astype(str)

            if len(cost_items) <= EXCEL_MAX_ROWS - 4:                
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
                            'Pathway Short Form',
                            'Case/Scenario',
                            'Parameter_A',
                            'Parameter_B',
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
                print(f"⚠️ DataFrame too big for Excel ({len(cost_items)} rows). Writing to parquet instead: dashboard_mfsp_itm_var.parquet")                
                cost_items.to_parquet(input_path_model + "/" + "dashboard_mfsp_itm_var.parquet", index=False)
                print(f"✅ Parquet file saved: dashboard_mfsp_itm_var.parquet")
            
            
        elif consider_variability_study & (consider_which_variabilities == 'Stream_LCA'):

            print ('writing to sheet lca_var')
            sheet_1 = wb.sheets['lca_var']
            sheet_1.range(str(4) + ':1048576').clear_contents()
            sheet_1['A4'].options(index=False, chunksize=10000).value =\
                LCA_items_agg[['variability_id',
                               'col_param',
                               'col_val',
                               'param_name',
                               'param_min',
                               'param_max',
                               'param_dist',
                               'dist_option',
                               'param_value',
                               'Pathway Short Form',
                               'Case/Scenario',
                               'LCA_metric',
                               'Total LCA: Unit (numerator)',
                               'Total LCA: Unit (denominator)',
                               'Production Year',
                               'Total LCA']]

            print ('writing to sheet mfsp')
            sheet_1 = wb.sheets['mfsp']
            sheet_1.range(str(4) + ':1048576').clear_contents()
            sheet_1['A4'].options(index=False, chunksize=10000).value =\
                MFSP_agg[[
                          'Pathway Short Form',
                          'Case/Scenario',
                          'Production Year',
                          'MFSP replacing fuel: Unit (numerator)',
                          'MFSP replacing fuel: Unit (denominator)',
                          'MFSP replacing fuel',
                          'Adjusted Cost Year'
                          ]]

            print ('writing to sheet mac_var_Stream_LCA')
            sheet_1 = wb.sheets['mac_var_Stream_LCA']
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
                        'Pathway Short Form',
                        'Case/Scenario',
                        'Biofuel Stream_LCA',
                        # 'Feedstock',
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
                        'Stream_LCA',
                        'CI replaced fuel: Unit (Numerator)',
                        'CI replaced fuel: Unit (Denominator)',
                        'CI replaced fuel',
                        'Metric_replaced fuel',
                        'Cost_replaced fuel',
                        'Year_Cost_replaced fuel',
                        'Cost replaced fuel: Unit (Numerator)',
                        'Cost replaced fuel: Unit (Denominator)',
                        'Adjusted Cost_replaced fuel',
                        'Adjusted Cost replaced fuel: Unit (Denominator)',
                        'Adjusted Cost replaced fuel: Unit (Numerator)',
                        'MAC_calculated',
                        'MAC_calculated: Unit (numerator)',
                        'MAC_calculated: Unit (denominator)',
                        'CI of replaced fuel higher',
                        'Cost of replaced fuel higher',
                        'Percent CI reduciton',
                        'Percent MFSP increase'
                        ]]

            print ('writing to sheet lca_itm_var')
            
            LCA_items = LCA_items.fillna('')
            LCA_items = LCA_items.astype(str)

            if len(LCA_items) <= EXCEL_MAX_ROWS - 4:
                sheet_1 = wb.sheets['lca_itm_var']
                sheet_1.range(str(4) + ':1048576').clear_contents()
                sheet_1['A4'].options(index=False, chunksize=10000).value =\
                    LCA_items[['variability_id',
                            'col_param',
                            'col_val',
                            'param_name',
                            'param_min',
                            'param_max',
                            'param_dist',
                            'dist_option',
                            'param_value',
                            'Pathway Short Form',
                            'Case/Scenario',
                            'Parameter_A',
                            'Parameter_B',
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
                            # 'Year',
                            # 'GREET1 sheet',
                            # 'Coproduct allocation method',
                            # 'GREET classification of coproduct',
                            'LCA: Unit (numerator)',
                            'LCA: Unit (denominator)',
                            'LCA_value',
                            'LCA_metric',
                            'Total LCA',
                            'Total LCA: Unit (numerator)',
                            'Total LCA: Unit (denominator)',
                            # 'Biofuel Flow: Unit (numerator)',
                            # 'Biofuel Flow: Unit (denominator)',
                            # 'Biofuel Flow'
                            ]]
            else:
                print(f"⚠️ DataFrame too big for Excel ({len(LCA_items)} rows). Writing to Parquet instead: dashboard_LCA_items.parquet")                
                LCA_items.to_parquet(input_path_model + "/" + "dashboard_LCA_items.parquet", index=False)
                print(f"✅ Parquet file saved: dashboard_LCA_items.parquet")

            print ('writing to sheet mfsp_itm')
            sheet_1 = wb.sheets['mfsp_itm']
            sheet_1.range(str(4) + ':1048576').clear_contents()
            sheet_1['A4'].options(index=False, chunksize=10000).value =\
                cost_items[[
                            'Pathway Short Form',
                            'Case/Scenario',
                            'Parameter_A',
                            'Parameter_B',
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

            # sheet_1 = wb.sheets['lca']
            wb.sheets['lca'].range(str(4) + ':1048576').clear_contents()
            wb.sheets['lca']['A4'].options(index=False, chunksize=10000).value =\
                LCA_items_agg[['Pathway Short Form',
                               'Case/Scenario',
                               'LCA_metric',
                               'Total LCA: Unit (numerator)',
                               'Total LCA: Unit (denominator)',
                               'Production Year',
                               'Total LCA']]

            # sheet_1 = wb.sheets['mfsp']
            wb.sheets['mfsp'].range(str(4) + ':1048576').clear_contents()
            wb.sheets['mfsp']['A4'].options(index=False, chunksize=10000).value =\
                MFSP_agg[['Pathway Short Form',
                          'Case/Scenario',
                          'Production Year',
                          'MFSP replacing fuel: Unit (numerator)',
                          'MFSP replacing fuel: Unit (denominator)',
                          'MFSP replacing fuel',
                          'Adjusted Cost Year'
                          ]]

            # sheet_1 = wb.sheets['mac']
            wb.sheets['mac'].range(str(4) + ':1048576').clear_contents()
            wb.sheets['mac']['A4'].options(index=False, chunksize=10000).value =\
                MAC_df[['Pathway Short Form',
                        'Case/Scenario',
                        'Biofuel Stream_LCA',
                        # 'Feedstock',
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
                        'Stream_LCA',
                        'CI replaced fuel: Unit (Numerator)',
                        'CI replaced fuel: Unit (Denominator)',
                        'CI replaced fuel',
                        'Metric_replaced fuel',
                        'Cost_replaced fuel',
                        'Year_Cost_replaced fuel',
                        'Cost replaced fuel: Unit (Numerator)',
                        'Cost replaced fuel: Unit (Denominator)',
                        'Adjusted Cost_replaced fuel',
                        'Adjusted Cost replaced fuel: Unit (Denominator)',
                        'Adjusted Cost replaced fuel: Unit (Numerator)',
                        'MAC_calculated',
                        'MAC_calculated: Unit (numerator)',
                        'MAC_calculated: Unit (denominator)',
                        'CI of replaced fuel higher',
                        'Cost of replaced fuel higher',
                        'Percent CI reduciton',
                        'Percent MFSP increase'

                        ]]

            # sheet_1 = wb.sheets['lca_itm']
            wb.sheets['lca_itm'].range(str(4) + ':1048576').clear_contents()
            wb.sheets['lca_itm']['A4'].options(index=False, chunksize=10000).value =\
                LCA_items[['Pathway Short Form',
                           'Case/Scenario',
                           'Parameter_A',
                           'Parameter_B',
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
                           # 'Year',
                           # 'GREET1 sheet',
                           # 'Coproduct allocation method',
                           # 'GREET classification of coproduct',
                           'LCA: Unit (numerator)',
                           'LCA: Unit (denominator)',
                           'LCA_value',
                           'LCA_metric',
                           'Total LCA',
                           'Total LCA: Unit (numerator)',
                           'Total LCA: Unit (denominator)',
                           # 'Biofuel Flow: Unit (numerator)',
                           # 'Biofuel Flow: Unit (denominator)',
                           # 'Biofuel Flow'
                           ]]

            # sheet_1 = wb.sheets['mfsp_itm']
            wb.sheets['mfsp_itm'].range(str(4) + ':1048576').clear_contents()
            wb.sheets['mfsp_itm']['A4'].options(index=False, chunksize=10000).value =\
                cost_items[['Pathway Short Form',
                            'Case/Scenario',
                            'Parameter_A',
                            'Parameter_B',
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

        # Write out electric grid CI at every run
        if decarb_electric_grid:            
            # sheet_1 = wb.sheets['EPS_CI']
            wb.sheets['EPS_CI'].range(str(4) + ':1048576').clear_contents()
            wb.sheets['EPS_CI']['A4'].options(index=False, chunksize=10000).value =\
                decarb_elec_CI[[
                    'Year',
                    'LCA: Unit (numerator)',
                    'LCA: Unit (denominator)',
                    'LCA_value',
                    'Parameter_B',
                    'LCA_metric']]
        else:
            tmpdf = corr_itemized_LCA.loc[corr_itemized_LCA['Stream_LCA'].isin(['Stationary Use: U.S. Mix']), :]
            wb.sheets['EPS_CI'].range(str(4) + ':1048576').clear_contents()
            wb.sheets['EPS_CI']['A4'].options(index=False, chunksize=10000).value =\
                tmpdf[[
                    'Year',
                    'LCA: Unit (numerator)',
                    'LCA: Unit (denominator)',
                    'LCA_value',
                    'Parameter_B',
                    'LCA_metric']]

        # resetting xlwings parameters before exiting
        wb.app.screen_updating = True

        print ('closing workbook ..')
        # wb.app.calculate()
        wb.save()
        wb.close()


print('    Elapsed time: ' + str(datetime.now() - init_time))

# %%

# Calculating percentage sensitivity when 'write_to_dashboard'=True and 'consider_variability_study'=True

# if write_to_dashboard:

# with xw.App(visible=False) as app:

# wb = xw.Book(input_path_model + '/' + f_model)

# if consider_variability_study:

# read mfsp

# read mfsp_var

# merge mfsp to mfsp_var

# wb.save()
# wb.close()
