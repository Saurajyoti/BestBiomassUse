# -*- coding: utf-8 -*-
"""
Project:EERE Decarbonization
Author: George G. Zaimes
Affiliation: Argonne National Laboratory
Date: 07/14/2021
Version: V1

Summary: This python script pulls time-series (2020-2050) data from EIA's AEO.
This data is combined with EPA's GHGI (2019), and used to project the GHG emissions
of the U.S. economy out until 2050.

Key Data Sources:
EIA AEO: https://www.eia.gov/outlooks/aeo/
EPA GHGI: https://cfpub.epa.gov/ghgdata/inventoryexplorer/chartindex.html
"""
#%%
#Import Python Libraries

# Python Packages
import pandas as pd
import requests
import matplotlib.pyplot as plt
import numpy as np
import seaborn as sns
from datetime import date
import matplotlib.ticker as ticker

#%%
# Set the EIA API Key (used to pull EIA data). EIA API Keys can be obtained via https://www.eia.gov/opendata/
# Create a Dictionary which maps EIA AEO cases to their API ID. AEO Cases represent different projections of the 
# U.S. Energy System. 'Looping' over the AEO dictionary, provides an easy way to extract EIA data across AEO cases.

# Obtain API Key from EIA, see URL: https://www.eia.gov/opendata/
api_key = ''

# Create a dictionary of AEO cases, and their corresponding API ID
aeo_case_dict = {'Reference case':'AEO.2021.REF2021.',
                 'High economic growth':'AEO.2021.HIGHMACRO.',
                 'Low economic growth':'AEO.2021.LOWMACRO.',
                 'High oil price':'AEO.2021.HIGHPRICE.',
                 'Low oil price':'AEO.2021.LOWPRICE.',
                 'High oil and gas supply':'AEO.2021.HIGHOGS.',
                 'Low oil and gas supply':'AEO.2021.LOWOGS.',
                 'High renewable cost':'AEO.2021.HIRENCST.',
                 'Low renewable cost':'AEO.2021.LORENCST.',
                 }

#%%
# Create a function to store sector-wide energy consumption and CO2 emissions

def eia_sector_import (sector, aeo_case): 
    """
    

    Parameters
    ----------
    sector : str
        Pulls EIA data for the selected U.S. sector. Choose one of the following options:
            'Residential'
            'Transportation'
            'Commercial'
            'Industrial'
            'Electric Power'

    aeo_case : str
        Pulls EIA data for the selected AEO case. Choose one of the following options:
            'Reference case'
            'High economic growth'
            'Low economic growth'
            'High oil price'
            'Low oil price'
            'High oil and gas supply'
            'Low oil and gas supply'
            'High renewable cost'
            'Low renewable cost'
    Returns
    -------
    eia_df : Pandas DataFrame
        Output is a pandas DataFrame that contains detailed estimates of energy consumption
        for the selected U.S. sector and AEO case over the 2020 to 2050 time-horizon.

    """
    
    # Create an temporary list to store time-series results from EIA
    temp_list = [] 
    
    # Load in EIA's AEO Series IDs / AEO Keys
    df_aeo_key = pd.read_excel('C:\\Users\\gzaimes\\Desktop\\EERE Decarbonization\\EIA AEO Data_v1.xlsx', sheet_name = sector)
    
    # Each sector has multiple data series that document the end-use applications, materials, and energy consumption. 
    # Based on the user-selected sector, loop through each data series, pulling EIA data based on the series ID / API Key
    # For each series store relevant metadata and results across the 2020 to 2050 timeframe. Append the results to the
    # temporary list. After looping through all series, concatenate results into one large dataframe. This dataframe
    # contains sector-wide energy consumption
    for row in df_aeo_key.itertuples():                
        series_id = aeo_case_dict[aeo_case] + row[7]
        url = 'http://api.eia.gov/series/?api_key=' + api_key +'&series_id=' + series_id
        r = requests.get(url)
        json_data = r.json()
        df_temp = pd.DataFrame(json_data.get('series')[0].get('data'),
                               columns = ['Date', 'Value'])
        df_temp['Data Source'] = row[1]
        df_temp['AEO Case'] = aeo_case
        df_temp['Sector'] = row[2]
        df_temp['Subcategory 1'] = row[3]      
        df_temp['Subcategory 2'] = row[4]   
        df_temp['Subcategory 3'] = row[5]
        df_temp['Energy Carrier'] = row[6]  
        df_temp['Metric'] = row[8]  
        df_temp['Unit'] = row[9]  
        df_temp['Series Id'] = series_id
        temp_list.append(df_temp)
    eia_sector_df = pd.concat(temp_list, axis=0)
    eia_sector_df = eia_sector_df.reset_index(drop=True)
    return eia_sector_df

# Use the function to create datasets for sector-specific and AEO cases: 
eia_sector_df = eia_sector_import(sector = 'Residential', aeo_case = 'Reference case')

#%%
# Create a function to store results across multiple combinations of AEO-cases and sectors

def eia_multi_sector_import (sectors, aeo_cases): 
    """
    

    Parameters
    ----------
    sectors : List
        Input is a list of EIA-AEO Sectors, choose one or more of the following sectors:
            'Residential'
            'Transportation'
            'Commercial'
            'Industrial'
            'Electric Power'
            
    aeo_cases : List
        Input is a list of AEO cases, choose one or more of the following cases:
            'Reference case'
            'High economic growth'
            'Low economic growth'
            'High oil price'
            'Low oil price'
            'High oil and gas supply'
            'Low oil and gas supply'
            'High renewable cost'
            'Low renewable cost'
    Returns
    -------
    eia_economy_wide_df : Pandas DataFrame
        Output is a pandas DataFrame that contains detailed estimates of energy consumption
        for the selected U.S. sectors and AEO cases over the 2020 to 2050 time-horizon.

    """
    
    # Create an temporary list to store results
    temp_list = [] 
    
    #Loop through every combination of AEO Case and Sector 
    for aeo_case in aeo_cases:
        for sector in sectors:
            eia_df_temp = eia_sector_import(sector = sector, aeo_case = aeo_case)
            temp_list.append(eia_df_temp)
    
    # Concatenate results into one large dataframe, containing all combinations of sectors and cases
    eia_economy_wide_df = pd.concat(temp_list, axis=0)
    eia_economy_wide_df = eia_economy_wide_df.reset_index(drop=True)
    
    return eia_economy_wide_df

eia_multi_sector_df = eia_multi_sector_import(sectors = ['Residential',
                                                         'Transportation',
                                                         'Commercial',
                                                         'Industrial',
                                                         'Electric Power'
                                                         ],
                                              
                                              aeo_cases = ['Reference case'
                                                           ]
                                              )

eia_multi_sector_df.to_csv('C:\\Users\\gzaimes\\Desktop\\EERE Decarbonization\\EIA Dataset.csv')