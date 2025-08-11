# -*- coding: utf-8 -*-
"""
Copyright © 2025, UChicago Argonne, LLC
The full description is available in the LICENSE file at location:
    https://github.com/Saurajyoti/BestBiomassUse/blob/master/LICENSE

@Project: Best Use of Biomass
@Title: Pre-processing and analysis of resource availability from the Billion Ton Study
@Authors: George G. Zaimes, Saurajyoti Kar
@Contact: skar@anl.gov
@Affiliation: Argonne National Laboratory 
@Data: https://bioenergykdf.net/sites/default/files/BillionTonDownloads/billionton_county_all.zip

@Created: March 8, 2023

"""
#%%

# Import Python Packages
import pandas as pd
import numpy as np
import os
import time

# Import user defined modules
code_path = 'C:\\Users\\skar\\repos\\BestBiomassUse'
os.chdir(code_path)

import unit_conversions as ut

#%%

# Set filepath to location of Billion Ton Study Data:
#filepath = 'C:\\Users\\gzaimes\\Desktop\\Billion Ton\\BT Data'
filepath = 'C:\\Users\\skar\\data\\Resource Availability\\Billion Ton Study\\Full Dataset [County-Level]'
 
# Set filepath to output result files
#results_filepath = 'C:\\Users\\gzaimes\\Desktop\\Billion Ton\\BT Results'
results_filepath = 'C:\\Users\\skar\\Box\\saura_self\\Proj - Best use of biomass\\data'

#%%


# Create a function ('bt_sceario') to process, subset, and aggregate data from the Billion Ton Study              
def bt_scenario(ag_case,
                forestry_case,
                waste_case,
                start_year,
                end_year,
                feedstock,
                biomass_price,
                price_logic,
                spatial_res):
    """
    Parameters
    ----------
    ag_case : String, or None
        This parameter defines which Billion Ton agricultural scenario will be loaded in the simulation.
        Choose between four different cases, a Basecase [ag_case = 'basecase'], 2% yield increase [ag_case ='2pct'],
        3% yield increase [ag_case ='3pct'], or 4% yield increase [ag_case ='4pct']. To exclude the agricultural sector from
        the results set ag_case = None
    forestry_case : String, or None
        This parameter defines which Billion Ton forestry scenario will be loaded in the simulation.
        Choose between six different cases, a Basecase [mediumhousinglowenergy = 'basecase'], 
        High housing high energy case [forestry_case ='highhousinghighenergy'],
        High housing medium energy case [forestry_case ='highhousingmediumenergy'], 
        High housing low energy case [forestry_case ='highhousinglowenergy'],
        Medium housing high energy case [forestry_case ='mediumhousinghighenergy'],
        Medium housing medium energy case [forestry_case ='mediumhousingmediumenergy'], 
        To exclude the forestry sector from the results set forestry_case = None
    waste_case : String, or None
        This parameter defines which Billion Ton waste scenario will be loaded in the simulation.
        At present, there is only one waste case, a Basecase [waste_case = 'basecase'].
        To exclude the waste sector from the results set waste_case = None
    start_year : Integer
        Choose the start year of the simulation (e.g. start_year = 2019). Billion Ton Data ranges from year 2014 to 2040,
        select a start year within this range.
    end_year : Integer
        Choose the end year of the simulation (e.g. end_year = 2019). Billion Ton Data ranges from year 2014 to 2040,
        select an end year within this range. Please ensure that end_year >= start_year. If end_year = start_year, 
        only results for that particular year will be populated. If end_year > start_year, results will be populated over the
        entire time horizon.
    feedstock : List, or None
        Subset the biofeedstocks to include in the analysis by passing a list of feedstocks, e.g. feedstock = ['Corn', 'Switchgrass']
        will subset the results to include only corn and Switchgrass. If you wish to include All feedstocks (e.g. no subsetting)
        please set feedstock = None. A list of all feedstocks is provided as follows:
            Barley
            Corn
            Cotton
            Hay
            Idle
            Oats
            Oats straw
            Pasture available
            Rice
            Sorghum
            Soybeans
            Wheat
            Wheat straw
            Willow
            Barley straw
            Corn stover
            Poplar
            Sorghum stubble
            Miscanthus
            Pine
            Switchgrass
            Biomass sorghum
            Eucalyptus
            Energy cane
            Hardwood, lowland, residue
            Hardwood, upland, residue
            Mixedwood, residue
            Softwood, natural, residue
            Softwood, planted, residue
            Hardwood, lowland, tree
            Hardwood, upland, tree
            Mixedwood, tree
            Softwood, natural, tree
            Softwood, planted, tree
            CD waste
            Citrus residues
            Noncitrus residues
            Other
            Other forest residue
            Paper and paperboard
            Plastics
            Primary mill residue
            Rubber and leather
            Secondary mill residue
            Textiles
            Tree nut residues
            Cotton gin trash
            Cotton residue
            Hogs, 1000+ head
            MSW wood
            Milk cows, 500+ head
            Rice hulls
            Sugarcane bagasse
            Sugarcane trash
            Rice straw
            Yard trimmings
            Other forest thinnings
            Food waste
    biomass_price : Float, Int, or None
        Subset the results based on the price of the biomass feedstock. Set the price logic, (e.g. >, =, <, etc.) using the
        variable price_logic. For example, if one wanted to consider only feedstocks with a price equal to $30/dt or less, set
        biomass_price = 30, and set price_logic = 'less than or equal to'. If you wish to include results across all biomass prices
        set biomass_price = None
    price_logic : String, or None
        Set the price_logic for the biomass_price variable. Select between the following options: 'less than', 'less than or equal to',
        'greater than', 'greater than or equal to', or 'equal to'
    spatial_res : String
        Set the spatial resolution of the dataset, select between 'County','State', or 'National'. The spatial resolution of 
        the original BT data is at the County level. If a user chooses either 'National' or 'State', county-level data will be
        aggregated to the user selected level of aggregation.
        If spatial_res = 'National'results will be aggregated at the national level. 
        If spatial_res = 'State', resilts will be aggregated at the state level.

    Returns
    -------
    bt_df : DataFrame
        This DataFrame contains results from the Billion Ton Study, subsetted and aggregated based on user-defined variables.
   """
    bt_df = pd.DataFrame()
    if ag_case is None:
        agri = pd.DataFrame()
    else:
        agri = pd.read_csv(filepath +'\\billionton_county_agriculture_'+ ag_case +'.csv')
    if forestry_case is None:
        forestry = pd.DataFrame()
    elif forestry_case == 'basecase':
        forestry = pd.read_csv(filepath +'\\billionton_county_forestry_mediumhousinglowenergy.csv')
    else:
        forestry = pd.read_csv(filepath +'\\billionton_county_forestry_'+ forestry_case +'.csv')
    if waste_case is None:
        waste = pd.DataFrame()
    else:
        waste = pd.read_csv(filepath +'\\billionton_county_wastes.csv')
    bt_df = bt_df.append([agri, forestry, waste])
    
    # free up some memory
    del agri
    del forestry
    del waste
    
    # Subset the data to exclude conventional crops as well as 'Idle' and 'Pasture available' categories
    bt_df = bt_df[~(bt_df['Crop Type'] == 'Conventional') & ~(bt_df['Feedstock'] == 'Idle') & ~(bt_df['Feedstock'] == 'Pasture available')]
    
    bt_df = bt_df[bt_df['Year'].isin(range(start_year, end_year + 1))]
    if feedstock is not None:
        bt_df = bt_df[bt_df['Feedstock'].isin(feedstock)]
    elif biomass_price is not None:
        if price_logic == 'less than':
            bt_df = bt_df[bt_df['Biomass Price'] < biomass_price]
        elif price_logic == 'less than or equal to':
            bt_df = bt_df[bt_df['Biomass Price'] <= biomass_price]
        elif price_logic == 'greater than':
            bt_df = bt_df[bt_df['Biomass Price'] > biomass_price]
        elif price_logic == 'greater than or equal to':
            bt_df = bt_df[bt_df['Biomass Price'] >= biomass_price]
        elif price_logic == 'equal to':
            bt_df = bt_df[bt_df['Biomass Price'] == biomass_price]
    
    # unit conversions
    to_unit = 'lb'
    bt_df['unit_conv'] = bt_df['Feedstock'] + '_' + to_unit + '_per_' + bt_df['Production Unit']
    bt_df['Production'] = np.where(
        [x in ut.unit1_per_unit2 for x in bt_df['unit_conv'] ],
        bt_df['Production'] * bt_df['unit_conv'].map(ut.unit1_per_unit2),
        bt_df['Production'])
    bt_df['Production Unit'] = np.where(
        [x in ut.unit1_per_unit2 for x in bt_df['unit_conv'] ],
        to_unit,
        bt_df['Production Unit'])
    
    to_unit = 'U.S.ton'
    bt_df['unit_conv'] = to_unit + '_per_' + bt_df['Production Unit']
    bt_df['Production'] = np.where(
        [x in ut.unit1_per_unit2 for x in bt_df['unit_conv'] ],
        bt_df['Production'] * bt_df['unit_conv'].map(ut.unit1_per_unit2),
        bt_df['Production'])
    bt_df['Production Unit'] = np.where(
        [x in ut.unit1_per_unit2 for x in bt_df['unit_conv'] ],
        to_unit,
        bt_df['Production Unit'])
    
    to_unit = 'dry'
    bt_df['unit_conv'] = bt_df['Feedstock'] + '_' + to_unit + '_per_' + 'wet'
    bt_df['Production'] = np.where(
        [x in ut.unit1_per_unit2 for x in bt_df['unit_conv'] ],
        bt_df['Production'] * bt_df['unit_conv'].map(ut.unit1_per_unit2),
        bt_df['Production'])
    bt_df['Production Unit'] = np.where(
        [x in ut.unit1_per_unit2 for x in bt_df['unit_conv'] ],
        'dt',
        bt_df['Production Unit'])
    bt_df.drop(columns = ['unit_conv', ], axis=1, inplace=True)

    # aggregrating to the required spatial level
   
    if spatial_res == 'County' or spatial_res == None:        
        grp_cols = ['Year', 'County', 'fips', 'State', 'Land Source', 'Crop Form', 'Crop Category', 'Crop Type', 
                    'Feedstock', 'Biomass Price', 'Production', 'Production Unit',
                    'Harvested Acres', 'Land Area']
    elif spatial_res == 'State':
        grp_cols = ['Year', 'State', 'Land Source', 'Crop Form', 'Crop Category', 'Crop Type', 
                    'Feedstock', 'Biomass Price', 'Production', 'Production Unit',
                    'Harvested Acres', 'Land Area']
    elif spatial_res == 'National':
        grp_cols = ['Year', 'Land Source', 'Crop Form', 'Crop Category', 'Crop Type', 
                    'Feedstock', 'Biomass Price', 'Production', 'Production Unit',
                    'Harvested Acres', 'Land Area']
    elif spatial_res == 'aggregrate_biomass':
        grp_cols = ['Year', 'Biomass Price', 'Production', 'Production Unit',
                    'Harvested Acres', 'Land Area']
    
    # filling in grouping variables when blank   
    bt_df.loc[bt_df['Land Source'].isnull(), 'Land Source'] = 'Other'
    bt_df.loc[bt_df['Crop Form'].isnull(), 'Crop Form'] = 'Other'

    
    # aggregrate only if a spatial level is mentioned
    if spatial_res != None:
        bt_df = bt_df.groupby([i for i in grp_cols if i not in ['Production','Harvested Acres','Land Area']], dropna = False, as_index = False)\
                               [['Production','Harvested Acres','Land Area']].sum().reset_index(drop=True)
    else:
        bt_df = bt_df[grp_cols]
        bt_df = bt_df.sort_values(grp_cols).reset_index(drop=True)
     
    
    # calculating incremental production, incremental cost, total cost, cummulative cost and average price at county level
    
    # create columns to compute
    bt_df[['min_biomass_price', 'inc_prod', 'inc_cost', 'total_cost', 'cuml_cost', 'avg_price']] = -1
       
    bt_df['min_biomass_price'] = bt_df.groupby(
        [i for i in grp_cols if i not in ['Production', 'Production Unit', 'Harvested Acres','Land Area', 'Biomass Price']])\
        ['Biomass Price'].transform('min')
    
    prod_shift = bt_df['Production'].shift(1).fillna(0)
    biomass_price_shift = bt_df['Biomass Price'].shift(1).fillna(0)
    
    conditions = [
        bt_df['Biomass Price'].values == bt_df['min_biomass_price'].values,
        True
        ]
    choices = [
        bt_df['Production'],
        bt_df['Production'] - prod_shift
        ]
    bt_df['inc_prod'] = np.select(conditions, choices, default = 'NA').astype('float')
    
    conditions = [
        bt_df['Biomass Price'].values == bt_df['min_biomass_price'].values,
        True
        ]
    choices = [
        bt_df['Biomass Price'],
        (bt_df['Biomass Price'] + biomass_price_shift) / 2
        ]
    bt_df['inc_cost'] = np.select(conditions, choices, default = 'NA').astype('float')
    
    bt_df['total_cost'] = bt_df['inc_prod'].multiply(bt_df['inc_cost'], fill_value=0)
    
    bt_df['cuml_cost'] = bt_df.groupby(
        [i for i in grp_cols if i not in ['Production', 'Production Unit', 'Harvested Acres','Land Area', 'Biomass Price']])\
        ['total_cost'].transform('cumsum')
    
    bt_df['avg_price'] = bt_df['cuml_cost'] / bt_df['Production']
    
    # calculating yield
    bt_df['Yield Unit'] = bt_df['Production Unit'] + '/ac'    
    bt_df['Yield'] = bt_df['Production'] / bt_df['Harvested Acres']
                  
    # Testing
    #tmp3 = bt_df.query("State == 'Alabama' & fips == 1011 & Year == 2025 & `Crop Form` == 'Herbaceous' & \
    #            `Crop Category` == 'Agriculture' & `Crop Type` == 'Energy' & Feedstock == 'Miscanthus' ")
    
    return bt_df
 
#%%    
 
# Run the function (bt_scenario) based on user-defined setting, save the results as a variable 'bt_case'

#spatial_res = [None]
#spatial_res = ['County']
#spatial_res = ['State']
#spatial_res = ['National']
spatial_res = ['All', 'County', 'State', 'National', 'aggregrate_biomass']

ag_case = ['basecase', '2pct', '3pct', '4pct']

forestry_case = ['basecase', 'highhousinghighenergy', 'highhousingmediumenergy', 'highhousinglowenergy', 'mediumhousinghighenergy', 'mediumhousingmediumenergy']

feedstock = None # no filtering, all feedstocks pulled

def call_func (spatial_res, ag_case, forestry_case):
    
    if spatial_res == 'All':
        spatial_res_param = None
    else:
        spatial_res_param = spatial_res
    
    print('\n**Working on spatial resolution: ' + spatial_res +
          '\n  Agriculture case: ' + ag_case + '\n  Forestry case: ' + forestry_case + '\n')
    bt_case = bt_scenario(ag_case = ag_case, 
                          forestry_case = forestry_case, 
                          waste_case = 'basecase',
                          start_year = 2020,
                          end_year = 2040,
                          feedstock = feedstock,
                          biomass_price = None,
                          price_logic = 'less than or equal to',
                          spatial_res = spatial_res_param)

    # Save the results as a CSV file or a python object
    bt_case.to_csv(results_filepath + '\\' + 'BT16_agcase_' + ag_case +  '_forestcase_' + forestry_case + '_spatialres_' + spatial_res + '.csv')
    

startT = time.time()

[[[call_func(res, ag, fr) for res in spatial_res] for ag in ag_case] for fr in forestry_case]

print ('Execution duration in minutes: ' + str((time.time() - startT)/60))
