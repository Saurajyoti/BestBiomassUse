# -*- coding: utf-8 -*-
"""
Author: George G. Zaimes
Affiliation: Argonne National Laboratory
Description: Pre-processing and analysis of resource availability from the Billion Ton Study
Data Source: https://bioenergykdf.net/sites/default/files/BillionTonDownloads/billionton_county_all.zip
"""
#%%

# Import Python Packages
import pandas as pd

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
    bt_df = bt_df[bt_df['Year'].isin(range(start_year, end_year + 1))]
    if feedstock is not None:
        bt_df = bt_df[bt_df['Feedstock'].isin(feedstock)]
    if biomass_price is not None:
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
    if spatial_res == 'National':
        var_list = bt_df.columns.tolist()
        remove_list = ['County', 'State', 'fips', 'Yield', 'Yield Unit', 'Production', 'Harvested Acres', 'Land Area','Diameter Class', 'Operation Type',
       'Owner', 'supply Class', 'Supply Target']
        agg_list = [i for i in var_list if i not in remove_list]
        bt_df = bt_df.groupby(agg_list, dropna = False, as_index = False)['Production','Harvested Acres','Land Area'].sum()
        bt_df['Yield Unit'] = bt_df['Production Unit'] + '/ac'
        bt_df['Yield'] = bt_df['Production'] / bt_df['Harvested Acres']
    elif spatial_res == 'State':
        var_list = bt_df.columns.tolist()
        remove_list = ['County', 'Yield', 'Yield Unit', 'Production', 'Harvested Acres', 'Land Area','Diameter Class', 'Operation Type',
       'Owner', 'supply Class', 'Supply Target', 'fips']
        agg_list = [i for i in var_list if i not in remove_list]
        bt_df = bt_df.groupby(agg_list, dropna = False, as_index = False)['Production','Harvested Acres','Land Area'].sum()
        bt_df['Yield Unit'] = bt_df['Production Unit'] + '/ac'
        bt_df['Yield'] = bt_df['Production'] / bt_df['Harvested Acres']
    return bt_df
 
#%%    
 
# Run the function (bt_scenario) based on user-defined setting, save the results as a variable 'bt_case'
bt_case = bt_scenario(ag_case = 'basecase', 
                      forestry_case = 'basecase', 
                      waste_case = 'basecase',
                      start_year = 2020,
                      end_year = 2050,
                      feedstock = None ,
                      biomass_price = None,
                      price_logic = 'less than or equal to',
                      spatial_res = 'National')

# Save the results as a CSV file 
bt_case = bt_case.sort_values(by=['Feedstock','Biomass Price'], ascending = True)
bt_case.to_csv(results_filepath + '\Billion Ton Results_Best_Use.csv')


