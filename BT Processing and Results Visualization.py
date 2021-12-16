# -*- coding: utf-8 -*-
"""
Author: George G. Zaimes
Affiliation: Argonne National Laboratory
Description: Subsetting of BT Data and generation of biomass cost-supply curves
Dependencies: This script uses aggregated national-level BT results obtained from the 'Billion Ton Study' Script
"""
#%% Import Python Packages

# Import Python Packages
import pandas as pd
import plotly.io as pio
import plotly.express as px
pio.renderers.default='browser'
import seaborn as sns
import numpy as np


#%%  Set File Pathways

# Set filepath to location of Billion Ton Study Data:
#filepath = 'C:\\Users\\gzaimes\\Desktop\\Billion Ton\\BT Results\\Billion Ton Results_Best_Use.csv'

# Time-Series U.S. National Biomass Cost-Supply
f_bm = 'C:\\Users\\skar\\Box\\saura_self\\Proj - Best use of biomass\\data\\Billion Ton Results_Best_Use_aggregrate_biomass.csv'

# Time-Series U.S. National Biomass Cost-Supply by Feedstock
f_us = 'C:\\Users\\skar\\Box\\saura_self\\Proj - Best use of biomass\\data\\Billion Ton Results_Best_Use_National.csv'
 
# Set filepath to output result folder
#results_filepath = 'C:\\Users\\gzaimes\\Desktop\\Billion Ton\\BT Results'
figs_path = 'C:\\Users\\skar\\Box\\saura_self\\Proj - Best use of biomass\\figs'

#%% Data Aquisiton and Subsetting

# Read in the Aggregate National-level Biomass availability from the BT Study
df_bm = pd.read_csv(f_bm)
df_us = pd.read_csv(f_us)

#%% Figure: National U.S. Biomass Cost-Supply Curve 

# Create Time-Series U.S. National Biomass Cost-Supply Curve
df_bm_sub = df_bm.groupby(['Year', 'Biomass Price'], as_index=False)['Production'].sum()

# If desired, subset based on 5-year increments, comment out line if not applicable
df_bm_sub = df_bm_sub[df_bm_sub['Year'].isin([2020,2025,2030,2035,2040])]

# Plot Biomass Cost-Supply Curve
g = sns.lineplot(data = df_bm_sub, x= "Production",y="Biomass Price", marker="o", hue = 'Year', palette = 'Paired', sort=False)
g.set(xlabel='Biomass Supply (Dry Tons)', ylabel='Biomass Price ($/Dry Ton)')

#%% Figure: National U.S. Biomass Cost-Supply Curve(s), by Crop Category 

# Aggregate data to reflect the desired data-format 
df_us_bycat = df_us.groupby(['Year', 'Biomass Price', 'Crop Category'], as_index=False)['Production'].sum()

# If desired, subset based on 5-year increments, comment out line if not applicable
df_us_bycat = df_us_bycat[df_us_bycat['Year'].isin([2020,2025,2030,2035,2040])]

# Create FacetGrid plot with individual biomass cost-supply curves
g = sns.FacetGrid(df_us_bycat, col="Crop Category", hue = 'Year',sharey = 'col')
g.map(sns.lineplot,"Production","Biomass Price", marker="o", sort=False)
g.add_legend()
g.set_axis_labels("Biomass Supply (Dry Tons)", "Biomass Price ($/Dry Ton)")

#%% Figure: National U.S. Biomass-Specific Cost-Supply Curve(s), for Top 3 Biomass Types

# Aggregate data to reflect the desired data-format 
df_us_sub = df_us.groupby(['Year', 'Biomass Price', 'Feedstock'], as_index=False)['Production'].sum()

# If desired, subset based on 5-year increments, comment out line if not applicable
df_us_sub = df_us_sub[df_us_sub['Year'].isin([2020,2025,2030,2035,2040])]
df_us_sub = df_us_sub[df_us_sub['Feedstock'].isin(['Corn stover', 'Miscanthus', 'Switchgrass'])]

# Create FacetGrid plot with individual biomass cost-supply curves
g = sns.FacetGrid(df_us_sub, col="Feedstock", hue = 'Year',sharey = 'col')
g.map(sns.lineplot,"Production","Biomass Price", marker="o", sort=False)
g.add_legend()
g.set_axis_labels("Biomass Supply (Dry Tons)", "Biomass Price ($/Dry Ton)")

#%% Figure: National U.S. Availability, Sunburst Plot

# Subset data based on the desired year
#df_subset_agg_year = df_subset[df_subset['Year'] == 2020]

# Generate sunburst plot




