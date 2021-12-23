# -*- coding: utf-8 -*-
"""
Created on Tue Dec 14 11:14:21 2021

@author: skar
@description: This script is to assess the overall decarbonization potential from biomass use
"""

"""
Calculation formula

biomass_pathway_ghg_reduction_potential, gCO2e = 
biomass dt x biomass_available_frac x dt_to_kg x avg_conv_yield MJ/kg x avg_fossil_CI gCO2e/MJ x biofuel_ghg_reduction
"""

import pandas as pd
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt

path_data = 'C:\\Users\\skar\\Box\\saura_self\\Proj - Best use of biomass\\data'
path_figs = 'C:\\Users\\skar\\Box\\saura_self\\Proj - Best use of biomass\\figs'
fname_bt = 'BT16_agcase_basecase_forestcase_basecase_spatialres_All.csv'

d = pd.read_csv(path_data + '\\' + fname_bt)
d.drop(columns = ['Unnamed: 0', ], axis=1, inplace=True)

d.head()
d.columns

# Monte Carlo sims for variables
# defining left, mod, right for triangular distribution

nsims = 10000 # number of samples to be collected

mc_frac_available = np.random.triangular(0.50, 0.75, 0.90, nsims)
mc_avg_conv_yield = np.random.triangular(2.58, 10.15, 25.00, nsims) # From Misclaneous cals.py -> calculate quantiles of biomass conversion efficiencies
mc_avg_fossil_CI = np.random.triangular(80, 90, 95, nsims)
mc_frac_bio_ghg_reduce = np.random.triangular(0.5, 0.7, 0.9, nsims) # Expected potential to reduce CI over fossil fuel CI by using the biomass feedstocks

mc_samples = mc_frac_available * mc_avg_conv_yield * mc_avg_fossil_CI * mc_frac_bio_ghg_reduce

# biomass prices to consider
biomass_price = [30,  40,  50,  60,  70,  80,  90, 100] # the biomass prices to consider


# Current data is already subsetted so no additional filtering occurs
# Subset the data to exclude conventional crops as well as 'Idle' and 'Pasture available' categories
d = d[~(d['Crop Type'] == 'Conventional') & ~(d['Feedstock'] == 'Idle') & ~(d['Feedstock'] == 'Pasture available')]

# subset Biomass price to consider
d.query('`Biomass Price` in @biomass_price', inplace = True)

d1 = d.groupby(['Year', 'Biomass Price'])['Production'].sum().reset_index()

mc_samples = np.repeat(mc_samples, d1.shape[0])

d1 = pd.concat([d1] * nsims, ignore_index = True)

d1['sim_gCO2e'] = d1['Production'] * mc_samples * 907.1847 * 1e-12 # MMT CO2e

# plotting

d1['Biomass Price'] = d1['Biomass Price'].astype('category')

sns.set_theme(style = 'white')

plt.figure(figsize = (10,10), )

g = sns.lineplot(x = 'Year', y = 'sim_gCO2e',
                 hue = 'Biomass Price',
             data = d1)
g.set(xlabel = 'Year', ylabel = 'MMT CO2e emission reduction')
g.set(xticks = range(d1['Year'].min(), d1['Year'].max()+2, 2))
g.figure.savefig(path_figs + '\\' + 'Expected decarbonization by biomass.jpg', dpi = 400)

# calculating quantiles from the data set and saving as a CSV file
class Quantile:
    # Writing class and setting up as functional calls ref: https://skeptric.com/pandas-aggregate-quantile/
    def __init__(self, q):
        self.q = q
        
    def __call__(self, x):
        return np.round(np.quantile(x.dropna(), self.q), 2)
    
year_filter = [2020, 2025, 2030, 2035, 2040]
d2 = d1.groupby(['Year', 'Biomass Price'])\
         .agg(sim_gCO2e_p10 = ('sim_gCO2e', Quantile(0.10)),
              sim_gCO2e_p90 = ('sim_gCO2e', Quantile(0.90)),
              sim_gCO2e_avg = ('sim_gCO2e', np.mean))
         
d2.query('Year in @year_filter', inplace = True)

d2.to_csv(path_data + '\\' + 'Expected decarbonization by biomass_quantiles.csv')    
#temp = d1[(d1['Year'] == 2021) & (d1['Biomass Price'].isin([30]))]['sim_gCO2e']
