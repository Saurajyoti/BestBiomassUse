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
fname_bt = 'Billion Ton Results_Best_Use.csv'

d = pd.read_csv(path_data + '\\' + fname_bt)
d.drop(columns = ['Unnamed: 0', ], axis=1, inplace=True)

d.head()
d.columns

# Monte Carlo sims for variables
# defining left, mod, right for triangular distribution

nsims = 1000 # number of samples to be collected

mc_frac_available = np.random.triangular(0.50, 0.75, 0.90, nsims)
mc_avg_conv_yield = np.random.triangular(2.58, 10.15, 25.00, nsims) # From Misclaneous cals.py -> calculate quantiles of biomass conversion efficiencies
mc_avg_fossil_CI = np.random.triangular(80, 90, 95, nsims)
mc_frac_bio_ghg_reduce = np.random.triangular(0.5, 0.7, 0.9, nsims) # Expected potential to reduce CI over fossil fuel CI by using the biomass feedstocks

mc_samples = mc_frac_available * mc_avg_conv_yield * mc_avg_fossil_CI * mc_frac_bio_ghg_reduce

# biomass prices to consider
biomass_price = [40, 60, 100] # the biomass price ranges to filter

# energy feedstocks to consider
feedstocks = [
'Barley straw',
'Biomass sorghum',
'CD waste',
'Citrus residues',
'Corn stover',
'Cotton gin trash',
'Cotton residue',
'Energy cane',
'Eucalyptus',
'Food waste',
'Hardwood, lowland, residue',
'Hardwood, lowland, tree',
'Hardwood, upland, residue',
'Hardwood, upland, tree',
'Hogs, 1000+ head',
'MSW wood',
'Milk cows, 500+ head',
'Miscanthus',
'Mixedwood, residue',
'Mixedwood, tree',
'Noncitrus residues',
'Oats straw',
'Other',
'Other forest residue',
'Other forest thinnings',
'Paper and paperboard',
'Pine',
'Plastics',
'Poplar',
'Primary mill residue',
'Rice hulls',
'Rice straw',
'Rubber and leather',
'Secondary mill residue',
'Softwood, natural, residue',
'Softwood, natural, tree',
'Softwood, planted, residue',
'Softwood, planted, tree',
'Sorghum stubble',
'Sugarcane bagasse',
'Sugarcane trash',
'Switchgrass',
'Textiles',
'Tree nut residues',
'Wheat straw',
'Willow',
'Yard trimmings'
]

d.query('`Biomass Price` in @biomass_price & Feedstock in @feedstocks', inplace = True)

d1 = d.groupby(['Year', 'Biomass Price'])['Production'].sum().reset_index().drop(columns = ['Biomass Price'], axis=1)

mc_samples = np.repeat(mc_samples, d1.shape[0])

d1 = pd.concat([d1] * nsims, ignore_index = True)

d1['sim_gCO2e'] = d1['Production'] * mc_samples * 907.1847 * 1e-12 # MMT CO2e

# plotting

#d1['Year'] = d1['Year'].astype('category')

sns.set_theme(style = 'darkgrid')

plt.figure(figsize = (10,10), )

g = sns.lineplot(x = 'Year', y = 'sim_gCO2e',
             data = d1)
g.set(xlabel = 'Year', ylabel = 'MMT CO2e emission reduction')
g.set(xticks = range(d1['Year'].min(), d1['Year'].max()+2, 2))
g.figure.savefig('Expected decarbonization by biomass.jpg', dpi = 400)
