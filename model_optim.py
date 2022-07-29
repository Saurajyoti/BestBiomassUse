# -*- coding: utf-8 -*-
"""
Created on Tue Jan  4 16:11:17 2022

@author: skar
"""

#%%
# Data loading

# BT16 data

# LCA data

# TEA data

# EIA cost data

# EIA fuel use projections data


#%%

# Data mapping, arrangement, and unit conversions

#%%

# Marginal GHG Abatement Cost (MAC) function

"""

Formula:      MAC               = (price_ref_fuel - price_bio_fuel) / (GHG_bio_fuel - GHG_ref_fuel)
Unit:   [USD / MT CO2e avoided] =       [USD / MMBtu]               /       [MT CO2e / MMBtu]

""" 

#%%

"""
Optimization design

Minimize: MAC

Constraints:
    1. biomass availability

Variables: 
    1. choice of quantity of feedstocks for the available feedstock-conversion-fuel pathways
    
Dimentions: temporal, 2020 to 2040

"""


#%%