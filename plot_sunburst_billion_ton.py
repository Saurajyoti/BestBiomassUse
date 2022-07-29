# -*- coding: utf-8 -*-
"""
Created on Thu Dec  2 14:49:41 2021

@author: Saurajyoti Kar
@affiliation: Argonne National Laboratory
@description: sunburst plot of billion ton data set
@data source: Billion Ton Script.py output data
"""

import os
import pandas as pd
import plotly.express as px
import plotly.io as pio

pio.renderers.default='browser'

code_path = 'C:\\Users\\skar\\repos\\BestBiomassUse'
os.chdir(code_path)

filepath = 'C:\\Users\\skar\\Box\\saura_self\\Proj - Best use of biomass\\data'
input_file = 'Billion Ton Results_Best_Use.csv'

d = pd.read_csv(filepath + '\\' + input_file)
d.drop(columns = ['Unnamed: 0', ], axis=1, inplace=True)

fltr_scenario = 'Basecase, all energy crops'
fltr_year = 2020
d1 = d.loc[(d['Scenario'] == fltr_scenario) & (d['Year'] == fltr_year), ]

d1 = d1[['Crop Form', 'Crop Type', 'Land Source', 'Production']]
d1 = d1.dropna()

fig = px.sunburst(d1, path=['Crop Form', 'Crop Type', 'Land Source'], values='Production')
fig.show()
