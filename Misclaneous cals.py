# -*- coding: utf-8 -*-
"""
Created on Mon Dec 13 16:50:16 2021

@author: Saurajyoti Kar
"""

#%%
# Calculating quantiles of biomass conversion effciencies

"""
The data tab 'Fuel Pathways' from the data table 'Biofuel Pathways v2.0' is used.
Column E is the data column. Values in ranges are considered equally
feasible, hence every value in the range are repeated once. Values with unit of 
MJ fuel/kg feedstock are only considered and values with units MJ (product)/MJ (landfill gas)
are omitted.
"""

import numpy as np

d = np.hstack(
    (np.array([3.4, 1.6, 3.0, 4.2, 4.1, 12.8, 7, 9]), 
     np.arange(25, 27+1, 1), 
     np.arange(23, 27+1, 1), 
     np.arange(23, 25+1, 1), 
     np.array([40]), 
     np.arange(5.8, 9.3+1, 1),
     np.arange(8.7, 14+1, 1),
     np.arange(8.5, 14+1, 1), 
     np.array([7.2, 9.7, 3.1, 11.5, 3.5, 13.7, 7.2, 1.6, 0.8, 9.6, 7.0, 2.4, 0.3, 9.6, 6.4, 4.9, 1.0, 0.5]), 
     np.arange(11, 17+1, 1)
    ))
d
np.quantile(d, [0.10, 0.50, 0.90])
# output: array([ 2.58, 10.15, 25.  ])

#%%