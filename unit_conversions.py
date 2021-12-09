# -*- coding: utf-8 -*-
"""
Created on Wed Dec  8 11:40:22 2021

@author: skar
"""

""" A dictionary for unit conversions and a caller \
function to return 1 in case the unit conversation is not available """
    
unit1_per_unit2 = {
 """
 sources:
 https://www.nrcs.usda.gov/Internet/FSE_DOCUMENTS/nrcs142p2_022760.pdf
 https://www.eia.gov/energyexplained/units-and-calculators/
 """
 # feedstock or fuel based physical units:

 'Barley_lb_per_bu' : 48,
 'Corn_lb_per_bu' : 56,
 'Oats_lb_per_bu' : 32,
 'Sorghum_lb_per_bu' : 56,
 'Soybeans_lb_per_bu' : 60,
 'Wheat_lb_per_bu' : 60,
 'Barley_dry_per_wet' : 0.3, # fraction of dry matter
 'Corn_dry_per_wet' : 0.3,
 'Oats_dry_per_wet' : 0.3,
 'Sorghum_dry_per_wet' : 0.3,
 'Soybeans_dry_per_wet' : 0.3,
 'Wheat_dry_per_wet' : 0.3,
 'crudeoil_MMBtu_per_barrel' : 5.691,
 
 # physical only units
 
 'U.S.ton_per_lb' : 0.0005
 } 

def unit_conv (conv):
    if conv in unit1_per_unit2:
        return unit1_per_unit2[conv]
    else:
        return 1