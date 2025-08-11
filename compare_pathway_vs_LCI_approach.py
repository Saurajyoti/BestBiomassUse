# -*- coding: utf-8 -*-
"""
Copyright © 2025, UChicago Argonne, LLC
The full description is available in the LICENSE file at location:
    https://github.com/Saurajyoti/BestBiomassUse/blob/master/LICENSE

@Project: Best Use of Biomass
@Title: Compare pathway-reported TEA, LCA, and MAC vs. itemized calculations
@Authors: Saurajyoti Kar
@Contact: skar@anl.gov
@Affiliation: Argonne National Laboratory 

Created on Wed Feb 22 16:49:37 2023

"""

# compare itemized MFSP, LCA, MAC vs. pathway reported TEA, LCA, and calculated MAC

#%%
input_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - Best use of biomass/data'
interim_path_prefix = input_path_prefix + '/interim'
qa_path_prefix = input_path_prefix + '/QA'

f_calc_mac = 'mac.csv'
f_qa_mac = 'mac_for_pathways.csv'
f_qa_out = 'compare_approaches.xlsx'

#%%

import pandas as pd

#%%

calc_mac = pd.read_csv(interim_path_prefix+'/'+f_calc_mac, index_col=None)
repo_mac = pd.read_csv(qa_path_prefix+'/'+f_qa_mac, index_col=None)

calc_mac = calc_mac[['Case/Scenario',
                     #'Biofuel Flow Name',
                     #'Feedstock',
                     #'Adjusted Cost Year',
                     'MFSP replacing fuel', 
                     'MFSP replacing fuel: Unit (numerator)',
                     'MFSP replacing fuel: Unit (denominator)',          
                     #'LCA_metric', 
                     'Total LCA',
                     'Total LCA: Unit (numerator)',
                     'Total LCA: Unit (denominator)',                     
                     #'Total LCA (g per MJ)', 
                     #'Replaced Fuel',
                     #'CI replaced fuel: Unit (Numerator)',
                     #'CI replaced fuel: Unit (Denominator)',
                     'CI replaced fuel',
                     #'CI Elec0_replaced fuel', 
                     'Adjusted Cost_replaced fuel',
                     #'Cost basis_replaced fuel',
                     #'Cost replaced fuel: Unit (Numerator)',
                     #'Cost replaced fuel: Unit (Denominator)',
                     'MAC_calculated', 
                     'MAC_calculated: Unit (numerator)',
                     'MAC_calculated: Unit (denominator)'
                     ]]
repo_mac = repo_mac[['Case/Scenario',
                     #'Process', 
                     #'Feedstock',
                     #'Feedstock type',
                     #'Feedstock pool', 
                     #'Output', 
                     #'Fuel pool', 
                     'CI', 
                     'CI_unit_numerator', 
                     'CI_unit_denominator',
                     'MFSP', 
                     #'MFSP_unit_numerator',
                     #'MFSP_unit_denominator',
                     #'MFSP_Year',
                     'CI_replaced',
                     #'CI_replaced_unit_numerator', 
                     #'CI_replaced_unit_denominator',
                     'mfsp_replaced', 
                     #'mfsp_replaced_unit_numerator',
                     #'mfsp_replaced_unit_denominator', 
                     'percent_CI_abated',
                     'mac',
                     'mac_unit_numerator', 
                     'mac_unit_denominator'
                   ]]


compare = pd.merge(calc_mac, repo_mac, how='inner',
                   on='Case/Scenario').reset_index(drop=True)

# ad-hoc unit convert MFSP_calculated usd/mmbtu to usd/MJ
compare.loc[compare['MFSP replacing fuel: Unit (denominator)'].isin(['MMBtu']), 'MFSP replacing fuel'] =\
  compare.loc[compare['MFSP replacing fuel: Unit (denominator)'].isin(['MMBtu']), 'MFSP replacing fuel'] / 1055.06
compare.loc[compare['MFSP replacing fuel: Unit (denominator)'].isin(['MMBtu']), 'MFSP replacing fuel: Unit (denominator)'] = 'MJ'

# harmonize mfsp of replaced fuel, CI of replaced fuel and recalculate MAC for itemized approach
compare['Adjusted Cost_replaced fuel'] = compare['mfsp_replaced']
compare['CI replaced fuel'] = compare['CI_replaced']
compare['MAC_calculated'] = (compare['MFSP replacing fuel'] - compare['Adjusted Cost_replaced fuel']) / \
                            (compare['CI replaced fuel'] - compare['Total LCA'])


compare['diff_mfsp'] = (compare['MFSP'] - compare['MFSP replacing fuel']) / compare['MFSP'] * 100
compare['diff_CI'] = (compare['CI'] - compare['Total LCA']) / compare['CI'] * 100
compare['diff_mac'] = (compare['mac'] - compare['MAC_calculated']) / compare['mac'] * 100

compare.to_excel(qa_path_prefix+'/'+f_qa_out, index=False)
