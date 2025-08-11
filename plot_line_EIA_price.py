# -*- coding: utf-8 -*-
"""
Copyright © 2025, UChicago Argonne, LLC
The full description is available in the LICENSE file at location:
    https://github.com/Saurajyoti/BestBiomassUse/blob/master/LICENSE

@Project: Best Use of Biomass
@Title: Plot line plots of EIA fuel prices
@Authors: Saurajyoti Kar
@Contact: skar@anl.gov
@Affiliation: Argonne National Laboratory

Created on Tue Dec  7 17:59:41 2021
"""

import os
import pandas as pd
import seaborn as sns

code_path = 'C:\\Users\\skar\\repos\\BestBiomassUse'
os.chdir(code_path)

import unit_conversions as ut

#pio.renderers.default='browser'
#pio.kaleido.scope.default_format = "jpg"

filepath = 'C:\\Users\\skar\\Box\\saura_self\\Proj - Best use of biomass\\data'
figpath = 'C:\\Users\\skar\\Box\\saura_self\\Proj - Best use of biomass\\figs'
input_file = 'EIA Dataset.csv'

d = pd.read_csv(filepath + '\\' + input_file)
d.drop(columns = ['Unnamed: 0', ], axis=1, inplace=True)

"""
Series considered for plotting:
'PRCE_NA_NA_NA_CL_MNMTH_NA_Y13DLRPTN.A', # Production price > Coal
'PRCE_SUP_NA_NA_NG_NA_L48_Y13DLRPMMBTU.A', # Production price > Natural gas, lower 48
'PRCE_NA_NA_NA_CR_WLHD_L48_Y13DLRPBBL.A', # Production price > Crude Oil
'PRCE_COMP_NA_NA_DSL_WHP_NA_Y13DLRPGLN.A', # Production price > Diesel
'PRCE_COMP_NA_NA_MGS_WHP_NA_Y13DLRPGLN.A', # Production price > Motor gasoline
'PRCE_COMP_NA_NA_JFL_WHP_NA_Y13DLRPGLN.A', # Production price > Jet fuel
'PRCE_COMP_RESD_NA_DSTL_WHP_NA_Y13DLRPGLN.A' # Production price > Residential distillate fuel oil

"""
# filtering data
fltr_series = 'PRCE_NA_NA_NA_CL_MNMTH_NA_Y13DLRPTN.A|\
PRCE_SUP_NA_NA_NG_NA_L48_Y13DLRPMMBTU.A|\
PRCE_NA_NA_NA_CR_WLHD_L48_Y13DLRPBBL.A|\
PRCE_COMP_NA_NA_DSL_WHP_NA_Y13DLRPGLN.A|\
PRCE_COMP_NA_NA_MGS_WHP_NA_Y13DLRPGLN.A|\
PRCE_COMP_NA_NA_JFL_WHP_NA_Y13DLRPGLN.A|\
PRCE_COMP_RESD_NA_DSTL_WHP_NA_Y13DLRPGLN.A'

d = d.query('`Series Id`.str.contains(@fltr_series)')
d.drop(columns = ['Series Id'], axis = 1, inplace = True)

# unit conversions
to_unit = 'MMBtu'
d[['numerator', 'denominator']] =  d.Unit.str.split('/', expand=True)
d['feedstock_perunit'] = d['Subcategory 1'] + '_' + d['denominator'] + '_per_' + to_unit

d['Value'] = d['Value'] * [ut.unit_conv(x) for x in d['feedstock_perunit']]
d['Unit'] = d['numerator'] + '/' + to_unit
d.drop(columns = ['numerator', "denominator", 'feedstock_perunit'], inplace = True)

# filter out Coal
d = d.loc[d['Subcategory 1'] != "Coal", ]

# plotting

custom_params = {"axes.spines.right": False, "axes.spines.top": False}
sns.set_theme(style="ticks", rc=custom_params)

g = sns.FacetGrid(d, col = "AEO Case", hue = "Subcategory 1",
                  col_wrap = 3, height = 3)
g.map(sns.lineplot, "Date", "Value")
g.set_titles("{col_name}")
g.set_axis_labels("Year", "2020 $/MMBtu")
g.set(xticks=range(min(d['Date']), max(d['Date'])+10, 10))
g.add_legend(title = "Fuel types")
g.figure.savefig(figpath+'\\'+'plot_line_EIA_price.jpg', dpi=400)