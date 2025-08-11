# -*- coding: utf-8 -*-
"""
Copyright © 2025, UChicago Argonne, LLC
The full description is available in the LICENSE file at location:
    https://github.com/Saurajyoti/BestBiomassUse/blob/master/LICENSE

@Project: Best Use of Biomass
@Title: Script to call data processing scripts, process data, perform calculations, and save output files
@Authors: Saurajyoti Kar
@Contact: skar@anl.gov
@Affiliation: Argonne National Laboratory

Created on Mon Jan 23 13:02:25 2023

"""

#%%
# Paths and Filenames

data_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - Best use of biomass/data/interim'
fig_path_prefix = 'C:/Users/skar/Box/saura_self/Proj - Best use of biomass/figs'

f_mfsp_itemized = 'mfsp_itemized.csv'
f_mfsp_agg = 'mfsp_agg.csv'
f_lca_itemized = 'lca_itemized.csv'
f_lca_agg = 'lca_agg.csv'
f_mac = 'mac.csv'

f_pathway_mac = 'mac_for_pathways.xlsx'

#%%
# Libraries
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt
#import os
#from datetime import datetime

#%%
# Load data tables

mfsp_agg = pd.read_csv(data_path_prefix + '/' + f_mfsp_agg, index_col=0)
lca_agg = pd.read_csv(data_path_prefix + '/' + f_lca_agg, index_col=0)
mac = pd.read_csv(data_path_prefix + '/' + f_mac, index_col=0)

#%%

# Get the full list of styles
# sns.axes_style()

sns.set_style("whitegrid")


# LCA (kgCO2e/MJ) vs pathways
lca_agg['ID'] = lca_agg.index

plt.figure(figsize=(12, 6))
g = sns.barplot(
    data=lca_agg,
    x='Case/Scenario', y='Total LCA',
    palette='dark') #, alpha=0.6, height=6
g.set_axis_labels("Pathways", "Emissions (g CO2 per MJ)", labelpad=10)
g.set_xticklabels(g.get_xticklabels(), rotation=45, horizontalalignment='right')
g.add_legend(title="")
g.figure.set_size_inches(6.5, 4.5)
g.figure.savefig(fig_path_prefix+'\\'+'plot_LCA.jpg', dpi=400)


# TEA ($/GGE) vs pathways
mfsp_agg['ID'] = mfsp_agg.index
g = sns.catplot(
    data=mfsp_agg, kind='bar',
    x='ID', y='MFSP replacing fuel',
    palette='dark') #, alpha=0.6, height=6
g.set_axis_labels("Pathways", "MFSP ($ per GGE)")
g.add_legend(title="")
g.figure.set_size_inches(6.5, 4.5)
g.figure.savefig(fig_path_prefix+'\\'+'plot_MFSP.jpg', dpi=400)


# MAC ($/kgCO2e) vs pathways
mac['ID'] = mac.index
g = sns.catplot(
    data=mac, kind='bar',
    x='ID', y='MAC_calculated',
    palette='dark') #, alpha=0.6, height=6
g.set_axis_labels("Pathways", "Cost of CO2 Abated ($ per MT CO2)")
g.add_legend(title="")
g.figure.set_size_inches(6.5, 4.5)
g.figure.savefig(fig_path_prefix+'\\'+'plot_MAC.jpg', dpi=400)


# LCA vs TEA (kg CO2e/ MJ vs $ per GGE)
lca_tea = pd.merge(mfsp_agg, lca_agg, how='left',
                   on=['Case/Scenario']).reset_index(drop=True)
lca_tea["Total LCA (kg/GGE)"] = lca_tea['Total LCA'] / 1000
#consider unique pathways
lca_tea = lca_tea[['Case/Scenario', 'MFSP replacing fuel', 'Total LCA (kg/GGE)']].drop_duplicates().reset_index(drop=True)
lca_tea['ID'] = lca_tea.index


g = sns.relplot(
    data=lca_tea,
    x="MFSP replacing fuel", y="Total LCA (kg/GGE)",
    #hue="year", size="mass",
    #palette=cmap, sizes=(10, 200),
)
#fig, ax = plt.subplots()
for (a, b, c) in zip(lca_tea['MFSP replacing fuel'], lca_tea['Total LCA (kg/GGE)'], lca_tea['ID']):
    g.ax.text(a+.1, b+0.2, c)

#g.set(xscale="log", yscale="log")
g.set_axis_labels("MFSP ($ per GGE)", "Emissions (kg CO2 per GGE)")
g.ax.xaxis.grid(True, "minor", linewidth=.25)
g.ax.yaxis.grid(True, "minor", linewidth=.25)
g.despine(left=True, bottom=True)
g.figure.set_size_inches(6.5, 4.5)
g.figure.savefig(fig_path_prefix+'\\'+'plot_LCA_MFSP.jpg', dpi=400)

# Four quard plot: ratio of ghg of alt fuels and the conv fuels vs. ratio of MFSP of alt fuels and conventional fuels


#%%

# Pathway level analysis - graphs

# perc CI reduction vs MAC

fig_pathway_prcCI_vs_MAC = 'pathway_prcCI_vs_MAC'


#%%

# carbon intensity variability analysis

