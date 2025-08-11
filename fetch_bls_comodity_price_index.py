# -*- coding: utf-8 -*-
"""
Copyright © 2025, UChicago Argonne, LLC
The full description is available in the LICENSE file at location:
    https://github.com/Saurajyoti/BestBiomassUse/blob/master/LICENSE

@Project: Best Use of Biomass
@Title: Fetch commodity price index data from BLS website
@Authors: Saurajyoti Kar
@Contact: skar@anl.gov
@Affiliation: Argonne National Laboratory 


Created on Thu Aug 31 16:24:34 2023

@author: skar
"""

import requests
from bs4 import BeautifulSoup

def recursiveUrl(url, link, depth):
    if depth == 5:
        return url
    else:
        print(link['href'])
        page = requests.get(url + link['href'])
        soup = BeautifulSoup(page.text, 'html.parser')
        newlink = soup.find('a')
        if len(newlink) == 0:
            return link
        else:
            return link, recursiveUrl(url, newlink, depth + 1)

def getLinks(url):
    page = requests.get(url)
    soup = BeautifulSoup(page.text, 'html.parser')
    links = soup.find_all('a')
    for link in links:
        links.append(recursiveUrl(url, link, 0))
    return links

links = getLinks("http://francaisauthentique.libsyn.com/")
print(links)
