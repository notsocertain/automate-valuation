#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jun 12 21:40:34 2024

@author: notsocertainwind
"""

import pandas as pd
import numpy as np


import pandas as pd

# JSON-like data
data = {
    "name_client": ["Mr. Niraj Bhandari"],
    "address_client": ["Kanakai Ward No.04, Ilam"],
    "owner": ["Mr. Devi Prasad Bhandari"],
    "plot": ["111,222"],
    "area": ["224,553"],
    "address": ["kapan,dang"],
    "ctzno": ["641/3269"],
    "issuedate": ["2036/10/16 AD"],
    "address_owner": ["Kanakai Ward No.04, Ilam"],
    "relationship": ["Grandfather"],
    "final": ["4000000"]
}

# Split 'plot' and 'area' fields into lists
plots = data["plot"][0].split(',')
areas = data["area"][0].split(',')
address = data["address"][0].split(',')

# Prepare data for DataFrame
df_data = {
    "name_client": [data["name_client"][0]] * len(plots),
    "address_client": [data["address_client"][0]] * len(plots),
    "owner": [data["owner"][0]] * len(plots),
    "plot": plots,
    "area": areas,
    "address": address,
    "ctzno": [data["ctzno"][0]] * len(plots),
    "issuedate": [data["issuedate"][0]] * len(plots),
    "address_owner": [data["address_owner"][0]] * len(plots),
    "relationship": [data["relationship"][0]] * len(plots),
    "final": [data["final"][0]] * len(plots)
}

# Create DataFrame
df = pd.DataFrame(df_data)

list(df)

df.head()


df.to_csv('second_input.csv')
