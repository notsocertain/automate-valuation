#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Wed Jun 12 21:45:08 2024

@author: notsocertainwind
"""

import pandas as pd
import numpy as np
class AreaConverter:
    def __init__(self):
        pass

    @staticmethod
    def sqft_to_rapd(sqft):
        ropani = int(sqft // 5476)
        remaining_sqft = sqft % 5476
        aana = int(remaining_sqft // 342.25)
        remaining_sqft %= 342.25
        paisa = int(remaining_sqft // 85.56)
        remaining_sqft %= 85.56
        daam = round(remaining_sqft / 21.39, 3)
        return (ropani, aana, paisa, daam)

    @staticmethod
    def sqm_to_rapd(sqm):
        sqft = sqm / 0.092903
        return AreaConverter.sqft_to_rapd(sqft)

    @staticmethod
    def rapd_to_sqft(ropani, aana, paisa, daam):
        sqft = (ropani * 5476) + (aana * 342.25) + (paisa * 85.56) + (daam * 21.39)
        sqft = round(sqft, 2)
        return sqft

    @staticmethod
    def rapd_to_sqm(ropani, aana, paisa, daam):
        sqft = AreaConverter.rapd_to_sqft(ropani, aana, paisa, daam)
        return AreaConverter.sqft_to_sqm(sqft)

    @staticmethod
    def sqm_to_sqft(sqm):
        sqft = sqm / 0.092903
        return round(sqft, 2)

    @staticmethod
    def sqft_to_sqm(sqft):
        sqm = sqft * 0.092903
        return round(sqm, 2)


class AreaHelper:
    def __init__(self):
        self.sqft_to_sqm = AreaConverter.sqft_to_sqm
        self.sqm_to_sqft = AreaConverter.sqm_to_sqft
        self.sqft_to_rapd = AreaConverter.sqft_to_rapd
        self.rapd_to_sqft = AreaConverter.rapd_to_sqft
        self.rapd_to_sqm = AreaConverter.rapd_to_sqm
        self.sqm_to_rapd =AreaConverter.sqm_to_rapd
        

    def ft2_to_m2(self, ft2):
        if pd.notnull(ft2):
            return self.sqft_to_sqm(ft2)
        else:
            return np.nan

    def m2_to_ft2(self, m2):
        if pd.notnull(m2):
            return self.sqm_to_sqft(m2)
        else:
            return np.nan

    def ft2_to_rapd(self, ft2):
        if pd.notnull(ft2):
            ropani, aana, paisa, daam = self.sqft_to_rapd(ft2)
            return f"{ropani}-{aana}-{paisa}-{daam}"
        else:
            return np.nan

    def rapd_to_ft2(self, rapd):
        if pd.notnull(rapd):
            ropani, aana, paisa, daam = rapd
            return self.rapd_to_sqft(ropani, aana, paisa, daam)
        else:
            return np.nan

    def rapd_to_m2(self, rapd):
        if pd.notnull(rapd):
            ropani, aana, paisa, daam = rapd
            return self.rapd_to_sqm(ropani, aana, paisa, daam)
        else:
            return np.nan

    def m2_to_rapd(self, m2):
        if pd.notnull(m2):
            ft2 = self.sqm_to_sqft(m2)
            return self.ft2_to_rapd(ft2)
        else:
            return np.nan

    def extract_rapd(self, rapd_str):
        if pd.notnull(rapd_str):
            parts = rapd_str.split("-")
            return tuple(float(part) for part in parts)
        else:
            return np.nan


# # Example usage
# area_helper = AreaHelper()

# ft2_value = 10000
# m2_value = 500
# rapd_str = "1-5-2-3"

# # Convert ft2 to m2
# print(f"{ft2_value} square feet is {area_helper.ft2_to_m2(ft2_value)} square meters")
