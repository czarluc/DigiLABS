import numpy as np
import pandas as pd
import glob
import openpyxl
import datetime as dt
import xlsxwriter

import requests
import json
import time
import os
from io import BytesIO

from word2number import w2n

sht1_dataset = pd.read_excel("SocMed_Geographic.xlsx", header=1, sheet_name="Sheet1")

sht1_dataset.rename(columns={sht1_dataset.columns[0]:'Social Media'}, inplace=True)
sht1_dataset = sht1_dataset.set_index(sht1_dataset.iloc[:, 0].name)
sht1_dataset = sht1_dataset.fillna("zero")
sht1_dataset.astype(str).apply(', '.join, axis=1)

for column in sht1_dataset.columns:
    colname_users = column + " users"
    colname_yr = column + " yr"

    sht1_dataset[colname_users] = sht1_dataset[column] 
    sht1_dataset[colname_users] = np.where(sht1_dataset[column].str.contains("zero"), "0", sht1_dataset[colname_users])
    sht1_dataset[colname_users] = np.where(sht1_dataset[column].str.contains("zero"), sht1_dataset[colname_users],
        (sht1_dataset[colname_users].str.split("(")).str[0])
    sht1_dataset[colname_users] = sht1_dataset[colname_users].str.replace(",", "")

    sht1_dataset[colname_yr] = sht1_dataset[column] 
    sht1_dataset[colname_yr] = np.where(sht1_dataset[column].str.contains("zero"), "0", sht1_dataset[colname_yr])
    sht1_dataset[colname_yr] = np.where(sht1_dataset[column].str.contains("zero"), sht1_dataset[colname_yr],
        (((sht1_dataset[colname_yr].str.split("(")).str[1]).str.split(",")).str[0] )

# sheet 2

sht2_dataset = pd.read_excel("SocMed_Geographic.xlsx", header=1, sheet_name="Sheet2")
sht2_dataset.rename(columns={sht2_dataset.columns[0]:'Social Media'}, inplace=True)
sht2_dataset = sht2_dataset.set_index(sht2_dataset.iloc[:, 0].name)
sht2_dataset = sht2_dataset.fillna("zero")
sht2_dataset.astype(str).apply(', '.join, axis=1)

for column in sht2_dataset.columns:
    colname_dur = column + " dur"

    colname_hr = column + " hr"
    colname_hr_letter = column + " hr letter"

    colname_min = column + " min"

    colname_yr = column + " yr"

    sht2_dataset[colname_dur] = sht2_dataset[column]
    sht2_dataset[colname_dur] = np.where(sht2_dataset[column].str.contains("zero"), "0", sht2_dataset[colname_dur])
    sht2_dataset[colname_dur] = np.where(sht2_dataset[column].str.contains("zero"), sht2_dataset[colname_dur],
        (sht2_dataset[colname_dur].str.split("(")).str[0])
    
    sht2_dataset[colname_hr] = np.where(sht2_dataset[column].str.contains("hour"), ((sht2_dataset[column].str.split("hour")).str[0]).str.strip(), "0")
    sht2_dataset[colname_hr_letter] = np.where( sht2_dataset[colname_hr].str.isalpha(), sht2_dataset[colname_hr], "zero")
    sht2_dataset[colname_hr] = np.where( sht2_dataset[colname_hr].str.isalpha(), "0", sht2_dataset[colname_hr])
    sht2_dataset[colname_hr_letter] = sht2_dataset[colname_hr_letter].apply(w2n.word_to_num)
    sht2_dataset[colname_hr] = sht2_dataset[colname_hr].astype(float)
    sht2_dataset[colname_hr_letter] = sht2_dataset[colname_hr_letter].astype(float)
    sht2_dataset[colname_hr] = (sht2_dataset[colname_hr]) + (sht2_dataset[colname_hr_letter])
    del sht2_dataset[colname_hr_letter]

    sht2_dataset[colname_min] = sht2_dataset[column]
    sht2_dataset[colname_min] = np.where(sht2_dataset[column].str.contains("zero"), "0", sht2_dataset[colname_min])
    sht2_dataset[colname_min] = np.where(sht2_dataset[column].str.contains("min"), ((sht2_dataset[column].str.split("min")).str[0]).str.strip(), "0")
    sht2_dataset[colname_min] = np.where(sht2_dataset[colname_min].str.contains("hour"), ((sht2_dataset[colname_min].str.split("hour")).str[1]), sht2_dataset[colname_min])
    sht2_dataset[colname_min] = np.where(sht2_dataset[column].str.contains("hour"), sht2_dataset[colname_min], 
        np.where(sht2_dataset[column].str.contains("min"), sht2_dataset[colname_min], sht2_dataset[colname_dur]) 
        ) # cells with no units are considered as minutes
    sht2_dataset[colname_min] = sht2_dataset[colname_min].astype(float)

    sht2_dataset[colname_dur] = (sht2_dataset[colname_hr] * 60) + sht2_dataset[colname_min]
    del sht2_dataset[colname_hr]
    del sht2_dataset[colname_min]

    sht2_dataset[colname_yr] = sht2_dataset[column] 
    sht2_dataset[colname_yr] = np.where(sht2_dataset[column].str.contains("zero"), "0", sht2_dataset[colname_yr])
    sht2_dataset[colname_yr] = np.where(sht2_dataset[column].str.contains("zero"), sht2_dataset[colname_yr],
        (((sht2_dataset[colname_yr].str.split("(")).str[1]).str.split(",")).str[0] )
    sht2_dataset[colname_yr] = sht2_dataset[colname_yr].fillna("2020") # cells with blank years are put to 2020 yr 

# population dataset

population_dataset = pd.read_csv("country_population.csv", header=2, usecols=["Country Name", "2018", "2019", "2020"])
population_dataset.rename(columns={population_dataset.columns[0]:'Country'}, inplace=True)
population_dataset = population_dataset.set_index(population_dataset.iloc[:, 0].name)

print(population_dataset.head(100))



