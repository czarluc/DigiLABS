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
    colname_users = column.strip() + " users"
    colname_yr = column.strip() + " yr"

    sht1_dataset[colname_users] = sht1_dataset[column] 
    sht1_dataset[colname_users] = np.where(sht1_dataset[column].str.contains("zero"), "0", sht1_dataset[colname_users])
    sht1_dataset[colname_users] = np.where(sht1_dataset[column].str.contains("zero"), sht1_dataset[colname_users],
        (sht1_dataset[colname_users].str.split("(")).str[0])
    sht1_dataset[colname_users] = (sht1_dataset[colname_users].str.replace(",", "")).str.strip()
    sht1_dataset[colname_users] = sht1_dataset[colname_users].str.replace('\W', '', regex=True)
    sht1_dataset[colname_users] = (sht1_dataset[colname_users].str.replace(" ", "")).str.strip()
    sht1_dataset[colname_users] = sht1_dataset[colname_users].astype(int)

    sht1_dataset[colname_yr] = sht1_dataset[column] 
    sht1_dataset[colname_yr] = np.where(sht1_dataset[column].str.contains("zero"), "0", sht1_dataset[colname_yr])
    sht1_dataset[colname_yr] = np.where(sht1_dataset[column].str.contains("zero"), sht1_dataset[colname_yr],
        (((sht1_dataset[colname_yr].str.split("(")).str[1]).str.split(",")).str[0] )
    sht1_dataset[colname_yr] = sht1_dataset[colname_yr].fillna("2020") # cells with blank years are put to 2020 yr
    sht1_dataset[colname_yr] = sht1_dataset[colname_yr].astype(int)
    
    del sht1_dataset[column]

# sheet 2

sht2_dataset = pd.read_excel("SocMed_Geographic.xlsx", header=1, sheet_name="Sheet2")
sht2_dataset.rename(columns={sht2_dataset.columns[0]:'Social Media'}, inplace=True)
sht2_dataset = sht2_dataset.set_index(sht2_dataset.iloc[:, 0].name)
sht2_dataset = sht2_dataset.fillna("zero")
sht2_dataset.astype(str).apply(', '.join, axis=1)

for column in sht2_dataset.columns:
    colname_dur = column.strip() + " dur"

    colname_hr = column.strip() + " hr"
    colname_hr_letter = column.strip() + " hr letter"

    colname_min = column.strip() + " min"

    colname_yr = column.strip() + " yr"

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
    sht2_dataset[colname_dur] = sht2_dataset[colname_dur].astype(float)

    sht2_dataset[colname_yr] = sht2_dataset[column] 
    sht2_dataset[colname_yr] = np.where(sht2_dataset[column].str.contains("zero"), "0", sht2_dataset[colname_yr])
    sht2_dataset[colname_yr] = np.where(sht2_dataset[column].str.contains("zero"), sht2_dataset[colname_yr],
        (((sht2_dataset[colname_yr].str.split("(")).str[1]).str.split(",")).str[0] )
    sht2_dataset[colname_yr] = sht2_dataset[colname_yr].fillna("2020") # cells with blank years are put to 2020 yr
    sht2_dataset[colname_yr] = sht2_dataset[colname_yr].astype(int)

    del sht2_dataset[column]

# population dataset
# SOURCE: https://data.worldbank.org/indicator/SP.POP.TOTL?name_desc=false

population_dataset = pd.read_csv("country_population.csv", header=2, usecols=["Country Name", "2018", "2019", "2020"])
population_dataset.rename(columns={population_dataset.columns[0]:'Country'}, inplace=True)
population_dataset = population_dataset.set_index(population_dataset.iloc[:, 0].name)

print("#### ANSWERS #####")

# #1
# print(sht1_dataset.index.values.tolist())

selected_col = sht1_dataset.index.values.tolist()[0]

selected_index = sht1_dataset.index.get_loc(selected_col)
selected_media = sht1_dataset.iloc[[selected_index]]

print("Number of "+ selected_col +" users per country")
for column in selected_media.columns:
    if "users" in column:
        country = (column.replace("users", "")).strip()
        users = '{:,}'.format(selected_media[column].values[0])
        print("Users in " + country + ": " + str(users))

print()

# #2
columns = sht2_dataset.columns.values.tolist()
countries = []

for col in columns:
    if "dur" in col:
        country = (col.replace("dur", "")).strip()
        countries.append(country)

selected_col = "United States"
selected_media = sht2_dataset.loc[:, selected_col + " dur"]
selected_media_sorted = selected_media.sort_values(ascending=False)

print("Social Media Duration usage (in minutes) in " + selected_col)
for i, number in selected_media_sorted.iteritems():
    print(str(i) + ": " + str(number) + " minutes")

print()

# #3
selected_col = "United States"

selected_country = sht1_dataset.loc[:, selected_col + " users"]
# print(selected_country)

selected_country_yr = sht1_dataset.loc[:, selected_col + " yr"]
# print(selected_country_yr)

pop_index = population_dataset.index.get_loc(selected_col)
pop_country = population_dataset.iloc[[pop_index]]

pop_list = []
for i, number in selected_country_yr.iteritems():
    # print(number) # yr
    if number != 0:
        pop_list.append(pop_country[str(number)].values[0])
    else:
        pop_list.append(0)

print("Percentage of users per Social Media in " + selected_col)
counter = 0
for i, number in selected_country.iteritems():
    user_percentage = 0
    if number != 0:
        user_percentage = number / pop_list[counter] 
    user_percentage = "{0:.0%}".format(user_percentage)
    print(str(i) + ": " + str(user_percentage))
    counter += 1

print()