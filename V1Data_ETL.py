# -*- coding: utf-8 -*-
"""
Created on Tue Oct 11 11:27:39 2022

@author: Eamonn Nemeh
"""

import pandas as pd
import glob

#Loading Files
path = '\\\Computer\\documents\Data'
df = pd.DataFrame()
files = glob.glob(path + r'\\*V1 Data.xlsx')

for f in files:
    df1 = pd.read_excel(f, sheet_name='V1 Data', skiprows=[2], header = [0,1])
    df1['Filename'] = f
    df = df.append(df1, ignore_index=True)
df.dropna(axis = 1,how = 'all', inplace = True)
df.columns = [' -- '.join(col).rstrip('_') for col in df.columns.values]

#Creating list of columns we do not want transposed
a = [col for col in df if col.startswith('Overall')]
b = [col for col in df if col.startswith('Assessment')]
c = [col for col in df if col.startswith('Filename')]
d = a+b+c

#Transposing the dataset to get Assessed and Result
df=df.melt(id_vars=d, var_name="Assessed",value_name="Result")

#Drop nas and fixing line breaks
df['Assessed'] = df['Assessed'].str.replace('\n', ' ')
df.dropna(subset=['Result'], inplace=True)

# #Separating one column to create Category and Assessed
new = df["Assessed"].str.split(" -- ", n=1, expand = True)
df.drop(columns =["Assessed"], inplace = True)
df["Category"] = new[0]
df["Assessed"] = new[1]

# Fixing Column names and editing formats
df.columns = df.columns.str.replace("Overall -- ", "")
df.columns = df.columns.str.replace("Assessment Information -- ", "")
df.columns = df.columns.str.replace("Filename -- ", "Filename")
df['Visit Date'] = df['Visit Date'].dt.strftime('%d/%m/%Y')
df['Assessed'] = df['Assessed'].str.title()
df['Category'] = df['Category'].str.title()

#Storing Files
df.to_csv(r'\\Computer\e$\QVAPPS\NewProjects\2.DataStore\CSV\V1Table.csv', index=False)