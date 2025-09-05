import pandas as pd
import openpyxl
import re
import xlsx_generator

def contains_numbers(string):
    """Check if the string contains any numbers."""
    return bool(re.search(r'\d', string))

df = pd.read_excel("SA727 Wage Details Report - June 2025.xlsx", skiprows=4, skipfooter=7)
df.drop(['AC', 'Account'], axis=1, inplace=True)

notna_mask = df['Name'].notna()
contains_numbers_mask = df.loc[notna_mask, 'Name'].apply(contains_numbers)
mask = notna_mask & contains_numbers_mask

df.loc[mask, 'Name'] = df.loc[mask, 'Name'].str.split(' ').str[0]
df.drop(df[df['Type'] == "Salaries and Wages"].index, inplace=True)

xlsx_generator.xlsx_gen(df)
