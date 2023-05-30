import pandas as pd
from pandas import ExcelWriter
import os
import numpy as np
import tkinter
from tkinter import filedialog

print("Please select the rent roll folder that you want to combine.")
root = tkinter.Tk()
root.update()
target_dir = filedialog.askdirectory()
key_stats = pd.DataFrame()

emo_ask = input("Do these files have an Expected Move Out column (Y/N)?")

for item in os.listdir(target_dir):
    if item[0] == "~" or item == "key_stats.xlsx" or item == "Unit Types.xlsx":
        continue
    abs_item_name = os.path.join(target_dir, item)

    filename, ext = os.path.splitext(item)
    rr_month = filename[-10:]

    rentroll_xls = pd.ExcelFile(abs_item_name)
    sheet = rentroll_xls.sheet_names[0]
    rentroll_df = pd.read_excel(rentroll_xls, sheet)

    print(rr_month)
    print(type(rentroll_df['Scheduled Rent'].values[0]))

    rentroll_df['Bldg-Unit'] = rentroll_df['Bldg-Unit'].astype(str)
    rentroll_df['Scheduled Rent'] = rentroll_df['Scheduled Rent'].replace(',', '', regex=True).astype(float)
    # rentroll_df['Scheduled Rent'] = pd.to_numeric(rentroll_df['Scheduled Rent'], errors='coerce')
    rentroll_df['Lease Start'] = pd.to_datetime(rentroll_df['Lease Start'])
    rentroll_df['Lease Expiration'] = pd.to_datetime(rentroll_df['Lease Expiration'])

    if emo_ask == "Y":
        rentroll_df['Move-Out'] = pd.to_datetime(rentroll_df['Move-Out'])
    else:
        rentroll_df['Move-Out'] = np.NaN

    rentroll_df['Month'] = rr_month

    key_stats = key_stats.append(rentroll_df[['Bldg-Unit', 'Resident', 'Lease Start',
                                              'Lease Expiration', 'Move-Out', 'Scheduled Rent', 'Month']],
                                 ignore_index=True)

key_stats['Resident'] = key_stats['Resident'].apply(lambda x: 'Vacant' if x.find('Vacant') > -1 else x)

key_stats.sort_values(by = ['Bldg-Unit','Month'], ascending = [True, True], inplace=True)

export_file = os.path.join(target_dir, "key_stats.xlsx")
writer = ExcelWriter(export_file)
key_stats.to_excel(writer, "Sheet1", index=False)
writer.save()
