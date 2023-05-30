import pandas as pd
from pandas import ExcelWriter
import os
import datetime
import numpy as np
import tkinter
from tkinter import filedialog

print("Please select the rent roll folder that you want to combine.")
root = tkinter.Tk()
root.update()
target_dir = filedialog.askdirectory()
key_stats = pd.DataFrame()
print(target_dir)

time = datetime.datetime.now().strftime(" %Y-%m-%dT%H-%M-%S")

# emo_ask = input("Do these files have a Move Out column (Y/N)?")

# key_stats_cols = ['Bldg-Unit', 'Resident', 'Move-In', 'Lease Start', 'Lease End',
#                   'Move-Out', 'Scheduled Rent', 'Market Rent', 'Month']

# if input("Do these files have a Balance column (Y/N)?") == "Y":
#     key_stats_cols.append('Balance')

for item in os.listdir(target_dir):
    if item[0] == "~":
        continue
    print(item)
    abs_item_name = os.path.join(target_dir, item)
    filename, ext = os.path.splitext(item)
    rr_month = filename[-10:]

    rentroll_xls = pd.ExcelFile(abs_item_name)
    rentroll_df = pd.read_excel(rentroll_xls)

    rentroll_df['Move-In'] = pd.to_datetime(rentroll_df['Move-In'])
    rentroll_df['Lease Start'] = pd.to_datetime(rentroll_df['Lease Start'])
    rentroll_df['Lease End'] = pd.to_datetime(rentroll_df['Lease End'])

    if "Expected Move-Out" in rentroll_df.columns:
        rentroll_df = rentroll_df.rename(columns={"Expected Move-Out": "Move-Out"})
        rentroll_df['Move-Out'] = pd.to_datetime(rentroll_df['Move-Out'])
    else:
        rentroll_df['Move-Out'] = np.NaN

    if "Deposit Held" in rentroll_df.columns:
        if rentroll_df['Deposit Held'].dtype not in ['float64','int64']:
            rentroll_df['Deposit Held'] = rentroll_df['Deposit Held'].str.replace(',', '')
            rentroll_df['Deposit Held'] = rentroll_df['Deposit Held'].astype(float)

    rentroll_df['Month'] = rr_month

    if "Resident" not in rentroll_df.columns:
        rentroll_df["Resident"] = rentroll_df["Unit Status"]

    key_stats = key_stats.append(rentroll_df, ignore_index=True)

key_stats.sort_values(by=['Bldg-Unit', 'Month'], ascending=[True, True], inplace=True)

export_file = os.path.join(target_dir, "key_stats" + time + ".xlsx")
writer = ExcelWriter(export_file)
key_stats.to_excel(writer, "Sheet1", index=False)
writer.save()
