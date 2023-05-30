import pandas as pd
from pandas import ExcelWriter
from datetime import datetime, timedelta
import tkinter
from tkinter import filedialog

start_date = datetime(2022, 1, 1)
end_date = datetime(2023, 3, 31)
sheet = "HOP"

tda = 60      # trailing days average

print("Please select the key lease rates file that you want to analyze.")
root = tkinter.Tk()
root.update()
key_stats_file_path = filedialog.askopenfilename()
key_stats_xls = pd.ExcelFile(key_stats_file_path)
key_stats_df = pd.read_excel(key_stats_xls, "Sheet1")

trailing_avg = []
date = start_date
while date <= end_date:

    print(date)
    current_avg_rents = {}
    running_total_variances = 0
    running_total_units = 0

    current_avg_rents['Date'] = date
    for unit_type in key_stats_df['Unit Type'].unique():

        current_key_stats = key_stats_df[(key_stats_df["Unit Type"] == unit_type) &
                                         (key_stats_df["Scheduled Rent"] > 0) &
                                         (key_stats_df["Lease Start"] >= date + timedelta(days=-tda)) &
                                         (key_stats_df["Lease Start"] <= date)]

        current_key_stats = current_key_stats[['Bldg-Unit', 'Lease Start', 'Scheduled Rent', 'Unit Type']]
        if len(current_key_stats['Scheduled Rent'].values) > 0:

            current_key_stats.drop_duplicates(inplace=True)
            avg_rent = sum(current_key_stats['Scheduled Rent'].values) / len(current_key_stats['Scheduled Rent'].values)

            current_avg_rents[unit_type + ' ($)'] = avg_rent
            current_avg_rents[unit_type + ' (#)'] = len(current_key_stats['Scheduled Rent'].values)

            running_total_variances += avg_rent * len(current_key_stats['Scheduled Rent'].values)
            running_total_units += len(current_key_stats['Scheduled Rent'].values)

    current_avg_rents['Total ($)'] = running_total_variances / running_total_units
    current_avg_rents['Total (#)'] = running_total_units

    trailing_avg.append(current_avg_rents)
    date = date + timedelta(days=1)

variances_stats_df = pd.DataFrame.from_dict(trailing_avg)
writer = ExcelWriter(str(tda) + "-day_trailing_avg_rent.xlsx")
variances_stats_df.to_excel(writer, "Sheet1", index=False)
writer.save()
