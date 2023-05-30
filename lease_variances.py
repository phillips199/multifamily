import pandas as pd
from pandas import ExcelWriter
from datetime import datetime, timedelta
import tkinter
from tkinter import filedialog

start_date = datetime(2022, 1, 1)
end_date = datetime(2023, 3, 31)
sheet = "HOP"
days_trailing = 30

print("Please select the key lease rates file that you want to analyze.")
root = tkinter.Tk()
root.update()
key_stats_file_path = filedialog.askopenfilename()
key_stats_xls = pd.ExcelFile(key_stats_file_path)
key_stats_df = pd.read_excel(key_stats_xls, "Sheet1")

average_rents = {}
for unit_type in key_stats_df['Unit Type'].unique():

    current_key_stats = key_stats_df[(key_stats_df["Unit Type"] == unit_type) &
                                     (key_stats_df["Scheduled Rent"] > 0)]

    current_key_stats = current_key_stats[['Bldg-Unit', 'Lease Start', 'Scheduled Rent', 'Unit Type']]
    if len(current_key_stats['Scheduled Rent'].values) > 0:

        current_key_stats.drop_duplicates(inplace=True)
        avg_rent = sum(current_key_stats['Scheduled Rent'].values) / len(current_key_stats['Scheduled Rent'].values)

        average_rents[unit_type] = avg_rent

variances = []
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
                                         (key_stats_df["Lease Start"] >= date + timedelta(days=-days_trailing)) &
                                         (key_stats_df["Lease Start"] <= date)]

        current_key_stats = current_key_stats[['Bldg-Unit', 'Lease Start', 'Scheduled Rent', 'Unit Type']]
        if len(current_key_stats['Scheduled Rent'].values) > 0:

            current_key_stats.drop_duplicates(inplace=True)
            avg_rent = sum(current_key_stats['Scheduled Rent'].values) / len(current_key_stats['Scheduled Rent'].values)

            current_avg_rents[unit_type + ' (%)'] = avg_rent / average_rents[unit_type] - 1
            current_avg_rents[unit_type + ' (#)'] = len(current_key_stats['Scheduled Rent'].values)

            running_total_variances += (avg_rent / average_rents[unit_type] - 1) * \
                                       len(current_key_stats['Scheduled Rent'].values)
            running_total_units += len(current_key_stats['Scheduled Rent'].values)

    current_avg_rents['Total (%)'] = running_total_variances / running_total_units
    current_avg_rents['Total (#)'] = running_total_units

    if 'Total (%)' in current_avg_rents.keys() and 'Total (#)' not in current_avg_rents.keys():
        print('missing ' + 'Total (#) for ' + str(date))

    if 'Total (%)' not in current_avg_rents.keys() and 'Total (#)' in current_avg_rents.keys():
        print('missing ' + 'Total (%) for ' + str(date))

    variances.append(current_avg_rents)
    date = date + timedelta(days=1)

variances_stats_df = pd.DataFrame.from_dict(variances)
writer = ExcelWriter("variances_stats_by_day.xlsx")
variances_stats_df.to_excel(writer, "Sheet1", index=False)
writer.save()
