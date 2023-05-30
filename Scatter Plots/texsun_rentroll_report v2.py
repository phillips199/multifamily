import pandas as pd
from pandas import ExcelWriter
import datetime
import numpy as np
import os
import matplotlib.pyplot as plt
import matplotlib.dates as dates
import tkinter
from tkinter import filedialog

print("Please select the key stats file that you want to analyze.")
root = tkinter.Tk()
root.update()
key_stats_file_path = filedialog.askopenfilename()
print(key_stats_file_path)
key_stats_xls = pd.ExcelFile(key_stats_file_path)
key_stats_df = pd.read_excel(key_stats_xls)
key_stats_df = key_stats_df.set_index('Bldg-Unit', drop=False)

print("Please select the unit types file that you want to analyze.")
root = tkinter.Tk()
root.update()
unit_types_file_path = filedialog.askopenfilename()
print(unit_types_file_path)
unit_types_xls = pd.ExcelFile(unit_types_file_path)
unit_types_df = pd.read_excel(unit_types_xls)
unit_types_df = unit_types_df.set_index('Bldg-Unit', drop=False)

time = datetime.datetime.now().strftime(" %Y-%m-%dT%H-%M-%S")

key_stats_new_df = key_stats_df.join(unit_types_df, lsuffix='_keystats', rsuffix='_types')

path, filename = os.path.split(key_stats_file_path)
name, ext = os.path.splitext(filename)

writer = ExcelWriter(path + name + "_reno" + ext)
key_stats_new_df.to_excel(writer, "Sheet1", index=False)
writer.save()


min_date = datetime.datetime(2020,7,1)
max_date = datetime.datetime(2023,5,31)

property_dict = {1: {'name': 'Bridge Hollow', 'date': datetime.datetime(2022,11,1)},
                2: {'name': 'Woodstone', 'date': datetime.datetime(2022,11,1)},
                3: {'name': 'Heights On Perrin', 'date': datetime.datetime(2022,10,1)},
                4: {'name': 'San Mateo', 'date': datetime.datetime(2022,10,1)},
                5: {'name': 'Trailside', 'date': datetime.datetime(2022,12,1)}}
cw_underwriting = {}

choice = int(input("""
Choose one of the following:
    1   Bridge Hollow
    2   Woodstone
    3   Heights On Perrin
    4   San Mateo
    5   Trailside
"""))

lease_xls = pd.ExcelFile('/Users/mikephillips/My Drive/Cottonwood/C. Equity/1. Project Documents/Leases/Lease Underwriting.xlsx')
for lease_sheet in lease_xls.sheet_names:
    lease_df = pd.read_excel(lease_xls, lease_sheet)
    counter = 0
    lease_details = {}
    for uw_type in lease_df["Type"]:
        lease_details[uw_type] = [lease_df["Before"][counter],lease_df["After"][counter]]
        counter += 1
    cw_underwriting[lease_sheet] = lease_details


for unit_type in key_stats_new_df["Unit Type"].unique():
    current_key_stats = key_stats_new_df[(key_stats_new_df["Unit Type"] == unit_type) &
                    (key_stats_new_df["Lease Start"] > min_date) &
                    (key_stats_new_df["Scheduled Rent"] > 0)]

    x = current_key_stats["Lease Start"].values
    y = current_key_stats["Scheduled Rent"].values
    plt.figure(figsize=(10,8))
    plt.subplots_adjust(top=0.9, bottom=0.25, left=0.2)
    ax = plt.axes()
    ax.scatter(x,y)
    plt.plot([property_dict[choice]['date'],property_dict[choice]['date']],[600,1800],color='grey', linestyle='dashed')
    plt.plot([min_date, property_dict[choice]['date']],[cw_underwriting[property_dict[choice]['name']][unit_type][0],
                                                        cw_underwriting[property_dict[choice]['name']][unit_type][0]],
             color='blue', linestyle='dashed')
    plt.plot([property_dict[choice]['date'], max_date],[cw_underwriting[property_dict[choice]['name']][unit_type][1],
                                                        cw_underwriting[property_dict[choice]['name']][unit_type][1]],
             color='green', linestyle='dashed')
    ax.set_title(property_dict[choice]['name'] + " - " + unit_type + "\n", fontsize=20)
    ax.set_ylabel('Rates\n',fontsize=20)
    ax.set_ylim(600, 1800)
    ax.set_xlabel('\nLease Dates',fontsize=20)
    ax.tick_params(axis='both', labelsize=20)
    ax.tick_params('x', labelrotation=30)
    ax.set_xlim(min_date, max_date)

    # Get values for the trend line analysis
    if len(x) > 1:
        x_num = dates.date2num(list(x))
        trend = np.polyfit(x_num, list(y), 1)
        fit = np.poly1d(trend)
        plt.plot(list(x), fit(list(x_num)), "r--")

    plt.savefig(property_dict[choice]['name'] + " - " + unit_type + time +".png")
    plt.close()
