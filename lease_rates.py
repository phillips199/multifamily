import pandas as pd
import numpy as np
from pandas import ExcelWriter
from datetime import datetime, timedelta
import tkinter
from tkinter import filedialog
import matplotlib.pyplot as plt
import matplotlib.dates as dates

start_date = datetime(2022, 6, 1)
end_date = datetime(2023, 3, 31)
sheet = "Quinn"
chart_folder = "/Users/mikephillips/My Drive/Cottonwood/C. Equity/3. Work Product/Leasing Charts/"

print("Please select the key stats file that you want to analyze.")
root = tkinter.Tk()
root.update()
key_stats_file_path = filedialog.askopenfilename()
key_stats_xls = pd.ExcelFile(key_stats_file_path)
key_stats_df = pd.read_excel(key_stats_xls, "Sheet1")
key_stats_df = key_stats_df.set_index('Bldg-Unit')

print("Please select the unit types file that you want to use.")
unit_types_file_path = filedialog.askopenfilename()
unit_types_xls = pd.ExcelFile(unit_types_file_path)
unit_types_df = pd.read_excel(unit_types_xls, "Sheet1")
unit_types_df = unit_types_df.set_index('Bldg-Unit')

key_stats_new_df = key_stats_df.join(unit_types_df, lsuffix='_keystats', rsuffix='_types')

writer = ExcelWriter("key_lease_rates.xlsx")
key_stats_new_df.to_excel(writer, "Sheet1", index=True)
writer.save()

for unit_type in key_stats_new_df['Unit Type'].unique():

    print("Unit Type: " + unit_type)
    current_key_stats = key_stats_new_df[(key_stats_new_df["Unit Type"] == unit_type) &
                    (key_stats_new_df["Lease Start"] >= start_date)]

    x = current_key_stats["Lease Start"].values
    y = current_key_stats["Scheduled Rent"].values
    plt.figure(figsize=(10, 8))
    plt.subplots_adjust(top=0.9, bottom=0.25, left=0.2)
    plt.yticks(np.arange(0, 2100, 50))
    ax = plt.axes()
    ax.scatter(x, y)
    ax.set_title(sheet + " - " + unit_type + "\n", fontsize=10)
    ax.set_ylabel('Rates\n', fontsize=5)
    ax.set_ylim(0, 2100)
    ax.set_xlabel('\nLease Dates', fontsize=5)
    ax.tick_params(axis='both', labelsize=5)
    ax.tick_params('x', labelrotation=30)
    ax.set_xlim(start_date, end_date)

    # Get values for the trend line analysis
    if len(x) > 1:
        x_num = dates.date2num(list(x))
        trend = np.polyfit(x_num, list(y), 1)
        fit = np.poly1d(trend)
        plt.plot(list(x), fit(list(x_num)), "r--")

    plt.savefig(chart_folder + sheet + " - " + unit_type + ".png")
    plt.close()
