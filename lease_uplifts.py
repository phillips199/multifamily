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
sheet = "HOP"
chart_folder = "/Users/mikephillips/My Drive/Cottonwood/C. Equity/3. Work Product/Leasing Charts/"

print("Please select the key stats vacancies file that you want to analyze.")
root = tkinter.Tk()
root.update()
key_stats_vacancies_file_path = filedialog.askopenfilename()
key_stats_vacancies_xls = pd.ExcelFile(key_stats_vacancies_file_path)
key_stats_vacancies_df = pd.read_excel(key_stats_vacancies_xls, "Sheet1")

for unit_type in key_stats_vacancies_df['Unit Type'].unique():

    print("Unit Type: " + unit_type)
    current_key_stats = key_stats_vacancies_df[(key_stats_vacancies_df["Unit Type"] == unit_type) &
                    (key_stats_vacancies_df["End Vacant Date"] >= start_date) &
                    (key_stats_vacancies_df["Prior Rent"] > 0) &
                    (key_stats_vacancies_df["Next Rent"] > 0)]

    x = current_key_stats["End Vacant Date"].values
    y = current_key_stats["Adjusted Uplift"].values
    plt.figure(figsize=(10, 8))
    plt.subplots_adjust(top=0.9, bottom=0.25, left=0.2)
    ax = plt.axes()
    ax.scatter(x, y)
    ax.set_title(sheet + " - " + unit_type + "\n", fontsize=10)
    ax.set_ylabel('Uplifts\n', fontsize=5)
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

    plt.savefig(chart_folder + sheet + " - Uplifts - " + unit_type + ".png")
    plt.close()
