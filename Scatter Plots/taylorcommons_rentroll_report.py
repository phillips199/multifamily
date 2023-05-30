
import pandas as pd
import datetime
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as dates

min_date = datetime.datetime(2020, 7, 1)
max_date = datetime.datetime(2023, 4, 30)

chart_folder = "/Users/mikephillips/My Drive/Cottonwood/C. Equity/3. Work Product/Leasing Charts/"

rentroll_xls = pd.ExcelFile('/Users/mikephillips/My Drive/Downloads/Taylor Commons RR 2023.02.28.xlsx')
sheet = "Sheet1"
key_stats = pd.read_excel(rentroll_xls, sheet)


key_stats["Lease Start"] = pd.to_datetime(key_stats["Lease Start"], format='%m/%d/%Y')

for unit_type in key_stats["Floorplan"].unique():
    current_key_stats = key_stats[(key_stats["Floorplan"] == unit_type) &
                    (key_stats["Lease Start"] > datetime.datetime(2020,7,1)) &
                    (key_stats["RENT"] > 0)]

    x = current_key_stats["Lease Start"].values
    y = current_key_stats["RENT"].values
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
    ax.set_xlim(min_date, max_date)

    # Get values for the trend line analysis
    if len(x) > 1:
        x_num = dates.date2num(list(x))
        trend = np.polyfit(x_num, list(y), 1)
        fit = np.poly1d(trend)
        plt.plot(list(x), fit(list(x_num)), "r--")

    # plt.show()
    plt.savefig(chart_folder + "Taylor Commons" + " - " + unit_type + ".png")
    plt.close()
