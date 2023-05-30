
import pandas as pd
import datetime
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as dates

min_date = datetime.datetime(2020, 7, 1)
max_date = datetime.datetime(2023, 4, 15)

chart_folder = "/Users/mikephillips/My Drive/Cottonwood/C. Equity/3. Work Product/Leasing Charts/"

cutoff_date = datetime.datetime(2023, 1, 1)
cw_underwriting = {}

lease_xls = pd.ExcelFile('/Users/mikephillips/My Drive/Cottonwood/C. Equity/1. Project Documents/Leases/Lease Underwriting.xlsx')
lease_df = pd.read_excel(lease_xls, "Raider")
counter = 0
lease_details = {}
for uw_type in lease_df["Type"]:
    lease_details[uw_type] = [lease_df["Before"][counter],lease_df["After"][counter]]
    counter += 1
cw_underwriting["Raider"] = lease_details

# tracker_xls = pd.ExcelFile("    ")
# tracker_df = pd.read_excel(tracker_xls, "Tracker")
# print(tracker_df.head())
# tracker_unit_col = tracker_df["Unit\n#"]

rentroll_xls = pd.ExcelFile("/Users/mikephillips/My Drive/Unique/Raider_Modified_Rentroll t-1681849747.527013.xlsx")
sheet = "Sheet1"
key_stats = pd.read_excel(rentroll_xls, sheet)
key_stats = key_stats.rename(columns={"Bldg-Unit": "Unit", "Unit Type": "Type",
                                      "Lease Start": "Date", "Rent": "Rate"})

# key_stats["Type"] = ""
# key_stats["Prior Reno"] = ""
# key_stats["Tracker Type"] = ""
# key_stats["TAM Reno"] = ""
# key_stats["Reno Complete"] = ""

length = len(key_stats["Unit"])

for unit_type in key_stats["Type"].unique():
    current_key_stats = key_stats[(key_stats["Type"] == unit_type) &
                                  (key_stats["Date"] > min_date) &
                                  (key_stats["Rate"] > 0)]

    x = current_key_stats["Date"].values
    y = current_key_stats["Rate"].values
    plt.figure(figsize=(10, 8))
    plt.subplots_adjust(top=0.9, bottom=0.25, left=0.2)
    ax = plt.axes()
    ax.scatter(x, y)
    plt.plot([cutoff_date, cutoff_date], [600, 2100], color='grey', linestyle='dashed')
    plt.plot([min_date, cutoff_date],
             [cw_underwriting["Raider"][unit_type][0], cw_underwriting["Raider"][unit_type][0]], color='blue',
             linestyle='dashed')
    plt.plot([cutoff_date, max_date],
             [cw_underwriting["Raider"][unit_type][1], cw_underwriting["Raider"][unit_type][1]], color='green',
             linestyle='dashed')
    ax.set_title("Raider - " + unit_type + "\n", fontsize=20)
    ax.set_ylabel('Rates\n', fontsize=20)
    ax.set_ylim(600, 2100)
    ax.set_xlabel('\nLease Dates', fontsize=20)
    ax.tick_params(axis='both', labelsize=20)
    ax.tick_params('x', labelrotation=30)
    ax.set_xlim(min_date, max_date)

    # Get values for the trend line analysis
    if len(x) > 1:
        x_num = dates.date2num(list(x))
        trend = np.polyfit(x_num, list(y), 1)
        fit = np.poly1d(trend)
        plt.plot(list(x), fit(list(x_num)), "r--")

    # plt.show()
    plt.savefig(chart_folder + "Raider" + " - " + unit_type + ".png")
    plt.close()
