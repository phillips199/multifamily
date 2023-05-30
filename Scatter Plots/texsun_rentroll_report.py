
import pandas as pd
import datetime
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as dates

min_date = datetime.datetime(2020,7,1)
max_date = datetime.datetime(2023,4,15)

cutoff_dates = {'Bridge Hollow': datetime.datetime(2022,11,1),
                'Woodstone': datetime.datetime(2022,11,1),
                'Heights On Perrin': datetime.datetime(2022,10,1),
                'San Mateo': datetime.datetime(2022,10,1),
                'Trailside': datetime.datetime(2022,12,1)
                }
cw_underwriting = {}

lease_xls = pd.ExcelFile('/Users/mikephillips/My Drive/Cottonwood/C. Equity/1. Project Documents/Leases/Lease Underwriting.xlsx')
for lease_sheet in lease_xls.sheet_names:
    lease_df = pd.read_excel(lease_xls, lease_sheet)
    counter = 0
    lease_details = {}
    for uw_type in lease_df["Type"]:
        lease_details[uw_type] = [lease_df["Before"][counter],lease_df["After"][counter]]
        counter += 1
    cw_underwriting[lease_sheet] = lease_details

rentroll_xls = pd.ExcelFile('/Users/mikephillips/My Drive/Cottonwood/C. Equity/1. Project Documents/Financials/Deals/RentRolls_February2023 t-1680044954.793422.xlsx')
for sheet in rentroll_xls.sheet_names:
    print("starting with " + sheet + "....")
    response = input("Do you want to run graphs for " + sheet + "?")
    if response == "Y":
        df = pd.read_excel(rentroll_xls, sheet)

        first_col = df["Unnamed: 0"]
        first_row = first_col[first_col == "Bldg-Unit"].index[0] + 1
        last_row = first_col[first_col == "Status Summary"].index[0] - 2

        key_stats = second_col = df[["Unnamed: 0","Unnamed: 1","Unnamed: 4","Unnamed: 6","Unnamed: 10"]][first_row:last_row]
        key_stats = key_stats.rename(columns={"Unnamed: 0": "Unit", "Unnamed: 1": "Type",
                                              "Unnamed: 4": "Resident", "Unnamed: 6": "Date", "Unnamed: 10": "Rate"})

        for unit_type in key_stats["Type"].unique():
            current_key_stats = key_stats[(key_stats["Type"] == unit_type) &
                            (key_stats["Date"] > datetime.datetime(2020,7,1)) &
                            (key_stats["Rate"] > 0)]

            x = current_key_stats["Date"].values
            y = current_key_stats["Rate"].values
            plt.figure(figsize=(10,8))
            plt.subplots_adjust(top=0.9, bottom=0.25, left=0.2)
            ax = plt.axes()
            ax.scatter(x,y)
            plt.plot([cutoff_dates[sheet],cutoff_dates[sheet]],[600,1800],color='grey', linestyle='dashed')
            plt.plot([min_date, cutoff_dates[sheet]],[cw_underwriting[sheet][unit_type][0],cw_underwriting[sheet][unit_type][0]],color='blue', linestyle='dashed')
            plt.plot([cutoff_dates[sheet], max_date],[cw_underwriting[sheet][unit_type][1],cw_underwriting[sheet][unit_type][1]],color='green', linestyle='dashed')
            ax.set_title(sheet + " - " + unit_type + "\n", fontsize=20)
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

            # plt.show()
            plt.savefig(sheet + " - " + unit_type + ".png")
            plt.close()
    else:
        print("Skipping " + sheet + "....")