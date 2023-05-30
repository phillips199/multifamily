import pandas as pd
from pandas import ExcelWriter
from datetime import datetime, timedelta
import tkinter
from tkinter import filedialog
import os
import numpy as np

time = datetime.now().strftime(" %Y-%m-%dT%H-%M-%S")

averages = []

average_period = 100
rolling_average_days = 30
start_date = datetime(2022, 1, 1)

print("Please select the key stats file that you want to analyze.")
root = tkinter.Tk()
root.update()
key_stats_file_path = filedialog.askopenfilename()
print(key_stats_file_path)
key_stats_xls = pd.ExcelFile(key_stats_file_path)
key_stats_df = pd.read_excel(key_stats_xls, "Sheet1")
key_stats_df['Bldg-Unit'] = key_stats_df['Bldg-Unit'].apply(lambda x: str(x).zfill(4))
key_stats_df = key_stats_df.set_index('Bldg-Unit', drop=False)
key_stats_df['Date'] = key_stats_df['Month'].apply(lambda x: datetime.strptime(x, "%Y.%m.%d"))
key_stats_df.dropna(subset=['Lease Start'], inplace=True)
key_stats_df = key_stats_df[key_stats_df['Lease Start'] <= datetime.now() - timedelta(days=average_period)]

for unit in key_stats_df['Bldg-Unit'].unique():

    print("Unit: " + str(unit))
    unit_df = key_stats_df[key_stats_df['Bldg-Unit'] == unit]

    for lease_date in unit_df['Lease Start'].unique():

        print("Lease Date: " + str(lease_date))
        unit_lease_df = unit_df[(unit_df['Lease Start'] == lease_date) &
                                (unit_df['Date'] <= lease_date + np.timedelta64(average_period,'D'))]

        if len(unit_lease_df) > 0:
            balance_average = unit_lease_df['Balance'].mean()
            balance_std = unit_lease_df['Balance'].std()

            balances_dict = {'Unit': str(unit).zfill(4), 'Resident': unit_lease_df['Resident'].values[-1],
                             'Lease Start': lease_date, 'Lease End': unit_lease_df['Lease End'].values[-1],
                             'Avg Balance': balance_average, 'Std Balance': balance_std,
                             'Move Out': unit_lease_df['Move-Out'].values[-1]}

            averages.append(balances_dict.copy())

print("Please select the unit types file that you want to use.")
unit_types_file_path = filedialog.askopenfilename()
print(unit_types_file_path)
unit_types_xls = pd.ExcelFile(unit_types_file_path)
unit_types_df = pd.read_excel(unit_types_xls, "Sheet1")
unit_types_df['Bldg-Unit'] = unit_types_df['Bldg-Unit'].apply(lambda x: str(x).zfill(4))
unit_types_df = unit_types_df.set_index('Bldg-Unit', drop=False)

average_balances_df = pd.DataFrame.from_dict(averages)
average_balances_df = average_balances_df.join(unit_types_df, rsuffix='_types')

path, filename = os.path.split(key_stats_file_path)
name, ext = os.path.splitext(filename)

writer = ExcelWriter(os.path.join(path, "average_" + str(average_period) + "day_balances" + time + ext))
average_balances_df.to_excel(writer, "Sheet1", index=False)
writer.save()

date = max(min(average_balances_df['Lease Start']) + timedelta(days=rolling_average_days), start_date)
end_date = max(average_balances_df['Lease Start'])
rolling_averages = []
while date <= end_date:

    rolling_average_df = average_balances_df[(average_balances_df['Lease Start'] >= date -
                                              timedelta(days=rolling_average_days)) &
                                             (average_balances_df['Lease Start'] <= date)]
    rolling_average = rolling_average_df['Avg Balance'].mean()
    rolling_averages.append({'Date': date, str(rolling_average_days) + 'day Rolling Avg': rolling_average})
    date = date + timedelta(days=1)

rolling_average_balances_df = pd.DataFrame.from_dict(rolling_averages)

writer = ExcelWriter(os.path.join(path, str(rolling_average_days) + 'day_rolling_avg_' + str(average_period) +
                                  'day_avg_balances' + time + ext))
rolling_average_balances_df.to_excel(writer, "Sheet1", index=False)
writer.save()

