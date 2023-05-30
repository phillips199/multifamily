import pandas as pd
import numpy as np
from pandas import ExcelWriter
from datetime import datetime, timedelta
import tkinter
from tkinter import filedialog
import os

tenancies = []

time = datetime.now().strftime(" %Y-%m-%dT%H-%M-%S")
review_period = 120
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

for unit in key_stats_df['Bldg-Unit'].unique():

    unit_df = key_stats_df[key_stats_df['Bldg-Unit'] == unit]
    last_lease_stats = {}

    last_vacant = False
    for index, row in unit_df.iterrows():

        rent_roll_date = datetime.strptime(row['Month'], "%Y.%m.%d")

        if last_lease_stats:
            if pd.isnull(row['Lease Start']):
                if not last_vacant:
                    if rent_roll_date < last_lease_stats['End Date']:
                        last_lease_stats['End Date'] = rent_roll_date
                        last_lease_stats['Broken'] = True
                    last_vacant = True

            else:

                if last_lease_stats['Lease Start'] != row['Lease Start']:

                    tenant_dict = {'Bldg-Unit': str(unit).zfill(4),
                                   'Resident': last_lease_stats['Resident'],
                                   'Scheduled Rent': last_lease_stats['Scheduled Rent'],
                                   'Lease Start': last_lease_stats['Lease Start'],
                                   'Lease End': last_lease_stats['Lease End'],
                                   'Broken': last_lease_stats['Broken']}

                    tenancies.append(tenant_dict)

                    last_lease_stats['Resident'] = row['Resident']
                    last_lease_stats['Scheduled Rent'] = row['Scheduled Rent']
                    last_lease_stats['Lease Start'] = row['Lease Start']
                    last_lease_stats['Lease End'] = row['Lease End']
                    last_lease_stats['Broken'] = np.nan
                    if pd.isnull(row['Move-Out']):
                        last_lease_stats['End Date'] = row['Lease End']
                    else:
                        last_lease_stats['End Date'] = row['Move-Out']
                    last_vacant = False

                else:
                    if rent_roll_date <= last_lease_stats['Lease Start'] + np.timedelta64(review_period, 'D'):

                        last_lease_stats['Resident'] = row['Resident']
                        last_lease_stats['Scheduled Rent'] = row['Scheduled Rent']
                        last_lease_stats['Lease Start'] = row['Lease Start']
                        last_lease_stats['Lease End'] = row['Lease End']
                        last_lease_stats['Broken'] = np.nan
                        if pd.isnull(row['Move-Out']):
                            last_lease_stats['End Date'] = row['Lease End']
                        else:
                            last_lease_stats['End Date'] = row['Move-Out']
                        last_vacant = False

        else:

            if pd.isnull(row['Lease Start']):
                last_vacant = True

            else:
                if row['Scheduled Rent'] > 0:
                    last_lease_stats['Resident'] = row['Resident']
                    last_lease_stats['Scheduled Rent'] = row['Scheduled Rent']
                    last_lease_stats['Lease Start'] = row['Lease Start']
                    last_lease_stats['Lease End'] = row['Lease End']
                    last_lease_stats['Broken'] = np.nan
                    if pd.isnull(row['Move-Out']):
                        last_lease_stats['End Date'] = row['Lease End']
                    else:
                        last_lease_stats['End Date'] = row['Move-Out']
                    last_vacant = False

    if last_lease_stats:
        tenant_dict = {'Bldg-Unit': str(unit).zfill(4),
                       'Resident': last_lease_stats['Resident'],
                       'Scheduled Rent': last_lease_stats['Scheduled Rent'],
                       'Lease Start': last_lease_stats['Lease Start'],
                       'Lease End': last_lease_stats['Lease End'],
                       'Broken': last_lease_stats['Broken']}

        tenancies.append(tenant_dict)

print("Please select the unit types file that you want to use.")
unit_types_file_path = filedialog.askopenfilename()
print(unit_types_file_path)
unit_types_xls = pd.ExcelFile(unit_types_file_path)
unit_types_df = pd.read_excel(unit_types_xls)
unit_types_df['Bldg-Unit'] = unit_types_df['Bldg-Unit'].apply(lambda x: str(x).zfill(4))
unit_types_df = unit_types_df.set_index('Bldg-Unit', drop=False)

tenancies_data_df = pd.DataFrame.from_dict(tenancies)
tenancies_data_df = tenancies_data_df.set_index('Bldg-Unit', drop=False)
tenancies_data_df = tenancies_data_df.join(unit_types_df, lsuffix='_keystats', rsuffix='_types')
tenancies_data_df['Duration'] = tenancies_data_df['Lease End'] - tenancies_data_df['Lease Start']

problem_report_data_df = tenancies_data_df.join(unit_types_df, rsuffix='_types')

path, filename = os.path.split(key_stats_file_path)
name, ext = os.path.splitext(filename)

writer = ExcelWriter(os.path.join(path, "tenant_durations" + time + ext))
problem_report_data_df.to_excel(writer, "Sheet1", index=False)
writer.save()

date = max(min(tenancies_data_df['Lease Start']) + timedelta(days=rolling_average_days), start_date)
end_date = max(tenancies_data_df['Lease Start'])
rolling_averages = []
while date <= end_date:

    rolling_average_df = tenancies_data_df[(tenancies_data_df['Lease Start'] >= date -
                                            timedelta(days=rolling_average_days)) &
                                           (tenancies_data_df['Lease Start'] <= date)]
    rolling_average = rolling_average_df['Duration'].mean()
    rolling_averages.append({'Date': date, str(rolling_average_days) + 'day Rolling Avg': rolling_average})
    date = date + timedelta(days=1)

rolling_average_balances_df = pd.DataFrame.from_dict(rolling_averages)

writer = ExcelWriter(os.path.join(path, str(rolling_average_days) + 'day_rolling_avg_duration' + time + ext))
rolling_average_balances_df.to_excel(writer, "Sheet1", index=False)
writer.save()