import pandas as pd
from pandas import ExcelWriter
from datetime import datetime, timedelta
import tkinter
from tkinter import filedialog
from pprint import pprint
import os

time = datetime.now().strftime(" %Y-%m-%dT%H-%M-%S")

buffer_days = 90
start_date_col = "Start Date"
# start_date_col = 'Const.\nStrt\nDate'

day_increment = 30
price_increment = 0.95
increment_levels = 5
thresholds = []
for i in range(1, increment_levels):
    thresholds.append((day_increment * i, price_increment ** i))

root = tkinter.Tk()
root.update()
print("Please select the key stats file that you want to analyze.")
key_stats_file_path = filedialog.askopenfilename()
print(key_stats_file_path)
key_stats_xls = pd.ExcelFile(key_stats_file_path)
key_stats_df = pd.read_excel(key_stats_xls, "Sheet1")
key_stats_df['Bldg-Unit'] = key_stats_df['Bldg-Unit'].apply(lambda x: str(x).zfill(4))
key_stats_df = key_stats_df.set_index('Bldg-Unit', drop=False)

print("Please select the renovation tracker file.")
renovation_file_path = filedialog.askopenfilename()
renovation_xls = pd.ExcelFile(renovation_file_path)
renovation_df = pd.read_excel(renovation_xls)
renovation_df['Bldg-Unit'] = renovation_df['Bldg-Unit'].apply(lambda x: str(x).zfill(4))
renovation_df = renovation_df.set_index('Bldg-Unit', drop=False)

key_stats_final_df = key_stats_df.join(renovation_df, lsuffix='_keystats', rsuffix='_types')

problem_report = []

for unit in key_stats_df['Bldg-Unit'].unique():

    print("Unit: " + str(unit))
    unit_df = key_stats_final_df[key_stats_final_df['Bldg-Unit_keystats'] == unit]

    vacant_rows = []
    longest_vacancy = {'Days Vacant': 0}
    for index, row in unit_df.iterrows():

        if len(vacant_rows) > 0:
            if pd.isnull(row['Lease Start']):
                date = datetime.strptime(row['Month'], "%Y.%m.%d")
                buffer_date = date + timedelta(days=buffer_days)
                const_start = row[start_date_col]
                if pd.isnull(const_start) or buffer_date < const_start:
                    rent = row['Market Rent']
                    for each in vacant_rows:
                        days_vacant = date - datetime.strptime(each['Month'], "%Y.%m.%d")
                        rent_ratio = rent / each['Market Rent']
                        for threshold in thresholds:
                            if days_vacant > timedelta(days=threshold[0]) and rent_ratio > threshold[1]:
                                int_days_vacant = days_vacant / timedelta(days=1)
                                if int_days_vacant > longest_vacancy['Days Vacant']:
                                    longest_vacancy = {'Bldg-Unit': unit,
                                                       'First Vacant Date': datetime.strptime(each['Month'],
                                                                                              "%Y.%m.%d"),
                                                       'Last Vacant Date': date,
                                                       'First Market Rent': each['Market Rent'],
                                                       'Last Market Rent': rent, 'Days Vacant': int_days_vacant,
                                                       'Rent Change': '{:.1%}'.format(rent_ratio - 1)}
                vacant_rows.append(row)

            else:
                if longest_vacancy['Days Vacant'] > 0:
                    pprint(longest_vacancy)
                    problem_report.append(longest_vacancy)
                    longest_vacancy = {'Days Vacant': 0}
                vacant_rows = []
        else:
            if pd.isnull(row['Lease Start']):
                vacant_rows.append(row)

    if longest_vacancy['Days Vacant'] > 0:
        pprint(longest_vacancy)
        problem_report.append(longest_vacancy)

problem_report_data_df = pd.DataFrame.from_dict(problem_report)
problem_report_data_df = problem_report_data_df.set_index('Bldg-Unit', drop=False)

print("Please select the unit types file that you want to use.")
unit_types_file_path = filedialog.askopenfilename()
print(unit_types_file_path)
unit_types_xls = pd.ExcelFile(unit_types_file_path)
unit_types_df = pd.read_excel(unit_types_xls)
unit_types_df['Bldg-Unit'] = unit_types_df['Bldg-Unit'].apply(lambda x: str(x).zfill(4))
unit_types_df = unit_types_df.set_index('Bldg-Unit', drop=False)

problem_report_data_df = problem_report_data_df.join(unit_types_df, rsuffix='_types')
problem_report_data_reno_df = problem_report_data_df.join(renovation_df, rsuffix='_types')

path, filename = os.path.split(key_stats_file_path)
name, ext = os.path.splitext(filename)

writer = ExcelWriter(os.path.join(path, "check_price_drops_reno" + time + ext))
problem_report_data_reno_df.to_excel(writer, "Sheet1", index=False)
writer.save()
