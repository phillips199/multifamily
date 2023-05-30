import numpy
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

threshold = 500

root = tkinter.Tk()
root.update()
print("Please select the key stats file that you want to analyze.")
key_stats_file_path = filedialog.askopenfilename()
print(key_stats_file_path)
key_stats_xls = pd.ExcelFile(key_stats_file_path)
key_stats_df = pd.read_excel(key_stats_xls, "Sheet1")
key_stats_df['Bldg-Unit'] = key_stats_df['Bldg-Unit'].apply(lambda x: str(x).zfill(4))
key_stats_df = key_stats_df.set_index('Bldg-Unit', drop=False)

# print("Please select the renovation tracker file.")
# renovation_file_path = filedialog.askopenfilename()
# renovation_xls = pd.ExcelFile(renovation_file_path)
# renovation_df = pd.read_excel(renovation_xls)
# renovation_df['Bldg-Unit'] = renovation_df['Bldg-Unit'].apply(lambda x: str(x).zfill(4))
# renovation_df = renovation_df.set_index('Bldg-Unit', drop=False)
#
# key_stats_final_df = key_stats_df.join(renovation_df, lsuffix='_keystats', rsuffix='_types')

problem_report = []

for unit in key_stats_df['Bldg-Unit'].unique():

    print("Unit: " + str(unit))
    unit_df = key_stats_df[key_stats_df['Bldg-Unit'] == unit]

    tenant_balances = {}
    balanced_owed_rows = []
    for index, row in unit_df.iterrows():

        if len(balanced_owed_rows) > 0:
            if row['Lease Start'] != balanced_owed_rows[-1]['Lease Start']:

                days_increasing = timedelta(days=0)
                days_above_threshold = timedelta(days=0)
                above_threshold = True
                increasing = True
                for i in range(1, len(balanced_owed_rows)):

                    date_last = datetime.strptime(balanced_owed_rows[-i]['Month'], "%Y.%m.%d")
                    date_before_last = datetime.strptime(balanced_owed_rows[-i - 1]['Month'], "%Y.%m.%d")
                    if balanced_owed_rows[-i]['Balance'] > balanced_owed_rows[-i - 1]['Balance'] and increasing:
                        days_increasing += date_last - date_before_last
                    else:
                        increasing = False
                    if balanced_owed_rows[-i]['Balance'] > threshold and above_threshold:
                        days_above_threshold += date_last - date_before_last
                    else:
                        above_threshold = False

                tenant_balances['Bldg-Unit'] = balanced_owed_rows[-1]['Bldg-Unit']
                tenant_balances['Resident'] = balanced_owed_rows[-1]['Resident']
                tenant_balances['Lease Start'] = balanced_owed_rows[-1]['Lease Start']
                tenant_balances['Lease End'] = balanced_owed_rows[-1]['Lease End']
                tenant_balances['Write-Off Date'] = datetime.strptime(row['Month'], "%Y.%m.%d")
                tenant_balances['Write-Off Amount'] = balanced_owed_rows[-1]['Balance']
                tenant_balances['Days Increasing'] = days_increasing
                tenant_balances['Days Above $' + str(threshold)] = days_increasing
                problem_report.append(tenant_balances.copy())
                balanced_owed_rows = []

        if row['Balance'] > 0:
            balanced_owed_rows.append(row)
        else:
            balanced_owed_rows = []

    if len(balanced_owed_rows) > 0:
        days_increasing = timedelta(days=0)
        days_above_threshold = timedelta(days=0)
        above_threshold = True
        increasing = True
        for i in range(1, len(balanced_owed_rows)):

            date_last = datetime.strptime(balanced_owed_rows[-i]['Month'], "%Y.%m.%d")
            date_before_last = datetime.strptime(balanced_owed_rows[-i - 1]['Month'], "%Y.%m.%d")
            if balanced_owed_rows[-i]['Balance'] > balanced_owed_rows[-i - 1]['Balance'] and increasing:
                days_increasing += date_last - date_before_last
            else:
                increasing = False
            if balanced_owed_rows[-i]['Balance'] > threshold and above_threshold:
                days_above_threshold += date_last - date_before_last
            else:
                above_threshold = False

        tenant_balances['Bldg-Unit'] = balanced_owed_rows[-1]['Bldg-Unit']
        tenant_balances['Resident'] = balanced_owed_rows[-1]['Resident']
        tenant_balances['Lease Start'] = balanced_owed_rows[-1]['Lease Start']
        tenant_balances['Lease End'] = balanced_owed_rows[-1]['Lease End']
        tenant_balances['Write-Off Date'] = numpy.nan
        tenant_balances['Write-Off Amount'] = numpy.nan
        tenant_balances['Balance Due'] = balanced_owed_rows[-1]['Balance']
        tenant_balances['Balance Date'] = date_last
        tenant_balances['Days Increasing'] = days_increasing
        tenant_balances['Days Above $' + str(threshold)] = days_increasing
        problem_report.append(tenant_balances.copy())

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

path, filename = os.path.split(key_stats_file_path)
name, ext = os.path.splitext(filename)

writer = ExcelWriter(os.path.join(path, "check_tenant_balances" + time + ext))
problem_report_data_df.to_excel(writer, "Sheet1", index=False)
writer.save()
