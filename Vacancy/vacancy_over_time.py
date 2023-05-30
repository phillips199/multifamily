import pandas as pd
import numpy as np
from pandas import ExcelWriter
from datetime import datetime, timedelta
import tkinter
from tkinter import filedialog

start_date = datetime(2022, 6, 1)
end_date = datetime(2023, 3, 31)

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
unit_types_df =unit_types_df.set_index('Bldg-Unit')

unit_types_dict = pd.pivot_table(unit_types_df, values=['Count'], index=['Unit Type'],
                    aggfunc=np.count_nonzero, fill_value=0).to_dict()['Count']

key_stats_new_df = key_stats_df.join(unit_types_df, lsuffix='_keystats', rsuffix='_types')

vacancy_stats = []
for month_col in key_stats_new_df['Month'].unique():

    print("Month: " + month_col)
    year, month, day = month_col.split('.')
    date = datetime(int(year), int(month), 1)
    while date.month == int(month):

        current_residents_df = key_stats_new_df[(key_stats_new_df['Lease Start'] <= pd.Timestamp(date)) &
                                                (key_stats_new_df['Lease End'] >= pd.Timestamp(date)) &
                                                (key_stats_new_df['Month'] == month_col)]
        current_residents_dict = pd.pivot_table(current_residents_df,values=['Month'],
                                                aggfunc=np.count_nonzero,index=['Unit Type']).to_dict()['Month']

        day_dict = {}
        for unit_type in unit_types_dict.keys():

            if unit_type in current_residents_dict.keys():
                day_dict[unit_type + ' (%)'] = current_residents_dict[unit_type] / unit_types_dict[unit_type]
                day_dict[unit_type + ' (#)'] = current_residents_dict[unit_type]

        day_dict['Total (%)'] = sum(current_residents_dict.values()) / sum(unit_types_dict.values())
        day_dict['Total (#)'] = sum(current_residents_dict.values())
        day_dict['Date'] = date
        vacancy_stats.append(day_dict)

        date = date + timedelta(days=1)

vacancy_stats_df = pd.DataFrame.from_dict(vacancy_stats)
writer = ExcelWriter("vacancy_stats_by_day.xlsx")
vacancy_stats_df.to_excel(writer, "Sheet1", index=False)
writer.save()
