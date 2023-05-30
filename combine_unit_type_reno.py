import pandas as pd
from pandas import ExcelWriter
import tkinter
from tkinter import filedialog

vacancies = []

print("Please select the unit types file that you want to use.")
unit_types_file_path = filedialog.askopenfilename()
print(unit_types_file_path)
unit_types_xls = pd.ExcelFile(unit_types_file_path)
unit_types_df = pd.read_excel(unit_types_xls)
unit_types_df['Bldg-Unit'] = unit_types_df['Bldg-Unit'].apply(lambda x: str(x).zfill(4))
unit_types_df = unit_types_df.set_index('Bldg-Unit', drop=False)

print("Please select the renovation tracker that you want to use.")
root = tkinter.Tk()
root.update()
reno_tracker_file_path = filedialog.askopenfilename()
print(reno_tracker_file_path)
reno_tracker_xls = pd.ExcelFile(reno_tracker_file_path)
reno_tracker_df = pd.read_excel(reno_tracker_xls)
reno_tracker_df = reno_tracker_df.set_index('Bldg-Unit', drop=False)

key_stats_new_df = reno_tracker_df.join(unit_types_df, lsuffix='_reno', rsuffix='_types', how='outer')

writer = ExcelWriter("Unit Types Reno.xlsx")
key_stats_new_df.to_excel(writer, "Sheet1", index=False)
writer.save()
