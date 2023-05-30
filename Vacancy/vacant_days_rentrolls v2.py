import pandas as pd
import numpy as np
from pandas import ExcelWriter
from datetime import datetime
import tkinter
from tkinter import filedialog
import os

vacancies = []

time = datetime.now().strftime(" %Y-%m-%dT%H-%M-%S")

start_date = datetime(2022, 9, 1)
end_date = datetime(2023, 3, 31)

print("Please select the key stats file that you want to analyze.")
root = tkinter.Tk()
root.update()
key_stats_file_path = filedialog.askopenfilename()
print(key_stats_file_path)
key_stats_xls = pd.ExcelFile(key_stats_file_path)
key_stats_df = pd.read_excel(key_stats_xls, "Sheet1")
key_stats_df['Bldg-Unit'] = key_stats_df['Bldg-Unit'].apply(lambda x: str(x).zfill(4))
key_stats_df = key_stats_df.set_index('Bldg-Unit', drop=False)

print("Please select the unit types file that you want to use.")
unit_types_file_path = filedialog.askopenfilename()
print(unit_types_file_path)
unit_types_xls = pd.ExcelFile(unit_types_file_path)
unit_types_df = pd.read_excel(unit_types_xls, "Sheet1")
unit_types_df['Bldg-Unit'] = unit_types_df['Bldg-Unit'].apply(lambda x: str(x).zfill(4))
unit_types_df = unit_types_df.set_index('Bldg-Unit', drop=False)

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
                if row['Scheduled Rent'] > 0:
                    if pd.Timestamp(last_lease_stats['End Date']) < pd.Timestamp(row['Lease Start']):

                        new_rent = row['Scheduled Rent']
                        old_rent = last_lease_stats['Rent']
                        days_vacant = pd.Timestamp(row['Lease Start']) - pd.Timestamp(last_lease_stats['End Date'])
                        if new_rent > 0 and old_rent > 0:
                            uplift = new_rent / old_rent - 1
                            adjusted_uplift = (new_rent / old_rent) ** (365 / (364 + days_vacant.days)) - 1
                        else:
                            uplift = ''
                            adjusted_uplift = ''

                        vacancy_dict = {'Bldg-Unit': str(unit).zfill(4),

                                        'Prior Resident': last_lease_stats['Resident'],
                                        'Prior Rent': last_lease_stats['Rent'],
                                        'Broken': last_lease_stats['Broken'],
                                        'Start Vacant Date': last_lease_stats['End Date'],

                                        'End Vacant Date': row['Lease Start'],
                                        'Next Rent': row['Scheduled Rent'],
                                        'Next Resident': row['Resident'],
                                        'Days Vacant': days_vacant,
                                        'Uplift': uplift,
                                        'Adjusted Uplift': adjusted_uplift}
                        vacancies.append(vacancy_dict)

                    last_lease_stats['Start Date'] = row['Lease Start']
                    last_lease_stats['Rent'] = row['Scheduled Rent']
                    last_lease_stats['Resident'] = row['Resident']
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

                if pd.Timestamp(row['Lease Start']) > start_date:
                    if row['Scheduled Rent'] > 0:
                        vacancy_dict = {'Bldg-Unit': str(unit).zfill(4),

                                        'Prior Resident': '',
                                        'Prior Rent': '',
                                        'Start Vacant Date': start_date,

                                        'End Vacant Date': row['Lease Start'],
                                        'Days Vacant': pd.Timestamp(row['Lease Start']) - pd.Timestamp(start_date),
                                        'Next Rent': row['Scheduled Rent'],
                                        'Next Resident': row['Resident']}

                        vacancies.append(vacancy_dict)

                    else:
                        print("$0 lease for resident " + row['Resident'])

                last_lease_stats['Start Date'] = row['Lease Start']
                last_lease_stats['Rent'] = row['Scheduled Rent']
                last_lease_stats['Resident'] = row['Resident']
                last_lease_stats['Broken'] = np.nan
                if pd.isnull(row['Move-Out']):
                    last_lease_stats['End Date'] = row['Lease End']

                else:
                    last_lease_stats['End Date'] = row['Move-Out']

    if pd.Timestamp(last_lease_stats['End Date']) < end_date:
        vacancy_dict = {'Bldg-Unit': str(unit).zfill(4),
                        'Prior Resident': last_lease_stats['Resident'],
                        'Broken': last_lease_stats['Broken'],
                        'Prior Rent': last_lease_stats['Rent'],
                        'Start Vacant Date': last_lease_stats['End Date'],

                        'Next Resident': 'Vacant',
                        'End Vacant Date': end_date,
                        'Days Vacant': pd.Timestamp(end_date) - pd.Timestamp(last_lease_stats['End Date'])}

        vacancies.append(vacancy_dict)


vacancies_data_df = pd.DataFrame.from_dict(vacancies)
vacancies_data_df = vacancies_data_df.set_index('Bldg-Unit', drop=False)
vacancies_data_df = vacancies_data_df.join(unit_types_df, lsuffix='_keystats', rsuffix='_types')

path, filename = os.path.split(key_stats_file_path)
name, ext = os.path.splitext(filename)

if input("Do you have a renovation tracker you want to use?") == "Y":
    print("Please select the renovation tracker file.")
    renovation_file_path = filedialog.askopenfilename()
    renovation_xls = pd.ExcelFile(renovation_file_path)
    renovation_df = pd.read_excel(renovation_xls)

    renovation_df = renovation_df.rename(columns={"Unit": "Bldg-Unit"})
    renovation_df['Bldg-Unit'] = renovation_df['Bldg-Unit'].apply(lambda x: str(x).zfill(4))

    key_stats_final_df = vacancies_data_df.merge(renovation_df, on="Bldg-Unit", how="outer")
    writer = ExcelWriter(path + "key_stats_vacancies_reno" + time + ext)
    key_stats_final_df.to_excel(writer, "Sheet1", index=False)
    writer.save()

else:

    writer = ExcelWriter(path + "key_stats_vacancies" + time + ext)
    vacancies_data_df.to_excel(writer, "Sheet1", index=False)
    writer.save()
