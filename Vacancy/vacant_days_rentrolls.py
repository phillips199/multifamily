import pandas as pd
from pandas import ExcelWriter
import datetime
import tkinter
from tkinter import filedialog

vacancies = []

start_date = datetime.datetime(2022, 9, 1)
end_date = datetime.datetime(2023, 3, 31)

print("Please select the key stats file that you want to analyze.")
root = tkinter.Tk()
root.update()
key_stats_file_path = filedialog.askopenfilename()
print(key_stats_file_path)
key_stats_xls = pd.ExcelFile(key_stats_file_path)
key_stats_df = pd.read_excel(key_stats_xls, "Sheet1")
key_stats_df = key_stats_df.set_index('Bldg-Unit', drop=False)

print("Please select the unit types file that you want to use.")
unit_types_file_path = filedialog.askopenfilename()
print(unit_types_file_path)
unit_types_xls = pd.ExcelFile(unit_types_file_path)
unit_types_df = pd.read_excel(unit_types_xls, "Sheet1")
unit_types_df = unit_types_df.set_index('Bldg-Unit', drop=False)

key_stats_new_df = key_stats_df.join(unit_types_df, rsuffix='_types')

for unit in key_stats_new_df['Bldg-Unit_keystats'].unique():

    print("Unit: " + str(unit))
    unit_df = key_stats_new_df[key_stats_new_df['Bldg-Unit'] == unit]
    last_lease_stats = {}

    unique_dates = unit_df['Lease Start'].unique()
    if len(unique_dates) == 1 and pd.isnull(unique_dates[0]):
        vacancy_dict = {'Unit': str(unit).zfill(4), 'Unit Type': unit_type, 'Start Vacant Date': start_date,
                        'Next Resident': 'Vacant',
                        'End Vacant Date': end_date,
                        'Days Vacant': pd.Timestamp(end_date) - pd.Timestamp(start_date)}

        vacancies.append(vacancy_dict)

    else:

        for lease_date in unit_df['Lease Start'].unique():

            if not pd.isnull(lease_date):
                print("Lease Date: " + str(lease_date))
                unit_lease_df = unit_df[unit_df['Lease Start'] == lease_date]
                if last_lease_stats:

                    if last_lease_stats['End Date'] < lease_date:

                        new_rent = unit_lease_df['Scheduled Rent'].values[-1]
                        old_rent = last_lease_stats['Rent']
                        days_vacant = pd.Timestamp(lease_date) - pd.Timestamp(last_lease_stats['End Date'])
                        if new_rent > 0 and old_rent > 0:
                            uplift = new_rent / old_rent - 1
                            adjusted_uplift = (new_rent / old_rent) ** (365 / (364 + days_vacant.days)) - 1
                        else:
                            uplift = 'N/A'
                            adjusted_uplift = 'N/A'

                        vacancy_dict = {'Unit': str(unit).zfill(4), 'Unit Type': unit_type,
                                        'Prior Resident': last_lease_stats['Resident'],
                                        'Prior Rent': last_lease_stats['Rent'],
                                        'Start Vacant Date': last_lease_stats['End Date'],
                                        'Move Out': last_lease_stats['Move Out'],
                                        'End Vacant Date': lease_date,
                                        'Next Rent': unit_lease_df['Scheduled Rent'].values[-1],
                                        'Next Resident': unit_lease_df['Resident'].values[-1],
                                        'Days Vacant': days_vacant,
                                        'Uplift': uplift,
                                        'Adjusted Uplift': adjusted_uplift}

                        vacancies.append(vacancy_dict)

                else:

                    if pd.Timestamp(unit_lease_df['Lease Start'].values[-1]) > start_date:

                        vacancy_dict = {'Unit': str(unit).zfill(4),
                                        'Unit Type': unit_type,
                                        'Prior Resident': 'N/A',
                                        'Prior Rent': 'N/A',
                                        'Start Vacant Date': start_date,
                                        'End Vacant Date': unit_lease_df['Lease Start'].values[-1],
                                        'Days Vacant': pd.Timestamp(unit_lease_df['Lease Start'].values[-1]) -
                                                         pd.Timestamp(start_date),
                                        'Next Rent': unit_lease_df['Scheduled Rent'].values[-1],
                                        'Next Resident': unit_lease_df['Resident'].values[-1]}

                        vacancies.append(vacancy_dict)

                last_lease_stats['Rent'] = unit_lease_df['Scheduled Rent'].values[-1]
                last_lease_stats['Resident'] = unit_lease_df['Resident'].values[-1]
                if pd.isnull(unit_lease_df['Move-Out'].values[-1]):

                    last_lease_stats['End Date'] = unit_lease_df['Lease End'].values[-1]
                    last_lease_stats['Move Out'] = True

                else:

                    last_lease_stats['End Date'] = unit_lease_df['Move-Out'].values[-1]
                    last_lease_stats['Move Out'] = True

        if pd.Timestamp(last_lease_stats['End Date']) < end_date:
            vacancy_dict = {'Unit': str(unit).zfill(4), 'Unit Type': unit_type,
                            'Prior Resident': last_lease_stats['Resident'],
                            'Prior Rent': last_lease_stats['Rent'],
                            'Start Vacant Date': last_lease_stats['End Date'],
                            'Move Out': last_lease_stats['Move Out'],
                            'Next Resident': 'Vacant',
                            'End Vacant Date': end_date,
                            'Days Vacant': pd.Timestamp(end_date) -
                                           pd.Timestamp(last_lease_stats['End Date'])}

            vacancies.append(vacancy_dict)

vacancies_data_df = pd.DataFrame.from_dict(vacancies)

if input("Do you have a renovation tracker you want to use?") == "Y":
    print("Please select the renovation tracker file.")
    renovation_file_path = filedialog.askopenfilename()
    renovation_xls = pd.ExcelFile(renovation_file_path)
    renovation_df = pd.read_excel(renovation_xls)
    renovation_df['Unit'] = renovation_df['Unit'].apply(lambda x: str(x).zfill(4))

    key_stats_final_df = vacancies_data_df.merge(renovation_df, on="Unit", how="outer")
    writer = ExcelWriter("key_stats_vacancies_renovation.xlsx")
    key_stats_final_df.to_excel(writer, "Sheet1", index=False)
    writer.save()

else:

    writer = ExcelWriter("key_stats_vacancies" ".xlsx")
    vacancies_data_df.to_excel(writer, "Sheet1", index=False)
    writer.save()
