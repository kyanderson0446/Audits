import time
from glob import glob
import os
from datetime import date
import csv
import pandas as pd
import xlwings as xw

# Paths for where the reports will be saved and the csv which maps facility names to legal names
audit_path = fr"C:\Users\kyle.anderson\Documents\audit_test\*.xlsx"
z_a_Facility = fr"C:\Users\kyle.anderson\PycharmProjects\Audits\ZtoA_facilities.csv"

# Read in csv
df = pd.read_csv(z_a_Facility)
df['Legal'] = df['Legal']
df['Facility'] = df['Facility']

date = date.today()
# Reporting for previous month
folder_month = date.today().month-1
folder_year = date.today().year

xlwings.App.display_alerts = False
xw.Interactive = False
xw.Visible = False

# Make subfolder for each month
try:
    os.mkdirs(fr"P:\PACS\Finance\Month End Close\All - Month End Reporting\Workday_AP_Aging\{folder_month} - {folder_year}")
except:
    pass

# Loop through each saved file and rename by using xlwings to grab name from workbook
for x in glob(audit_path):

    app = xw.App(add_book=False)
    wb = xw.Book(x, update_links=False)

    try:
        f_name = wb.sheets[0].range("A2").value
        time.sleep(1)
        if df.loc[df['Legal'] == f_name, 'Facility'].values[0]:
            new_name = df.loc[df['Legal'] == f_name, 'Facility'].values[0]
            print(new_name, " Saving...")
            wb.save(
                fr"P:\PACS\Finance\Month End Close\All - Month End Reporting\Workday_AP_Aging\{folder_month} - {folder_year}\{folder_year} {folder_month} {new_name} Payables Aging.xlsx")
            wb.close()

        # For new reports or when the facility hasn't been added to the mapper, save the file under the cell name

        else:
            print(f_name, "- Add name to mapping")
            wb.save(
                fr"P:\PACS\Finance\Month End Close\All - Month End Reporting\Workday_AP_Aging\{folder_month} - {folder_year}\{folder_year} {folder_month} {f_name} Payables Aging.xlsx")
            wb.close()
        app.quit()
    # Some downloads have different headers
    except:
        f_name = wb.sheets[0].range("B2").value
        time.sleep(1)

        # Using the csv to map the cell name to the 'Facility' name

        if df.loc[df['Legal'] == f_name, 'Facility'].values[0]:
            new_name = df.loc[df['Legal'] == f_name, 'Facility'].values[0]
            print(new_name, " Saving...")
            wb.save(fr"P:\PACS\Finance\Month End Close\All - Month End Reporting\Workday_AP_Aging\{folder_month} - {folder_year}\{folder_year} {folder_month} {new_name} Payables Aging.xlsx")
            wb.close()

        # For new reports or when the facility hasn't been added to the mapper, save the file under the cell name

        else:
            print(f_name, "- Add name to mapping")
            wb.save(
                fr"P:\PACS\Finance\Month End Close\All - Month End Reporting\Workday_AP_Aging\{folder_month} - {folder_year}\{folder_year} {folder_month} {f_name} Payables Aging.xlsx")
            wb.close()
        app.quit()


