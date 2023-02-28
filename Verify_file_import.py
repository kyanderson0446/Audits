import os
import glob
import csv

# Get the list of all files and directories
folder = "P:\PACS\Finance\Month End Close\All - Month End Reporting\Workday_testing"
dir_list = os.listdir(folder)

# prints all files
print(dir_list)

# Example.csv gets created in the current working directory
with open(fr'P:\PACS\Finance\Month End Close\All - Month End Reporting\Workday_testing\Audit_log_list_.csv', 'w', newline='') as csvfile:
    my_writer = csv.writer(csvfile, delimiter=' ')
    my_writer.writerow(dir_list)

