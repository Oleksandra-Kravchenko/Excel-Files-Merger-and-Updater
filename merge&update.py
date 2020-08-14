"""
This utility merges excel files from a common directory and updates the main workbook by adding data from the merged files into it.
It is scheduled to run automatically every Sunday. 
"""
import shutil
import os
import pandas as pd
from openpyxl import load_workbook
import datetime
import schedule
import time

def merge_files():
    # move file from the "Current Period" directory to the current working directory
    for file in os.listdir("Current Period/"):
        shutil.move(os.getcwd() + '\\Current Period\\' + file, os.getcwd())
    # create a list of files in the current working directory
    files = os.listdir(os.getcwd())
    # create a data frame where files will be merged 
    df = pd.DataFrame()
    for file in files: 
        if file.endswith('.xlsx') and file != 'main.xlsx':
            df = df.append(pd.read_excel(file), ignore_index=True)
    # save data frame as an excel file
    df.to_excel(f"period_total_{datetime.date.today()}.xlsx", 
                sheet_name="Sheet1",
                index=False)
    # transfer data frame into the "Processed" directory
    # delete the files that were merged
    for file in files:
        if file.endswith('.xlsx') and file != 'main.xlsx':
            # os.remove(os.getcwd() + '\\' + file)
            shutil.move(os.getcwd() + '\\' + file, os.getcwd() + '\\Processed')
    try: 
        shutil.move(os.getcwd() + '\\' + f"period_total_{datetime.date.today()}.xlsx",
                    os.getcwd() + '\\Processed')
    except:
        os.remove(os.getcwd() + '\\' + f"period_total_{datetime.date.today()}.xlsx")
    return(df)
    
def update_workbook():
    df = merge_files()
    # load main workbook
    wb_name = 'main.xlsx'
    wb = load_workbook(wb_name)
    sheet = wb['Tabelle1']
    # add new data to the workbook 
    new_data = df.values.tolist()
    for i in new_data:
        sheet.append(i)
    # save the updated workbook
    wb.save(filename = wb_name)
    
    print(f"Workbook {wb_name} has been updated!")

# schedule the program to update the workbook every Sunday     
schedule.every().sunday.do(update_workbook)
print('The programme is scheduled to run every Sunday')
while True:
    schedule.run_pending()
    time.sleep(1)
