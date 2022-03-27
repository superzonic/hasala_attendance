import pandas as pd
import glob
import os


def summary(location):
    excel_files = glob.glob(location)

    df1 = pd.DataFrame()

    for excel_file in excel_files:
        df2 = pd.read_excel(excel_file, header=None)
        df1 = pd.concat([df1, df2], ignore_index=True, )

    df1.fillna(value="", inplace=True)
    directory_file = open("directory_attendance_file")
    directory_file_path = str(directory_file.read()) + "/attendance_files/"
    df1.to_excel(directory_file_path+"summary.xlsx", index=False)

