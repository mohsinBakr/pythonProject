# import glob
# import os
# import tkinter
# from tkinter import messagebox
# from tkinter import ttk
#
# import openpyxl
# import openpyxl as ox
import pandas as pd
# import pandasql as ps
# from openpyxl.styles import Protection, PatternFill
# from openpyxl.worksheet.protection import SheetProtection
# from pandasql import sqldf

import pandas as pd


def read_jmeter_csv(csv_file):
    try:
        df = pd.read_csv(csv_file)
        return df
    except FileNotFoundError:
        print("File not found.")
        return None
    except Exception as e:
        print("An error occurred:", str(e))
        return None


def main():
    # Path to your JMeter Aggregate Report CSV file
    csv_file = "D:/Performance/Projects/aggregate.csv"

    df = read_jmeter_csv(csv_file)
    if df is not None:
        print("Data read successfully:")
        print(df.head())  # Print the first few rows of the dataframe


if __name__ == "__main__":
    main()
