import pandas as pd
import os


folder_path = 'Subjects_Sheets/Term_1'
combined_data = pd.DataFrame()

for root, dirs, files in os.walk(folder_path):
    for file in files:
        if file.endswith('.xlsx'):
            file_path = os.path.join(root, file)
            df = pd.read_excel(file_path)
            combined_data = combined_data.append(df, ignore_index=True)
            print(file_path)

output_path = 'Combined_Data.xlsx'
combined_data.to_excel(output_path, index=False)