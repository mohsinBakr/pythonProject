import os

import pandas as pd
from pandasql import sqldf
import re

df = pd.read_excel('ddd.xlsx')
distinct_values = sqldf("SELECT DISTINCT SubjectId FROM df")
for value in distinct_values['SubjectId']:
    subjectName = sqldf(f"SELECT DISTINCT SubjectNameEn FROM df WHERE SubjectId = '{value}'")
    subjectName2 = subjectName.to_string(index=1).replace("  SubjectNameEn\n0    ", "").replace(" ", "")

    GradeId = sqldf(f"SELECT DISTINCT GradeId FROM df WHERE SubjectId = '{value}'")
    GradeId2 = GradeId.to_string(index=1).replace("GradeId\n0","").replace(" ", "")
    # print(GradeId2)

    LanguageType = sqldf(f"SELECT DISTINCT LanguageType FROM df WHERE SubjectId = '{value}'")
    LanguageType2 = LanguageType.to_string(index=1).replace("LanguageType\n0","").replace(" ", "")
    # print(LanguageType2)
    # print(subjectName2)

    selected_rows = sqldf(f"SELECT * FROM df WHERE SubjectId = '{value}'")

    if not os.path.exists("Subjects_Sheets"):
        os.makedirs("Subjects_Sheets")
    if not os.path.exists(f"Subjects_Sheets/Grade_{GradeId2}"):
        os.makedirs(f"Subjects_Sheets/Grade_{GradeId2}")
    if not os.path.exists(f"Subjects_Sheets/Grade_{GradeId2}/Language Type_{LanguageType2}"):
        os.makedirs(f"Subjects_Sheets/Grade_{GradeId2}/Language Type_{LanguageType2}")

    selected_rows.to_excel(f"Subjects_Sheets/Grade_{GradeId2}/Language Type_{LanguageType2}/{subjectName2}_Id:{value}.xlsx", index=False)

    print("Created:" + f"Subjects_Sheets/Grade_{GradeId2}/Language Type_{LanguageType2}/{subjectName2}_Id:{value}.xlsx")
    # print(subjectName)
