#!/usr/bin/env python
# coding: utf-8
import shutil
import tkinter
from tkinter import messagebox
from tkinter import ttk

# ### Importing packages

# In[13]:


# import necessary libraries
import pandas as pd
import os
import glob
import pandasql as ps
import jinja2
from docxtpl import DocxTemplate

# Define Term and Month


# in the folder -------------------------------------------------------------------------------------------------------
path = os.getcwd()
DataBase = glob.glob(os.path.join(path, "DataBase.xlsx"))


def Generate_Certificates():
    Term = Term_combobox.get()
    Month = Month_combobox.get()
    Grade = Grade_combobox.get()
    # loop over the list of csv files

    for f in DataBase:
        # read the csv file
        grades = pd.read_excel(f, sheet_name="Grades")
        students = pd.read_excel(f, sheet_name="Students")
        subjects = pd.read_excel(f, sheet_name="Subjects")
        exams = pd.read_excel(f, sheet_name="Exams")
        teachers = pd.read_excel(f, sheet_name="Teachers")

    # ## Term {Term} Month {Month}

    # In[2]:

    # KG1 new
    # Language
    if Grade == "KG1":

        query = f'''

        select st.Id "StudentId", st.StudentNameEn ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 'KG1' and st.LanguageType = 'Language'
        order by 1

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 'KG1'
        group by 1 ,2
        order by 1

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameEn', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('A', inplace=True)

        doc = DocxTemplate("Templates/Templates/KG_Term_Month_L.docx")
        path2 = os.path.join(path, f"Grade KG1 Term {Term} Month {Month} Language")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameEn': row['Q1-StudentNameEn'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1R': row['Q1-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'TS': row["('Total', 'Science')"]

                       }

            doc.render(context)
            doc.save(
                f"Grade KG1 Term {Term} Month {Month} Language/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameEn']}.docx")

            # Arabic

        query = f'''

        select st.Id "StudentId", st.StudentNameAr ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 'KG1' and st.LanguageType = 'Arabic'
        order by 1

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 'KG1'
        group by 1 ,2
        order by 1

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameAr', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('غ', inplace=True)

        doc = DocxTemplate("Templates/KG_Term_Month_A.docx")
        path2 = os.path.join(path, f"Grade KG1 Term {Term} Month {Month} Arabic")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameAr': row['Q1-StudentNameAr'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1R': row['Q1-Religion'],
                       # 'Q2R' : row['Q2-Religion'],
                       # 'Q3R' : row['Q3-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'TS': row["('Total', 'Science')"]

                       }

            doc.render(context)
            doc.save(
                f"Grade KG1 Term {Term} Month {Month} Arabic/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameAr']}.docx")

    # In[5]:

    # KG2 new
    # Language
    elif Grade == "KG2":

        query = f'''

        select st.Id "StudentId", st.StudentNameEn ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 'KG2' and st.LanguageType = 'Language'
        order by 1

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 'KG2'
        group by 1 ,2
        order by 1

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameEn', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('A', inplace=True)

        doc = DocxTemplate("Templates/KG_Term_Month_L.docx")
        path2 = os.path.join(path, f"Grade KG2 Term {Term} Month {Month} Language")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameEn': row['Q1-StudentNameEn'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1R': row['Q3-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'TS': row["('Total', 'Science')"]

                       }

            doc.render(context)
            doc.save(
                f"Grade KG2 Term {Term} Month {Month} Language/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameEn']}.docx")

        # Arabic

        query = f'''

        select st.Id "StudentId", st.StudentNameAr ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 'KG2' and st.LanguageType = 'Arabic'
        order by 1

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 'KG2'
        group by 1 ,2
        order by 1

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameAr', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('غ', inplace=True)

        doc = DocxTemplate("Templates/KG_Term_Month_A.docx")
        path2 = os.path.join(path, f"Grade KG2 Term {Term} Month {Month} Arabic")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameAr': row['Q1-StudentNameAr'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1R': row['Q3-Religion'],
                       # 'Q2R' : row['Q2-Religion'],
                       # 'Q3R' : row['Q3-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'TS': row["('Total', 'Science')"]

                       }

            doc.render(context)
            doc.save(
                f"Grade KG2 Term {Term} Month {Month} Arabic/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameAr']}.docx")

    # In[6]:

    # Grade 1 new
    # Language
    elif Grade == "1":

        query = f'''

        select st.Id "StudentId", st.StudentNameEn ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 1 and st.LanguageType = 'Language'
        order by 1

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 1
        group by 1 ,2
        order by 1

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameEn', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Participation = Pivot['Mark']['Participation'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()
        Behavior = Pivot['Mark']['Behavior'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Participation.columns = "Participation-" + Participation.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns
        Behavior.columns = "Behavior-" + Behavior.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Behavior b on Q1."Q1-StudentId" = b."Behavior-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Participation p on Q1."Q1-StudentId" = p."Participation-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('A', inplace=True)

        doc = DocxTemplate("Templates/3_Term_Month_L.docx")
        path2 = os.path.join(path, f"Grade 1 Term {Term} Month {Month} Language")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameEn': row['Q1-StudentNameEn'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'BA': row['Behavior-Arabic'],
                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'PA': row['Participation-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'BE': row['Behavior-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'PE': row['Participation-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'BM': row['Behavior-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'PM': row['Participation-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1I': row['Q1-ICT'],
                       # 'Q2I' : row['Q2-ICT'],
                       # 'Q3I' : row['Q3-ICT'],
                       'BI': row['Behavior-ICT'],
                       'EI': row['Evaluation-ICT'],
                       'AI': row['Attendance-ICT'],
                       'PI': row['Participation-ICT'],
                       'TI': row["('Total', 'ICT')"],

                       'Q1R': row['Q1-Religion'],
                       # 'Q2R' : row['Q2-Religion'],
                       # 'Q3R' : row['Q3-Religion'],
                       'BR': row['Behavior-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'PR': row['Participation-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'BS': row['Behavior-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'PS': row['Participation-Science'],
                       'TS': row["('Total', 'Science')"],

                       'Q1SL': row['Q1-Second Language'],
                       'Q2SL': row['Q2-Second Language'],
                       'Q3SL': row['Q3-Second Language'],
                       'BSL': row['Behavior-Second Language'],
                       'ESL': row['Evaluation-Second Language'],
                       'ASL': row['Attendance-Second Language'],
                       'PSL': row['Participation-Second Language'],
                       'TSL': row["('Total', 'Second Language')"],

                       }

            doc.render(context)
            doc.save(
                f"Grade 1 Term {Term} Month {Month} Language/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameEn']}.docx")

        # Arabic

        query = f'''

        select st.Id "StudentId", st.StudentNameAr ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 1 and st.LanguageType = 'Arabic'
        order by 1

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 1
        group by 1 ,2
        order by 1

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameAr', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Participation = Pivot['Mark']['Participation'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()
        Behavior = Pivot['Mark']['Behavior'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Participation.columns = "Participation-" + Participation.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns
        Behavior.columns = "Behavior-" + Behavior.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Behavior b on Q1."Q1-StudentId" = b."Behavior-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Participation p on Q1."Q1-StudentId" = p."Participation-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('غ', inplace=True)

        doc = DocxTemplate("Templates/3_Term_Month_A.docx")
        path2 = os.path.join(path, f"Grade 1 Term {Term} Month {Month} Arabic")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameAr': row['Q1-StudentNameAr'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'BA': row['Behavior-Arabic'],
                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'PA': row['Participation-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'BE': row['Behavior-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'PE': row['Participation-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'BM': row['Behavior-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'PM': row['Participation-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1I': row['Q1-ICT'],
                       # 'Q2I' : row['Q2-ICT'],
                       # 'Q3I' : row['Q3-ICT'],
                       'BI': row['Behavior-ICT'],
                       'EI': row['Evaluation-ICT'],
                       'AI': row['Attendance-ICT'],
                       'PI': row['Participation-ICT'],
                       'TI': row["('Total', 'ICT')"],

                       'Q1R': row['Q1-Religion'],
                       # 'Q2R' : row['Q2-Religion'],
                       # 'Q3R' : row['Q3-Religion'],
                       'BR': row['Behavior-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'PR': row['Participation-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'BS': row['Behavior-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'PS': row['Participation-Science'],
                       'TS': row["('Total', 'Science')"],

                       'Q1SL': row['Q1-Second Language'],
                       'Q2SL': row['Q2-Second Language'],
                       'Q3SL': row['Q3-Second Language'],
                       'BSL': row['Behavior-Second Language'],
                       'ESL': row['Evaluation-Second Language'],
                       'ASL': row['Attendance-Second Language'],
                       'PSL': row['Participation-Second Language'],
                       'TSL': row["('Total', 'Second Language')"],

                       }

            doc.render(context)
            doc.save(
                f"Grade 1 Term {Term} Month {Month} Arabic/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameAr']}.docx")

    # In[7]:

    # Grade 2 new

    # Language
    elif Grade == "2":

        query = f'''

        select st.Id "StudentId", st.StudentNameEn ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language' 
                     when sb.SubjectNameEn like 'Religion%'then 'Religion' 
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 2 and st.LanguageType = 'Language'
        order by 1 

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language' 
                     when sb.SubjectNameEn like 'Religion%'then 'Religion' 
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 2
        group by 1 ,2
        order by 1 

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameEn', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Participation = Pivot['Mark']['Participation'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()
        Behavior = Pivot['Mark']['Behavior'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Participation.columns = "Participation-" + Participation.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns
        Behavior.columns = "Behavior-" + Behavior.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Behavior b on Q1."Q1-StudentId" = b."Behavior-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Participation p on Q1."Q1-StudentId" = p."Participation-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('A', inplace=True)

        doc = DocxTemplate("Templates/3_Term_Month_L.docx")
        path2 = os.path.join(path, f"Grade 2 Term {Term} Month {Month} Language")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameEn': row['Q1-StudentNameEn'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'BA': row['Behavior-Arabic'],
                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'PA': row['Participation-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'BE': row['Behavior-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'PE': row['Participation-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'BM': row['Behavior-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'PM': row['Participation-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1I': row['Q1-ICT'],
                       # 'Q2I' : row['Q2-ICT'],
                       # 'Q3I' : row['Q3-ICT'],
                       'BI': row['Behavior-ICT'],
                       'EI': row['Evaluation-ICT'],
                       'AI': row['Attendance-ICT'],
                       'PI': row['Participation-ICT'],
                       'TI': row["('Total', 'ICT')"],

                       'Q1R': row['Q1-Religion'],
                       # 'Q2R' : row['Q2-Religion'],
                       # 'Q3R' : row['Q3-Religion'],
                       'BR': row['Behavior-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'PR': row['Participation-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'BS': row['Behavior-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'PS': row['Participation-Science'],
                       'TS': row["('Total', 'Science')"],

                       'Q1SL': row['Q1-Second Language'],
                       'Q2SL': row['Q2-Second Language'],
                       'Q3SL': row['Q3-Second Language'],
                       'BSL': row['Behavior-Second Language'],
                       'ESL': row['Evaluation-Second Language'],
                       'ASL': row['Attendance-Second Language'],
                       'PSL': row['Participation-Second Language'],
                       'TSL': row["('Total', 'Second Language')"],

                       }

            doc.render(context)
            doc.save(
                f"Grade 2 Term {Term} Month {Month} Language/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameEn']}.docx")

        # Arabic

        query = f'''

        select st.Id "StudentId", st.StudentNameAr ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language' 
                     when sb.SubjectNameEn like 'Religion%'then 'Religion' 
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 2 and st.LanguageType = 'Arabic'
        order by 1 

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language' 
                     when sb.SubjectNameEn like 'Religion%'then 'Religion' 
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 2
        group by 1 ,2
        order by 1 

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameAr', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Participation = Pivot['Mark']['Participation'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()
        Behavior = Pivot['Mark']['Behavior'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Participation.columns = "Participation-" + Participation.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns
        Behavior.columns = "Behavior-" + Behavior.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Behavior b on Q1."Q1-StudentId" = b."Behavior-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Participation p on Q1."Q1-StudentId" = p."Participation-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('غ', inplace=True)

        doc = DocxTemplate("Templates/3_Term_Month_A.docx")
        path2 = os.path.join(path, f"Grade 2 Term {Term} Month {Month} Arabic")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameAr': row['Q1-StudentNameAr'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'BA': row['Behavior-Arabic'],
                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'PA': row['Participation-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'BE': row['Behavior-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'PE': row['Participation-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'BM': row['Behavior-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'PM': row['Participation-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1I': row['Q1-ICT'],
                       # 'Q2I' : row['Q2-ICT'],
                       # 'Q3I' : row['Q3-ICT'],
                       'BI': row['Behavior-ICT'],
                       'EI': row['Evaluation-ICT'],
                       'AI': row['Attendance-ICT'],
                       'PI': row['Participation-ICT'],
                       'TI': row["('Total', 'ICT')"],

                       'Q1R': row['Q1-Religion'],
                       # 'Q2R' : row['Q2-Religion'],
                       # 'Q3R' : row['Q3-Religion'],
                       'BR': row['Behavior-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'PR': row['Participation-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'BS': row['Behavior-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'PS': row['Participation-Science'],
                       'TS': row["('Total', 'Science')"],

                       'Q1SL': row['Q1-Second Language'],
                       'Q2SL': row['Q2-Second Language'],
                       'Q3SL': row['Q3-Second Language'],
                       'BSL': row['Behavior-Second Language'],
                       'ESL': row['Evaluation-Second Language'],
                       'ASL': row['Attendance-Second Language'],
                       'PSL': row['Participation-Second Language'],
                       'TSL': row["('Total', 'Second Language')"],

                       }

            doc.render(context)
            doc.save(
                f"Grade 2 Term {Term} Month {Month} Arabic/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameAr']}.docx")

    # In[8]:

    # Grade 3 new
    elif Grade == "3":

        # Language

        query = f'''

        select st.Id "StudentId", st.StudentNameEn ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language' 
                     when sb.SubjectNameEn like 'Religion%'then 'Religion' 
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 3 and st.LanguageType = 'Language'
        order by 1 

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language' 
                     when sb.SubjectNameEn like 'Religion%'then 'Religion' 
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 3
        group by 1 ,2
        order by 1 

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameEn', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Participation = Pivot['Mark']['Participation'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()
        Behavior = Pivot['Mark']['Behavior'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Participation.columns = "Participation-" + Participation.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns
        Behavior.columns = "Behavior-" + Behavior.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Behavior b on Q1."Q1-StudentId" = b."Behavior-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Participation p on Q1."Q1-StudentId" = p."Participation-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('A', inplace=True)

        doc = DocxTemplate("Templates/3_Term_Month_L.docx")
        path2 = os.path.join(path, f"Grade 3 Term {Term} Month {Month} Language")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameEn': row['Q1-StudentNameEn'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'BA': row['Behavior-Arabic'],
                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'PA': row['Participation-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'BE': row['Behavior-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'PE': row['Participation-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'BM': row['Behavior-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'PM': row['Participation-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1I': row['Q1-ICT'],
                       # 'Q2I' : row['Q2-ICT'],
                       # 'Q3I' : row['Q3-ICT'],
                       'BI': row['Behavior-ICT'],
                       'EI': row['Evaluation-ICT'],
                       'AI': row['Attendance-ICT'],
                       'PI': row['Participation-ICT'],
                       'TI': row["('Total', 'ICT')"],

                       'Q1R': row['Q1-Religion'],
                       # 'Q2R' : row['Q2-Religion'],
                       # 'Q3R' : row['Q3-Religion'],
                       'BR': row['Behavior-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'PR': row['Participation-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'BS': row['Behavior-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'PS': row['Participation-Science'],
                       'TS': row["('Total', 'Science')"],

                       'Q1SL': row['Q1-Second Language'],
                       'Q2SL': row['Q2-Second Language'],
                       'Q3SL': row['Q3-Second Language'],
                       'BSL': row['Behavior-Second Language'],
                       'ESL': row['Evaluation-Second Language'],
                       'ASL': row['Attendance-Second Language'],
                       'PSL': row['Participation-Second Language'],
                       'TSL': row["('Total', 'Second Language')"],

                       }

            doc.render(context)
            doc.save(
                f"Grade 3 Term {Term} Month {Month} Language/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameEn']}.docx")

        # Arabic

        query = f'''

        select st.Id "StudentId", st.StudentNameAr ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language' 
                     when sb.SubjectNameEn like 'Religion%'then 'Religion' 
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 3 and st.LanguageType = 'Arabic'
        order by 1 

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language' 
                     when sb.SubjectNameEn like 'Religion%'then 'Religion' 
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 3
        group by 1 ,2
        order by 1 

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameAr', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Participation = Pivot['Mark']['Participation'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()
        Behavior = Pivot['Mark']['Behavior'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Participation.columns = "Participation-" + Participation.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns
        Behavior.columns = "Behavior-" + Behavior.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Behavior b on Q1."Q1-StudentId" = b."Behavior-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Participation p on Q1."Q1-StudentId" = p."Participation-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('غ', inplace=True)

        doc = DocxTemplate("Templates/3_Term_Month_A.docx")
        path2 = os.path.join(path, f"Grade 3 Term {Term} Month {Month} Arabic")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameAr': row['Q1-StudentNameAr'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'BA': row['Behavior-Arabic'],
                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'PA': row['Participation-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'BE': row['Behavior-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'PE': row['Participation-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'BM': row['Behavior-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'PM': row['Participation-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1I': row['Q1-ICT'],
                       # 'Q2I' : row['Q2-ICT'],
                       # 'Q3I' : row['Q3-ICT'],
                       'BI': row['Behavior-ICT'],
                       'EI': row['Evaluation-ICT'],
                       'AI': row['Attendance-ICT'],
                       'PI': row['Participation-ICT'],
                       'TI': row["('Total', 'ICT')"],

                       'Q1R': row['Q1-Religion'],
                       # 'Q2R' : row['Q2-Religion'],
                       # 'Q3R' : row['Q3-Religion'],
                       'BR': row['Behavior-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'PR': row['Participation-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'BS': row['Behavior-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'PS': row['Participation-Science'],
                       'TS': row["('Total', 'Science')"],

                       'Q1SL': row['Q1-Second Language'],
                       'Q2SL': row['Q2-Second Language'],
                       'Q3SL': row['Q3-Second Language'],
                       'BSL': row['Behavior-Second Language'],
                       'ESL': row['Evaluation-Second Language'],
                       'ASL': row['Attendance-Second Language'],
                       'PSL': row['Participation-Second Language'],
                       'TSL': row["('Total', 'Second Language')"],

                       }

            doc.render(context)
            doc.save(
                f"Grade 3 Term {Term} Month {Month} Arabic/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameAr']}.docx")

    # In[9]:

    # Grade 4 new
    elif Grade == "4":

        # Language

        query = f'''

        select st.Id "StudentId", st.StudentNameEn ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language' 
                     when sb.SubjectNameEn like 'Religion%'then 'Religion' 
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 4 and st.LanguageType = 'Language'
        order by 1 

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language' 
                     when sb.SubjectNameEn like 'Religion%'then 'Religion' 
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 4
        group by 1 ,2
        order by 1 

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameEn', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()
        Behavior = Pivot['Mark']['Behavior'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns
        Behavior.columns = "Behavior-" + Behavior.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Behavior b on Q1."Q1-StudentId" = b."Behavior-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('A', inplace=True)

        doc = DocxTemplate("Templates/4_Term_Month_L.docx")
        path2 = os.path.join(path, f"Grade 4 Term {Term} Month {Month} Language")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameEn': row['Q1-StudentNameEn'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'BA': row['Behavior-Arabic'],
                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'BE': row['Behavior-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'BM': row['Behavior-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1I': row['Q1-ICT'],
                       'Q2I': row['Q2-ICT'],
                       'Q3I': row['Q3-ICT'],
                       'BI': row['Behavior-ICT'],
                       'EI': row['Evaluation-ICT'],
                       'AI': row['Attendance-ICT'],
                       'TI': row["('Total', 'ICT')"],

                       'Q1R': row['Q1-Religion'],
                       # 'Q2R' : row['Q2-Religion'],
                       # 'Q3R' : row['Q3-Religion'],
                       'BR': row['Behavior-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'BS': row['Behavior-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'TS': row["('Total', 'Science')"],

                       'Q1SL': row['Q1-Second Language'],
                       'Q2SL': row['Q2-Second Language'],
                       'Q3SL': row['Q3-Second Language'],
                       'BSL': row['Behavior-Second Language'],
                       'ESL': row['Evaluation-Second Language'],
                       'ASL': row['Attendance-Second Language'],
                       'TSL': row["('Total', 'Second Language')"],

                       'Q1So': row['Q1-Social'],
                       'Q2So': row['Q2-Social'],
                       'Q3So': row['Q3-Social'],
                       'BSo': row['Behavior-Social'],
                       'ESo': row['Evaluation-Social'],
                       'ASo': row['Attendance-Social'],
                       'TSo': row["('Total', 'Social')"],

                       'Q1Sk': row['Q1-Skills'],
                       'Q2Sk': row['Q2-Skills'],
                       'Q3Sk': row['Q3-Skills'],
                       'BSk': row['Behavior-Skills'],
                       'ESk': row['Evaluation-Skills'],
                       'ASk': row['Attendance-Skills'],
                       'TSk': row["('Total', 'Skills')"]

                       }

            doc.render(context)
            doc.save(
                f"Grade 4 Term {Term} Month {Month} Language/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameEn']}.docx")

        # Arabic

        query = f'''

        select st.Id "StudentId", st.StudentNameAr ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language' 
                     when sb.SubjectNameEn like 'Religion%'then 'Religion' 
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 4 and st.LanguageType = 'Arabic'
        order by 1 

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language' 
                     when sb.SubjectNameEn like 'Religion%'then 'Religion' 
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 4
        group by 1 ,2
        order by 1 

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameAr', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()
        Behavior = Pivot['Mark']['Behavior'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns
        Behavior.columns = "Behavior-" + Behavior.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Behavior b on Q1."Q1-StudentId" = b."Behavior-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('غ', inplace=True)

        doc = DocxTemplate("Templates/4_Term_Month_A.docx")
        path2 = os.path.join(path, f"Grade 4 Term {Term} Month {Month} Arabic")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameAr': row['Q1-StudentNameAr'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'BA': row['Behavior-Arabic'],
                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'BE': row['Behavior-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],

                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'BM': row['Behavior-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1I': row['Q1-ICT'],
                       'Q2I': row['Q2-ICT'],
                       'Q3I': row['Q3-ICT'],
                       'BI': row['Behavior-ICT'],
                       'EI': row['Evaluation-ICT'],
                       'AI': row['Attendance-ICT'],
                       'TI': row["('Total', 'ICT')"],

                       'Q1R': row['Q1-Religion'],
                       # 'Q2R' : row['Q2-Religion'],
                       # 'Q3R' : row['Q3-Religion'],
                       'BR': row['Behavior-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'BS': row['Behavior-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'TS': row["('Total', 'Science')"],

                       'Q1SL': row['Q1-Second Language'],
                       'Q2SL': row['Q2-Second Language'],
                       'Q3SL': row['Q3-Second Language'],
                       'BSL': row['Behavior-Second Language'],
                       'ESL': row['Evaluation-Second Language'],
                       'ASL': row['Attendance-Second Language'],
                       'TSL': row["('Total', 'Second Language')"],

                       'Q1So': row['Q1-Social'],
                       'Q2So': row['Q2-Social'],
                       'Q3So': row['Q3-Social'],
                       'BSo': row['Behavior-Social'],
                       'ESo': row['Evaluation-Social'],
                       'ASo': row['Attendance-Social'],
                       'TSo': row["('Total', 'Social')"],

                       'Q1Sk': row['Q1-Skills'],
                       'Q2Sk': row['Q2-Skills'],
                       'Q3Sk': row['Q3-Skills'],
                       'BSk': row['Behavior-Skills'],
                       'ESk': row['Evaluation-Skills'],
                       'ASk': row['Attendance-Skills'],
                       'TSk': row["('Total', 'Skills')"]

                       }

            doc.render(context)
            doc.save(
                f"Grade 4 Term {Term} Month {Month} Arabic/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameAr']}.docx")

    # In[10]:

    # Grade 5 new
    elif Grade == "5":

        # Language

        query = f'''

        select st.Id "StudentId", st.StudentNameEn ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language' 
                     when sb.SubjectNameEn like 'Religion%'then 'Religion' 
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 5 and st.LanguageType = 'Language'
        order by 1 

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language' 
                     when sb.SubjectNameEn like 'Religion%'then 'Religion' 
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 5
        group by 1 ,2
        order by 1 

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameEn', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()
        Behavior = Pivot['Mark']['Behavior'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns
        Behavior.columns = "Behavior-" + Behavior.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Behavior b on Q1."Q1-StudentId" = b."Behavior-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('A', inplace=True)

        doc = DocxTemplate("Templates/5_Term_Month_L.docx")
        path2 = os.path.join(path, f"Grade 5 Term {Term} Month {Month} Language")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameEn': row['Q1-StudentNameEn'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'BA': row['Behavior-Arabic'],
                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'BE': row['Behavior-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'BM': row['Behavior-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1I': row['Q1-ICT'],
                       'Q2I': row['Q2-ICT'],
                       'Q3I': row['Q3-ICT'],
                       'BI': row['Behavior-ICT'],
                       'EI': row['Evaluation-ICT'],
                       'AI': row['Attendance-ICT'],
                       'TI': row["('Total', 'ICT')"],

                       'Q1R': row['Q1-Religion'],
                       # 'Q2R' : row['Q2-Religion'],
                       # 'Q3R' : row['Q3-Religion'],
                       'BR': row['Behavior-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'BS': row['Behavior-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'TS': row["('Total', 'Science')"],

                       'Q1SL': row['Q1-Second Language'],
                       'Q2SL': row['Q2-Second Language'],
                       'Q3SL': row['Q3-Second Language'],
                       'BSL': row['Behavior-Second Language'],
                       'ESL': row['Evaluation-Second Language'],
                       'ASL': row['Attendance-Second Language'],
                       'TSL': row["('Total', 'Second Language')"],

                       'Q1So': row['Q1-Social'],
                       'Q2So': row['Q2-Social'],
                       'Q3So': row['Q3-Social'],
                       'BSo': row['Behavior-Social'],
                       'ESo': row['Evaluation-Social'],
                       'ASo': row['Attendance-Social'],
                       'TSo': row["('Total', 'Social')"],

                       'Q1Sk': row['Q1-Skills'],
                       'Q2Sk': row['Q2-Skills'],
                       'Q3Sk': row['Q3-Skills'],
                       'BSk': row['Behavior-Skills'],
                       'ESk': row['Evaluation-Skills'],
                       'ASk': row['Attendance-Skills'],
                       'TSk': row["('Total', 'Skills')"]

                       }

            doc.render(context)
            doc.save(
                f"Grade 5 Term {Term} Month {Month} Language/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameEn']}.docx")

        # Arabic

        query = f'''

        select st.Id "StudentId", st.StudentNameAr ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language' 
                     when sb.SubjectNameEn like 'Religion%'then 'Religion' 
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 5 and st.LanguageType = 'Arabic'
        order by 1 

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language' 
                     when sb.SubjectNameEn like 'Religion%'then 'Religion' 
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 5
        group by 1 ,2
        order by 1 

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameAr', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()
        Behavior = Pivot['Mark']['Behavior'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns
        Behavior.columns = "Behavior-" + Behavior.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Behavior b on Q1."Q1-StudentId" = b."Behavior-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('غ', inplace=True)

        doc = DocxTemplate("Templates/5_Term_Month_A.docx")
        path2 = os.path.join(path, f"Grade 5 Term {Term} Month {Month} Arabic")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameAr': row['Q1-StudentNameAr'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'BA': row['Behavior-Arabic'],
                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'BE': row['Behavior-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'BM': row['Behavior-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1I': row['Q1-ICT'],
                       'Q2I': row['Q2-ICT'],
                       'Q3I': row['Q3-ICT'],
                       'BI': row['Behavior-ICT'],
                       'EI': row['Evaluation-ICT'],
                       'AI': row['Attendance-ICT'],
                       'TI': row["('Total', 'ICT')"],

                       'Q1R': row['Q1-Religion'],
                       # 'Q2R' : row['Q2-Religion'],
                       # 'Q3R' : row['Q3-Religion'],
                       'BR': row['Behavior-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'BS': row['Behavior-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'TS': row["('Total', 'Science')"],

                       'Q1SL': row['Q1-Second Language'],
                       'Q2SL': row['Q2-Second Language'],
                       'Q3SL': row['Q3-Second Language'],
                       'BSL': row['Behavior-Second Language'],
                       'ESL': row['Evaluation-Second Language'],
                       'ASL': row['Attendance-Second Language'],
                       'TSL': row["('Total', 'Second Language')"],

                       'Q1So': row['Q1-Social'],
                       'Q2So': row['Q2-Social'],
                       'Q3So': row['Q3-Social'],
                       'BSo': row['Behavior-Social'],
                       'ESo': row['Evaluation-Social'],
                       'ASo': row['Attendance-Social'],
                       'TSo': row["('Total', 'Social')"],

                       'Q1Sk': row['Q1-Skills'],
                       'Q2Sk': row['Q2-Skills'],
                       'Q3Sk': row['Q3-Skills'],
                       'BSk': row['Behavior-Skills'],
                       'ESk': row['Evaluation-Skills'],
                       'ASk': row['Attendance-Skills'],
                       'TSk': row["('Total', 'Skills')"]

                       }

            doc.render(context)
            doc.save(
                f"Grade 5 Term {Term} Month {Month} Arabic/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameAr']}.docx")

    # In[15]:

    # Grade 6 new
    elif Grade == "6":

        # Language

        query = f'''

        select st.Id "StudentId", st.StudentNameEn ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 6 and st.LanguageType = 'Language'
        order by 1

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 6
        group by 1 ,2
        order by 1

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameEn', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Participation = Pivot['Mark']['Participation'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()
        Behavior = Pivot['Mark']['Behavior'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Participation.columns = "Participation-" + Participation.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns
        Behavior.columns = "Behavior-" + Behavior.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Behavior b on Q1."Q1-StudentId" = b."Behavior-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Participation p on Q1."Q1-StudentId" = p."Participation-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('A', inplace=True)

        doc = DocxTemplate("Templates/6_Term_Month_L.docx")
        path2 = os.path.join(path, f"Grade 6 Term {Term} Month {Month} Language")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameEn': row['Q1-StudentNameEn'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'BA': row['Behavior-Arabic'],
                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'PA': row['Participation-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'BE': row['Behavior-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'PE': row['Participation-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'BM': row['Behavior-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'PM': row['Participation-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1I': row['Q1-ICT'],
                       'Q2I': row['Q2-ICT'],
                       'Q3I': row['Q3-ICT'],
                       'BI': row['Behavior-ICT'],
                       'EI': row['Evaluation-ICT'],
                       'AI': row['Attendance-ICT'],
                       'PI': row['Participation-ICT'],
                       'TI': row["('Total', 'ICT')"],

                       'Q1R': row['Q1-Religion'],
                       # 'Q2R' : row['Q2-Religion'],
                       # 'Q3R' : row['Q3-Religion'],
                       'BR': row['Behavior-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'PR': row['Participation-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'BS': row['Behavior-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'PS': row['Participation-Science'],
                       'TS': row["('Total', 'Science')"],

                       'Q1SL': row['Q1-Second Language'],
                       'Q2SL': row['Q2-Second Language'],
                       'Q3SL': row['Q3-Second Language'],
                       'BSL': row['Behavior-Second Language'],
                       'ESL': row['Evaluation-Second Language'],
                       'ASL': row['Attendance-Second Language'],
                       'PSL': row['Participation-Second Language'],
                       'TSL': row["('Total', 'Second Language')"],

                       'Q1So': row['Q1-Social'],
                       'Q2So': row['Q2-Social'],
                       'Q3So': row['Q3-Social'],
                       'BSo': row['Behavior-Social'],
                       'ESo': row['Evaluation-Social'],
                       'ASo': row['Attendance-Social'],
                       'PSo': row['Participation-Social'],
                       'TSo': row["('Total', 'Social')"]

                       }

            doc.render(context)
            doc.save(
                f"Grade 6 Term {Term} Month {Month} Language/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameEn']}.docx")

        # Arabic

        query = f'''

        select st.Id "StudentId", st.StudentNameAr ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 6 and st.LanguageType = 'Arabic'
        order by 1

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 6
        group by 1 ,2
        order by 1

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameAr', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Participation = Pivot['Mark']['Participation'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()
        Behavior = Pivot['Mark']['Behavior'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Participation.columns = "Participation-" + Participation.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns
        Behavior.columns = "Behavior-" + Behavior.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Behavior b on Q1."Q1-StudentId" = b."Behavior-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Participation p on Q1."Q1-StudentId" = p."Participation-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('غ', inplace=True)

        doc = DocxTemplate("Templates/6_Term_Month_A.docx")
        path2 = os.path.join(path, f"Grade 6 Term {Term} Month {Month} Arabic")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameAr': row['Q1-StudentNameAr'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'BA': row['Behavior-Arabic'],
                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'PA': row['Participation-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'BE': row['Behavior-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'PE': row['Participation-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'BM': row['Behavior-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'PM': row['Participation-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1I': row['Q1-ICT'],
                       'Q2I': row['Q2-ICT'],
                       'Q3I': row['Q3-ICT'],
                       'BI': row['Behavior-ICT'],
                       'EI': row['Evaluation-ICT'],
                       'AI': row['Attendance-ICT'],
                       'PI': row['Participation-ICT'],
                       'TI': row["('Total', 'ICT')"],

                       'Q1R': row['Q1-Religion'],
                       # 'Q2R' : row['Q2-Religion'],
                       # 'Q3R' : row['Q3-Religion'],
                       'BR': row['Behavior-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'PR': row['Participation-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'BS': row['Behavior-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'PS': row['Participation-Science'],
                       'TS': row["('Total', 'Science')"],

                       'Q1SL': row['Q1-Second Language'],
                       'Q2SL': row['Q2-Second Language'],
                       'Q3SL': row['Q3-Second Language'],
                       'BSL': row['Behavior-Second Language'],
                       'ESL': row['Evaluation-Second Language'],
                       'ASL': row['Attendance-Second Language'],
                       'PSL': row['Participation-Second Language'],
                       'TSL': row["('Total', 'Second Language')"],

                       'Q1So': row['Q1-Social'],
                       'Q2So': row['Q2-Social'],
                       'Q3So': row['Q3-Social'],
                       'BSo': row['Behavior-Social'],
                       'ESo': row['Evaluation-Social'],
                       'ASo': row['Attendance-Social'],
                       'PSo': row['Participation-Social'],
                       'TSo': row["('Total', 'Social')"]

                       }

            doc.render(context)
            doc.save(
                f"Grade 6 Term {Term} Month {Month} Arabic/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameAr']}.docx")

    # In[16]:

    # Grade 7 new
    elif Grade == "7":

        # Language

        query = f'''

        select st.Id "StudentId", st.StudentNameEn ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 7 and st.LanguageType = 'Language'
        order by 1

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 7
        group by 1 ,2
        order by 1

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameEn', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Participation = Pivot['Mark']['Participation'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()
        Behavior = Pivot['Mark']['Behavior'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Participation.columns = "Participation-" + Participation.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns
        Behavior.columns = "Behavior-" + Behavior.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Behavior b on Q1."Q1-StudentId" = b."Behavior-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Participation p on Q1."Q1-StudentId" = p."Participation-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('A', inplace=True)

        doc = DocxTemplate("Templates/6_Term_Month_L.docx")
        path2 = os.path.join(path, f"Grade 7 Term {Term} Month {Month} Language")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameEn': row['Q1-StudentNameEn'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'BA': row['Behavior-Arabic'],
                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'PA': row['Participation-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'BE': row['Behavior-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'PE': row['Participation-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'BM': row['Behavior-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'PM': row['Participation-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1I': row['Q1-ICT'],
                       'Q2I': row['Q2-ICT'],
                       'Q3I': row['Q3-ICT'],
                       'BI': row['Behavior-ICT'],
                       'EI': row['Evaluation-ICT'],
                       'AI': row['Attendance-ICT'],
                       'PI': row['Participation-ICT'],
                       'TI': row["('Total', 'ICT')"],

                       'Q1R': row['Q1-Religion'],
                       # 'Q2R' : row['Q2-Religion'],
                       # 'Q3R' : row['Q3-Religion'],
                       'BR': row['Behavior-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'PR': row['Participation-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'BS': row['Behavior-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'PS': row['Participation-Science'],
                       'TS': row["('Total', 'Science')"],

                       'Q1SL': row['Q1-Second Language'],
                       'Q2SL': row['Q2-Second Language'],
                       'Q3SL': row['Q3-Second Language'],
                       'BSL': row['Behavior-Second Language'],
                       'ESL': row['Evaluation-Second Language'],
                       'ASL': row['Attendance-Second Language'],
                       'PSL': row['Participation-Second Language'],
                       'TSL': row["('Total', 'Second Language')"],

                       'Q1So': row['Q1-Social'],
                       'Q2So': row['Q2-Social'],
                       'Q3So': row['Q3-Social'],
                       'BSo': row['Behavior-Social'],
                       'ESo': row['Evaluation-Social'],
                       'ASo': row['Attendance-Social'],
                       'PSo': row['Participation-Social'],
                       'TSo': row["('Total', 'Social')"]

                       }

            doc.render(context)
            doc.save(
                f"Grade 7 Term {Term} Month {Month} Language/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameEn']}.docx")

        # Arabic

        query = f'''

        select st.Id "StudentId", st.StudentNameAr ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 7 and st.LanguageType = 'Arabic'
        order by 1

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 7
        group by 1 ,2
        order by 1

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameAr', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Participation = Pivot['Mark']['Participation'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()
        Behavior = Pivot['Mark']['Behavior'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Participation.columns = "Participation-" + Participation.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns
        Behavior.columns = "Behavior-" + Behavior.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Behavior b on Q1."Q1-StudentId" = b."Behavior-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Participation p on Q1."Q1-StudentId" = p."Participation-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('غ', inplace=True)

        doc = DocxTemplate("Templates/6_Term_Month_A.docx")
        path2 = os.path.join(path, f"Grade 7 Term {Term} Month {Month} Arabic")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameAr': row['Q1-StudentNameAr'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'BA': row['Behavior-Arabic'],
                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'PA': row['Participation-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'BE': row['Behavior-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'PE': row['Participation-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'BM': row['Behavior-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'PM': row['Participation-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1I': row['Q1-ICT'],
                       'Q2I': row['Q2-ICT'],
                       'Q3I': row['Q3-ICT'],
                       'BI': row['Behavior-ICT'],
                       'EI': row['Evaluation-ICT'],
                       'AI': row['Attendance-ICT'],
                       'PI': row['Participation-ICT'],
                       'TI': row["('Total', 'ICT')"],

                       'Q1R': row['Q1-Religion'],
                       # 'Q2R' : row['Q2-Religion'],
                       # 'Q3R' : row['Q3-Religion'],
                       'BR': row['Behavior-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'PR': row['Participation-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'BS': row['Behavior-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'PS': row['Participation-Science'],
                       'TS': row["('Total', 'Science')"],

                       'Q1SL': row['Q1-Second Language'],
                       'Q2SL': row['Q2-Second Language'],
                       'Q3SL': row['Q3-Second Language'],
                       'BSL': row['Behavior-Second Language'],
                       'ESL': row['Evaluation-Second Language'],
                       'ASL': row['Attendance-Second Language'],
                       'PSL': row['Participation-Second Language'],
                       'TSL': row["('Total', 'Second Language')"],

                       'Q1So': row['Q1-Social'],
                       'Q2So': row['Q2-Social'],
                       'Q3So': row['Q3-Social'],
                       'BSo': row['Behavior-Social'],
                       'ESo': row['Evaluation-Social'],
                       'ASo': row['Attendance-Social'],
                       'PSo': row['Participation-Social'],
                       'TSo': row["('Total', 'Social')"]

                       }

            doc.render(context)
            doc.save(
                f"Grade 7 Term {Term} Month {Month} Arabic/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameAr']}.docx")

    # In[17]:

    # Grade 8 new
    elif Grade == "8":

        # Language

        query = f'''

        select st.Id "StudentId", st.StudentNameEn ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 8 and st.LanguageType = 'Language'
        order by 1

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 8
        group by 1 ,2
        order by 1

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameEn', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Participation = Pivot['Mark']['Participation'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()
        Behavior = Pivot['Mark']['Behavior'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Participation.columns = "Participation-" + Participation.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns
        Behavior.columns = "Behavior-" + Behavior.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Behavior b on Q1."Q1-StudentId" = b."Behavior-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Participation p on Q1."Q1-StudentId" = p."Participation-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('A', inplace=True)

        doc = DocxTemplate("Templates/6_Term_Month_L.docx")
        path2 = os.path.join(path, f"Grade 8 Term {Term} Month {Month} Language")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameEn': row['Q1-StudentNameEn'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'BA': row['Behavior-Arabic'],
                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'PA': row['Participation-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'BE': row['Behavior-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'PE': row['Participation-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'BM': row['Behavior-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'PM': row['Participation-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1I': row['Q1-ICT'],
                       'Q2I': row['Q2-ICT'],
                       'Q3I': row['Q3-ICT'],
                       'BI': row['Behavior-ICT'],
                       'EI': row['Evaluation-ICT'],
                       'AI': row['Attendance-ICT'],
                       'PI': row['Participation-ICT'],
                       'TI': row["('Total', 'ICT')"],

                       'Q1R': row['Q1-Religion'],
                       # 'Q2R' : row['Q2-Religion'],
                       # 'Q3R' : row['Q3-Religion'],
                       'BR': row['Behavior-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'PR': row['Participation-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'BS': row['Behavior-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'PS': row['Participation-Science'],
                       'TS': row["('Total', 'Science')"],

                       'Q1SL': row['Q1-Second Language'],
                       'Q2SL': row['Q2-Second Language'],
                       'Q3SL': row['Q3-Second Language'],
                       'BSL': row['Behavior-Second Language'],
                       'ESL': row['Evaluation-Second Language'],
                       'ASL': row['Attendance-Second Language'],
                       'PSL': row['Participation-Second Language'],
                       'TSL': row["('Total', 'Second Language')"],

                       'Q1So': row['Q1-Social'],
                       'Q2So': row['Q2-Social'],
                       'Q3So': row['Q3-Social'],
                       'BSo': row['Behavior-Social'],
                       'ESo': row['Evaluation-Social'],
                       'ASo': row['Attendance-Social'],
                       'PSo': row['Participation-Social'],
                       'TSo': row["('Total', 'Social')"]

                       }

            doc.render(context)
            doc.save(
                f"Grade 8 Term {Term} Month {Month} Language/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameEn']}.docx")

        # Arabic

        query = f'''

        select st.Id "StudentId", st.StudentNameAr ,st.GradeId,st.Class, st.Religion, st.SecondLanguage, sb.Id "SubjectId"
              , case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,e.QuizType, e.Term, e.Month,e.Mark
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 8 and st.LanguageType = 'Arabic'
        order by 1

        '''
        query_output = ps.sqldf(query)

        Totals = f'''

        select st.Id "StudentId",case when sb.SubjectNameEn = 'German' or sb.SubjectNameEn = 'French' then 'Second Language'
                     when sb.SubjectNameEn like 'Religion%'then 'Religion'
                     else sb.SubjectNameEn end "SubjectNameEn"
              ,sum(Mark) "Total"
        from students st
        left join subjects sb on st.GradeId = sb.GradeId and st.LanguageType = sb.LanguageType and st.Class = sb.Class
        left join exams e on sb.Id = e.SubjectId and e.StudentId = st.Id
        Where Term = {Term} and Month = {Month} and st.GradeId = 8
        group by 1 ,2
        order by 1

        '''

        Totals = ps.sqldf(Totals)
        Totals = Totals.pivot_table(index=['StudentId'], columns=['SubjectNameEn'], values=['Total']).reset_index()
        Total = '''
        select * from Totals
        '''
        Total = ps.sqldf(Total)

        Pivot = query_output.pivot_table(index=['StudentId', 'StudentNameAr', 'GradeId', 'Class', 'Term', 'Month'],
                                         columns=['QuizType', 'SubjectNameEn'], values=['Mark'])
        Q1 = Pivot['Mark']['Q1'].reset_index()
        Q2 = Pivot['Mark']['Q2'].reset_index()
        Q3 = Pivot['Mark']['Q3'].reset_index()
        Participation = Pivot['Mark']['Participation'].reset_index()
        Attendance = Pivot['Mark']['Attendance'].reset_index()
        Evaluation = Pivot['Mark']['Evaluation'].reset_index()
        Behavior = Pivot['Mark']['Behavior'].reset_index()

        Q1.columns = "Q1-" + Q1.columns
        Q2.columns = "Q2-" + Q2.columns
        Q3.columns = "Q3-" + Q3.columns
        Participation.columns = "Participation-" + Participation.columns
        Attendance.columns = "Attendance-" + Attendance.columns
        Evaluation.columns = "Evaluation-" + Evaluation.columns
        Behavior.columns = "Behavior-" + Behavior.columns

        Table = ''' select *
        from Q1
        left join Q2 on Q1."Q1-StudentId" = Q2."Q2-StudentId"
        left join Q3 on Q1."Q1-StudentId"= Q3."Q3-StudentId"
        left join Behavior b on Q1."Q1-StudentId" = b."Behavior-StudentId"
        left join Evaluation e on Q1."Q1-StudentId" = e."Evaluation-StudentId"
        left join Attendance a on Q1."Q1-StudentId" = a."Attendance-StudentId"
        left join Participation p on Q1."Q1-StudentId" = p."Participation-StudentId"
        left join Total t on Q1."Q1-StudentId" = t."('StudentId', '')"
        '''
        Table = ps.sqldf(Table)
        Table.fillna('غ', inplace=True)

        doc = DocxTemplate("Templates/6_Term_Month_A.docx")
        path2 = os.path.join(path, f"Grade 8 Term {Term} Month {Month} Arabic")
        if os.path.exists(path2):
            # Remove existing directory
            shutil.rmtree(path2)
        os.mkdir(path2, 0o700)

        for index, row in Table.iterrows():
            print(f"Running Student ID: {row['Q1-StudentId']} if failed and this is last Student, then check its Marks")
            context = {'Id': row['Q1-StudentId'],
                       'Term': row['Q1-Term'],
                       'Month': row['Q1-Month'],
                       'StudentNameAr': row['Q1-StudentNameAr'],
                       'GradeId': row['Q1-GradeId'],
                       'Class': row['Q1-Class'],
                       'Q1A': row['Q1-Arabic'],
                       'Q2A': row['Q2-Arabic'],
                       'Q3A': row['Q3-Arabic'],

                       'BA': row['Behavior-Arabic'],
                       'EA': row['Evaluation-Arabic'],
                       'AA': row['Attendance-Arabic'],
                       'PA': row['Participation-Arabic'],
                       'TA': row["('Total', 'Arabic')"],

                       'Q1E': row['Q1-English'],
                       'Q2E': row['Q2-English'],
                       'Q3E': row['Q3-English'],
                       'BE': row['Behavior-English'],
                       'EE': row['Evaluation-English'],
                       'AE': row['Attendance-English'],
                       'PE': row['Participation-English'],
                       'TE': row["('Total', 'English')"],

                       'EE2': row['Evaluation-English2'],

                       'Q1M': row['Q1-Math'],
                       'Q2M': row['Q2-Math'],
                       'Q3M': row['Q3-Math'],
                       'BM': row['Behavior-Math'],
                       'EM': row['Evaluation-Math'],
                       'AM': row['Attendance-Math'],
                       'PM': row['Participation-Math'],
                       'TM': row["('Total', 'Math')"],

                       'Q1I': row['Q1-ICT'],
                       'Q2I': row['Q2-ICT'],
                       'Q3I': row['Q3-ICT'],
                       'BI': row['Behavior-ICT'],
                       'EI': row['Evaluation-ICT'],
                       'AI': row['Attendance-ICT'],
                       'PI': row['Participation-ICT'],
                       'TI': row["('Total', 'ICT')"],

                       'Q1R': row['Q1-Religion'],
                       # 'Q2R' : row['Q2-Religion'],
                       # 'Q3R' : row['Q3-Religion'],
                       'BR': row['Behavior-Religion'],
                       'ER': row['Evaluation-Religion'],
                       'AR': row['Attendance-Religion'],
                       'PR': row['Participation-Religion'],
                       'TR': row["('Total', 'Religion')"],

                       'Q1S': row['Q1-Science'],
                       'Q2S': row['Q2-Science'],
                       'Q3S': row['Q3-Science'],
                       'BS': row['Behavior-Science'],
                       'ES': row['Evaluation-Science'],
                       'AS': row['Attendance-Science'],
                       'PS': row['Participation-Science'],
                       'TS': row["('Total', 'Science')"],

                       'Q1SL': row['Q1-Second Language'],
                       'Q2SL': row['Q2-Second Language'],
                       'Q3SL': row['Q3-Second Language'],
                       'BSL': row['Behavior-Second Language'],
                       'ESL': row['Evaluation-Second Language'],
                       'ASL': row['Attendance-Second Language'],
                       'PSL': row['Participation-Second Language'],
                       'TSL': row["('Total', 'Second Language')"],

                       'Q1So': row['Q1-Social'],
                       'Q2So': row['Q2-Social'],
                       'Q3So': row['Q3-Social'],
                       'BSo': row['Behavior-Social'],
                       'ESo': row['Evaluation-Social'],
                       'ASo': row['Attendance-Social'],
                       'PSo': row['Participation-Social'],
                       'TSo': row["('Total', 'Social')"]

                       }

            doc.render(context)
            doc.save(
                f"Grade 8 Term {Term} Month {Month} Arabic/ID-{row['Q1-StudentId']} Grade-{row['Q1-GradeId']} Class-{row['Q1-Class']} {row['Q1-StudentNameAr']}.docx")
    else:
        tkinter.messagebox.showerror(title="Error", message="Grade Not Selected")

    tkinter.messagebox.showinfo(title='Process Completed',
                                message='Certificates Created Successfully')

    # In[ ]:


# User Interface--------------------------------------------------------------------------------------------------------
window = tkinter.Tk()
tabControl = ttk.Notebook(window)
tab1 = ttk.Frame(tabControl)

tabControl.add(tab1, text='Certificates Generator')
tabControl.pack(expand=1, fill="both")

window.title("Generate Certificates per Grade")
frame = tkinter.Frame(window)
frame.pack()
Hint_label = tkinter.Label(tab1,
                           text="Make sure that you select Term, Month and Grade that has correct data in DataBase.xlsx file\nCertificates templates and Database.xlsx should be added to the following path" + os.getcwd())
Hint_label.grid(row=0, column=0)

Term_label = tkinter.Label(tab1, text="Term")
Term_combobox = ttk.Combobox(tab1, values=["1", "2"], state="readonly")
Term_label.grid(row=1, column=0)
Term_combobox.grid(row=2, column=0)

Month_label = tkinter.Label(tab1, text="Month")
Month_combobox = ttk.Combobox(tab1, values=["1", "2", "3"], state="readonly")
Month_label.grid(row=3, column=0)
Month_combobox.grid(row=4, column=0)

Grade_label = tkinter.Label(tab1, text="Grade")
Grade_combobox = ttk.Combobox(tab1, values=["KG1", "KG2", "1", "2", "3", "4", "5", "6", "7", "8"], state="readonly")
Grade_label.grid(row=5, column=0)
Grade_combobox.grid(row=6, column=0)

Create_Final_Marks_sheet_button = tkinter.Button(tab1, text="Generate Certificates", command=Generate_Certificates,
                                                 state="normal")
Create_Final_Marks_sheet_button.grid(row=7, column=0, sticky="news", padx=20, pady=10)

Notes_label = tkinter.Label(tab1,
                            text="You can find generated certificates in the following path" + os.getcwd() + "/#Certificates Folder#")
Notes_label.grid(row=8, column=0)

window.mainloop()