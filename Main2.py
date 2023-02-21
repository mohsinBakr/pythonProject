import pandas as pd
import docx
from docx import Document


# def read_doc_table () :
document = Document("/Users/mohsenbakr/Downloads/FGS/Gr3L - French.docx")
table_num = 2
nheader = 2
table = document.tables[table_num-1]
data = [[cell.text for cell in row.cells] for row in table.rows]
df = pd.DataFrame(data)
if nheader == 1:
    df = df.rename(columns=df.iloc[0]).drop(df.index[0]).reset_index(drop=True)
elif nheader == 2:
    outside_col, inside_col = df.iloc[0], df. iloc[1]
    hier_index = pd.MultiIndex.from_tuples(list(zip(outside_col,inside_col)))
    df = pd.DataFrame(data, columns=hier_index).drop(df.index[[0,1]]).reset_index(drop=True)
elif nheader > 2:
    print ("More than two headers not currently supported")
    # df = pd.DataFrame()
df = pd.DataFrame()['Seat num.']
print(df)
    # return df

