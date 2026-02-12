import os
import pandas as pd

dir_name = r"C:\Users\chika\\"
file_name = "2021切花売上.xlsx"

os.chdir(dir_name)
df = pd.read_excel(file_name,header=0,index_col=0,sheet_name='Sheet5')
df = df.cumsum()
print(df)

df.to_excel(dir_name+file_name, sheet_name='Sheet6')
