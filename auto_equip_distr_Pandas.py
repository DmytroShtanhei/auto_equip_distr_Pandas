"""
Script
for processing spreadsheet of files "Договір.xlsx" and "Групування.xlsx"
and creating a new file "Рознарядка.xlsx"
with spreadsheets "Договір", "Групування", "Рознарядка".
"""

import pandas as pd
import datetime

pd.set_option("expand_frame_repr", False)

df_grouping = pd.read_excel("Групування-КодиЛВУ-Позиції.xlsx", sheet_name="Групування", usecols="D, H, M")

# Fill out missing values for "Поз. в Договорі"
df_grouping = df_grouping.fillna(method="pad")

# Add information about Warehouses
df_lvu_warehouses = pd.read_excel("Групування-КодиЛВУ-Позиції.xlsx", sheet_name="КодиЛВУ", usecols="A, E")
df_grouping = pd.merge(df_grouping, df_lvu_warehouses, how="left", on="Код Заводу")

# Group rows by common criteria and calculate sums for each group
df_grouping = df_grouping.groupby(["Завод зберігання", "Поз. в Договорі", "Код Заводу"]).sum()
# print(df_grouping)

# Reshape table where "Код Заводу" will be a column, and replace NaNs with 0s
df_grouping = df_grouping.unstack("Код Заводу", fill_value=0)

# Rename columns names
# Read codes and names columns
df_lvu_names = pd.read_excel("Групування-КодиЛВУ-Позиції.xlsx", sheet_name="КодиЛВУ", usecols="A, C")
# Convert Dataframe to Series
s_lvu_names = pd.Series(df_lvu_names["Актуальна назва підрозділу"].values, index=df_lvu_names["Код Заводу"])
# Rename codes to names in column names
df_grouping = df_grouping.rename(columns=s_lvu_names)

# Get rid of indexes that were added after grouping by Warehouses and Products
df_grouping.reset_index(inplace=True)
df_grouping.columns = [' '.join(col).strip("Кільк.").strip() for col in df_grouping.columns.values]

# Add more info about Warehouses (at the end of DF)
df_warehouses_info = pd.read_excel("Групування-КодиЛВУ-Позиції.xlsx", sheet_name="КодиЛВУ", usecols="E:H")
df_warehouses_info.drop_duplicates(inplace=True)
df_grouping = pd.merge(df_grouping, df_warehouses_info, how="left", on="Завод зберігання")
# Move last three columns with info about Warehouses to second position
df_grouping = pd.concat(
    [df_grouping.iloc[:, :1], df_grouping.iloc[:, -3:], df_grouping.iloc[:, 1:-3]],
    axis=1)

# Add more info about Products
df_contract = pd.read_excel("Групування-КодиЛВУ-Позиції.xlsx", sheet_name="Позиції", usecols="A:C, E:F")
df_grouping = pd.merge(df_grouping, df_contract, how="left", on="Поз. в Договорі")
# Move last four columns with info about Products to fifth position
df_grouping = pd.concat(
    [df_grouping.iloc[:, :5], df_grouping.iloc[:, -4:], df_grouping.iloc[:, 5:-4]],
    axis=1)

df_grouping.drop(["Поз. в Договорі"], axis=1, inplace=True)

# Start index from 1
df_grouping.index += 1
# Insert column with index like numbers
df_grouping.insert(0, "N за/п", df_grouping.index)

col_list = list(df_grouping)[9:]
df_grouping.insert(7, "К-сть", df_grouping[col_list].sum(axis=1))
print(df_grouping)

# df_grouping.style.applymap(style_negative, props='color:red;').highlight_max(axis=0)

# Write to Excel file
df_grouping.to_excel(f"Рознарядка {datetime.datetime.now().strftime('%Y-%m-%d_T%H%M%S')}.xlsx", index=False)
