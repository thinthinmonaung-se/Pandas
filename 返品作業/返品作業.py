返品表 = Excel.Cs_OpenV3(self, "C:\\Python Learning\\Work\\品質不良集計場所\\品質不良集計場所 - RPA用\\返品8月(2024)　-　新マクロ.xlsm", 1, "MicrosoftExcel", 1, 1, "", "", 0, 0, var_ret=0, skip_err=0, delay_before=0, delay_after=1000) 
#@_コードを挿入する "年月" 
import datetime

# Get current month number (1 to 12)
月 = datetime.datetime.now().month

# Dictionary for month mapping
month = {
    1: "10月",  2: "11月", 3: "12月", 4: "1月", 5: "2月",
    6: "3月",  7: "4月",  8: "5月",  9: "6月",  10: "7月",
    11: "8月", 12: "9月"
}

# Determine the search_year based on the mapping
current_year = datetime.datetime.now().year
if 月 in [1, 2, 3]:  
    search_year = current_year - 1
else:
    search_year = current_year

col_month = {
    1: "K",  2: "I",  3: "Dec-24",  4: "B",  5: "C", 
    6: "D",  7: "E",  8: "F",  9: "G",  10: "H", 
    11: "I", 12: "J"
}
 
#@ 
#@_コードを挿入する "Main df 作成" 
import pandas as pd

# Read the data from the Excel file
df = pd.read_excel(返品表, sheet_name=month[月], header=2, dtype=str).fillna('')
df.index = range(4, len(df) + 4)

# Remove alphabets until digits are found in the '商品ｺｰﾄﾞ' column
df['商品ｺｰﾄﾞ'] = df['商品ｺｰﾄﾞ'].str.replace(r'(?<=\d)[A-Za-z]+$', '', regex=True)

# Update the column based on the condition
#condition = (df["処理"] == "サンプル") & (df["検品後ｸﾚｰﾑ番号"] != "201")
#df.loc[condition, "検品後ｸﾚｰﾑ番号"] = "正常"

#(~df['返品先'].str.contains("⇒", na=False))]


# Filter the DataFrame with the given conditions
df_with_706 = df[
    ((df['検品後ｸﾚｰﾑ番号'] < "700") | (df['検品後ｸﾚｰﾑ番号'] == "706")) & 
    (df['検品後ｸﾚｰﾑ番号'] != "610") &         
    (df['検品後ｸﾚｰﾑ番号'] != "609") &        
    (df['検品後ｸﾚｰﾑ番号'] != "608") &        
    (df['検品後ｸﾚｰﾑ番号'] != "正常") 
]

df_without_706 = df[
    (df['検品後ｸﾚｰﾑ番号'] < "700") &  
    (df['検品後ｸﾚｰﾑ番号'] != "610") &         
    (df['検品後ｸﾚｰﾑ番号'] != "609") &        
    (df['検品後ｸﾚｰﾑ番号'] != "608") &        
    (df['検品後ｸﾚｰﾑ番号'] != "正常")
]

# List of sheet names
品質課題_sheet_names = ['品質課題', '品質課題_rpa']
p_table_sheet_names = ['p_table_rpa', 'p_table1_rpa']
df_names = [df_with_706, df_without_706]
 
#@ 
# #@_折りたたみ"売上仕入実績表" 
# 売上仕入実績表 = Excel.Cs_OpenV3(self, "C:\\Python Learning\\Work\\品質不良集計場所\\品質不良集計場所 - RPA用\\Obic7データ\\売上仕入実績表_(2024_8月).xlsx", 1, "MicrosoftExcel", 1, 1, "", "", 0, 0, var_ret=0, skip_err=0, delay_before=0, delay_after=0) 
# #@_コードを挿入する "df作成" 
# import pandas as pd
# 
# # Read the data from {i}月 sheet
# 売上df = pd.read_excel(売上仕入実績表, sheet_name="Sheet1", header=0, dtype=str).fillna('')
# 売上df.index = range(2, len(売上df) + 2)
# 
# # Convert the DataFrame values to a list of lists
# 売上 = 売上df.values.tolist() 
# #@ 
# Excel.Cs_WriteContent(self, 返品表, "売上仕入実績表", {"column": "A", "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": 2, "startColumn": "A", "startRow": 1, "type": 0}, 売上, 0, skip_err=0, delay_before=0, delay_after=0) 
# Excel.Cs_Close(self, 売上仕入実績表, 1, 0, skip_err=0, delay_before=0, delay_after=0) 
# #@ 
#@_コードを挿入する "OEMコード削除" 
import pandas as pd

# Read the data from {i}月 sheet
oem_df = pd.read_excel(返品表, sheet_name="OEM", header=None, dtype=str).fillna('')

# Rename the column
oem_df = oem_df.rename(columns={3: '商品ｺｰﾄﾞ'})

# Read the data from {i}月　sheet
大阪_df = pd.read_excel(r"C:\Python Learning\Work\品質不良集計場所\品質不良集計場所 - RPA用\物流Gﾃﾞｰﾀ\【大阪】返品2024年8月.xlsx", sheet_name=month[月], header=2, dtype=str).fillna('')
大阪_df.index = range(4, len(大阪_df) + 4)

# Find rows where '商品ｺｰﾄﾞ' match in both DataFrames and drop them from 大阪_df
matching_codes = oem_df['商品ｺｰﾄﾞ']
大阪_df = 大阪_df[~大阪_df['商品ｺｰﾄﾞ'].isin(matching_codes)]

# Convert the DataFrame values to a list of lists
大阪_data = 大阪_df.values.tolist()

# Read the data from {i}月　sheet
福岡_df = pd.read_excel(r"C:\Python Learning\Work\品質不良集計場所\品質不良集計場所 - RPA用\物流Gﾃﾞｰﾀ\【福岡】返品2024年8月 .xlsx", sheet_name=month[月], header=2, dtype=str).fillna('')
福岡_df.index = range(4, len(福岡_df) + 4)

# Find rows where '商品ｺｰﾄﾞ' match in both DataFrames and drop them from 大阪_df
matching_codes = oem_df['商品ｺｰﾄﾞ']
福岡_df = 福岡_df[~福岡_df['商品ｺｰﾄﾞ'].isin(matching_codes)]

# Convert the DataFrame values to a list of lists
福岡_data = 福岡_df.values.tolist() 
#@ 
for i in range(2): 
    #@_コードを挿入する "Process data" 
    df_to_process = df_names[i]
    p_table_sheet = p_table_sheet_names[i]
    品質課題_sheet = 品質課題_sheet_names[i]
    
    # Ensure "数量" column is numeric BEFORE summing
    df_to_process["数量"] = pd.to_numeric(df_to_process["数量"], errors="coerce")
    
    # Group by "商品ｺｰﾄﾞ" and take the last entry in each group
    data_df = df_to_process.groupby("商品ｺｰﾄﾞ", as_index=False).last()
    
    # Calculate the cumulative sum for the last row of each group
    data_df["数量"] = df_to_process.groupby("商品ｺｰﾄﾞ")["数量"].sum().values
    
    # Prepare content to write
    content_to_write = data_df[
        ["倉庫略称", "伝票日付", "商品ｺｰﾄﾞ", "数量", "検品前ｸﾚｰﾑ番号", "検品後ｸﾚｰﾑ番号", "ｸﾚｰﾑ内容", "返品先", "処理"]
    ].values.tolist()
    
    
    # Calculate the total sum of the "数量" column
    total_数量 = data_df["数量"].sum()
    
    p_table_write = data_df[["商品ｺｰﾄﾞ", "数量"]].values.tolist()
    
    # Sort the DataFrame by "数量" in descending order
    data_df = data_df.sort_values(by="数量", ascending=False)
    
    # Read the data from 売上仕入実績表 sheet
    売上仕入_df = pd.read_excel(返品表, sheet_name="売上仕入実績表", header=0, dtype={"当月売上数量":int})
    売上仕入_df.index = range(2, len(売上仕入_df) + 2)
    
    # Remove alphabets until digits are found in the '商品ｺｰﾄﾞ' column
    売上仕入_df['商品コード'] = 売上仕入_df['商品コード'].str.replace(r'(?<=\d)[A-Za-z]+$', '', regex=True)
    
    # Calculate the cumulative sum for the "当月売上数量" within each group
    売上仕入_df["当月売上数量"] = 売上仕入_df.groupby("商品コード")["当月売上数量"].transform("sum")
    
    # Remove duplicates based on the '商品ｺｰﾄﾞ' column
    売上仕入_df = 売上仕入_df.drop_duplicates(subset=["商品コード"])
    
    # Merging the DataFrames on the 商品ｺｰﾄﾞ and 商品コード columns
    data_df = pd.merge(data_df, 売上仕入_df, how='inner', left_on='商品ｺｰﾄﾞ', right_on='商品コード')
    
    data_df = data_df.rename(columns={
        '数量': '不良数', 
        '当月売上数量': '販売数'
    })
    
    # 不良率計算 (with condition for 0販売数)
    data_df["不良率"] = data_df.apply(lambda row: 0 if row["販売数"] == 0 else (row["不良数"] / row["販売数"]) * 100, axis=1)
    data_df["不良率"] = data_df["不良率"].round(2)
    data_df['不良率'] = data_df['不良率'].astype(str) + "%"
    
    # Read the data from the "単価" sheet
    単価表_df = pd.read_excel(返品表, sheet_name="単価", header=0, dtype={"当月在庫評価単価":int}).fillna('')
    単価表_df.index = range(2, len(単価表_df) + 2) 
    
    # Get the row with the max "当月在庫評価単価" for each 商品ｺｰﾄﾞ
    max_valuation_df = 単価表_df.loc[単価表_df.groupby("商品コード")["当月在庫評価単価"].idxmax()]
    
    # Use 'left' join to keep all rows from p_table_df
    data_df = pd.merge(data_df, max_valuation_df, how='left', left_on='商品ｺｰﾄﾞ', right_on='商品コード')
    
    data_df['ロスコスト'] = data_df['当月在庫評価単価'] * data_df['不良数']
    
    # Read the data from 製造工場 sheet
    工場_df = pd.read_excel(返品表, sheet_name="製造工場", header=0, dtype=str).fillna('')
    工場_df.index = range(2, len(工場_df) + 2)
    工場_df.rename(columns={'Unnamed: 0': '商品ｺｰﾄﾞ'}, inplace=True)
    
    # Function to get the first matching 製造工場 based on substring match
    def get_first_manufacturer(code):
        matching_rows = 工場_df[工場_df['商品ｺｰﾄﾞ'].str.contains(code, na=False)]
        if not matching_rows.empty:
            return matching_rows.iloc[0]['工場名']
        else:
            return '不明'
    
    # Apply the function to get 製造工場 for each 商品ｺｰﾄﾞ in p_table_df
    data_df['製造工場'] = data_df['商品ｺｰﾄﾞ'].apply(lambda x: get_first_manufacturer(str(x)))
    
    # Fill missing values in "ロスコスト" with 0
    data_df["ロスコスト"] = data_df["ロスコスト"].fillna(0)
    
    # Convert the DataFrame to a list of lists for writing
    data_to_write = data_df[['商品ｺｰﾄﾞ', '不良数', '販売数', '不良率', 'ロスコスト', '製造工場']].values.tolist()
    
    円_data = data_df[['商品ｺｰﾄﾞ', '不良数']].values.tolist() 
    #@ 
    Excel.Cs_WriteContent(self, 返品表, 品質課題_sheet, {"column": "A", "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": 2, "startColumn": "A", "startRow": 1, "type": 0}, content_to_write, 1, skip_err=0, delay_before=0, delay_after=0) 
    Excel.Cs_WriteContent(self, 返品表, p_table_sheet, {"column": "A", "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": 4, "startColumn": "A", "startRow": 1, "type": 0}, p_table_write, 1, skip_err=0, delay_before=0, delay_after=0) 
    pA_row = Excel.Cs_GetLastEmptyCell(self, 返品表, p_table_sheet, "指定一列", "A", 1, var_ret=0, skip_err=0, delay_before=0, delay_after=0) 
    Excel.Cs_WriteRowV2(self, 返品表, p_table_sheet, pA_row , "A", ["総計",str(total_数量)], 1, skip_err=0, delay_before=0, delay_after=0) 
    Excel.Cs_WriteCellV2(self, 返品表, p_table_sheet, "C", 2, str(total_数量), 1, skip_err=0, delay_before=0, delay_after=0) 
    Excel.Cs_WriteContent(self, 返品表, p_table_sheet, {"column": "D", "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": 2, "startColumn": "A", "startRow": 1, "type": 0}, data_to_write, 1, skip_err=0, delay_before=0, delay_after=0) 
    d_row = Excel.Cs_GetLastEmptyCell(self, 返品表, p_table_sheet, "指定一列", "D", 1, var_ret=0, skip_err=0, delay_before=0, delay_after=0) 
    #@_折りたたみ"Sorting 不良率" 
    DH_data = Excel.Cs_ReadContentV2(self, 返品表, p_table_sheet, {"column": "A", "lastColumn": "H", "lastRow": d_row - 1, "option": 3, "range": "A1:B1", "row": 1, "startColumn": "D", "startRow": 2, "type": 0}, 0, 0, var_ret=0, skip_err=0, delay_before=0, delay_after=0) 
    #@_コードを挿入する "Sorting by 不良率" 
    # Sort the data by the 4th column (multiplied by 100)
    sorted_DH_data = sorted(DH_data, key=lambda x: x[3] * 100, reverse=True)
    
    # Format the percentages and update the original sorted list
    for i in range(len(sorted_DH_data)):
        sorted_DH_data[i][3] = f"{sorted_DH_data[i][3] * 100:.2f}%"
     
    #@ 
    Excel.Cs_WriteContent(self, 返品表, p_table_sheet, {"column": "K", "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": 2, "startColumn": "A", "startRow": 1, "type": 0}, sorted_DH_data, 0, skip_err=0, delay_before=0, delay_after=0) 
    #@ 
    if (p_table_sheet == "p_table1_rpa"): 
        #@_コードを挿入する "p_table2" 
        # Calculate the sum of "不良数" for each "製造工場"
        data_df['不良数合計'] = data_df.groupby('製造工場')['不良数'].transform('sum')
        data_df['ロスコスト合計'] = data_df.groupby('製造工場')['ロスコスト'].transform('sum')
        
        # Drop duplicates based on "製造工場" and keep the first occurrence
        data_df_unique = data_df[['製造工場', '不良数合計', 'ロスコスト合計']].drop_duplicates()
        
        # Sort by 不良数合計 and ロスコスト合計 in descending order
        data_df_unique = data_df_unique.sort_values(by=['不良数合計'], ascending=False)
        data_unique = data_df_unique.sort_values(by=['ロスコスト合計'], ascending=False)
        
        # Convert the grouped data to a list of lists
        p2_不良数 = data_df_unique[['製造工場', '不良数合計']].values.tolist()
        p2_ロスコスト = data_unique[['製造工場', 'ロスコスト合計']].values.tolist()
        
        # Calculate the total sum for 不良数合計 and ロスコスト
        total_不良数 = data_df_unique['不良数合計'].sum()
        total_ロスコスト = data_unique['ロスコスト合計'].sum() 
        #@ 
        Excel.Cs_WriteContent(self, 返品表, "p_table2", {"column": "A", "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": 2, "startColumn": "A", "startRow": 1, "type": 0}, p2_不良数, 1, skip_err=0, delay_before=0, delay_after=0) 
        A_end_row = Excel.Cs_GetLastEmptyCell(self, 返品表, "p_table2", "指定一列", "A", 1, var_ret=0, skip_err=0, delay_before=0, delay_after=0) 
        Excel.Cs_WriteRowV2(self, 返品表, "p_table2", A_end_row + 1, "A", ["総計",str(total_不良数)], 1, skip_err=0, delay_before=0, delay_after=0) 
        Excel.Cs_WriteRowV2(self, 返品表, "p_table2", 25, "A", ["行ラベル","合計 / ロスコスト"], 1, skip_err=0, delay_before=0, delay_after=0) 
        Excel.Cs_WriteContent(self, 返品表, "p_table2", {"column": "A", "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": 26, "startColumn": "A", "startRow": 1, "type": 0}, p2_ロスコスト, 1, skip_err=0, delay_before=0, delay_after=0) 
        A1_end_row = Excel.Cs_GetLastEmptyCell(self, 返品表, "p_table2", "指定一列", "A", 1, var_ret=0, skip_err=0, delay_before=0, delay_after=0) 
        Excel.Cs_WriteRowV2(self, 返品表, "p_table2", A1_end_row + 1, "A", ["総計",str(total_ロスコスト)], 1, skip_err=0, delay_before=0, delay_after=0) 
    if (p_table_sheet == "p_table_rpa"): 
        Excel.Cs_WriteContent(self, 返品表, "円グラフ_rpa", {"column": "A", "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": 2, "startColumn": "A", "startRow": 1, "type": 0}, 円_data, 0, skip_err=0, delay_before=0, delay_after=0) 
#@_コードを挿入する "売上実績照会" 
import pandas as pd

df = pd.read_excel(
    r"C:\Python Learning\Work\品質不良集計場所\品質不良集計場所 - RPA用\返品8月(2024)　-　新マクロ.xlsm", 
    sheet_name="円グラフ (工場別)_rpa", 
    header=None, 
    skiprows=4, 
    nrows=33, 
    usecols=[9, 10, 11],  
    names=["主要仕入先名", "純売上数", "純売上金額"]  
)

# Read the data from 売上実績照会表
照会_df = pd.read_excel(
    r"C:\Python Learning\Work\品質不良集計場所\品質不良集計場所 - RPA用\Obic7データ\売上実績照会_2024_8_主要仕入れ先.xlsx", 
    sheet_name="Sheet1", 
    header=None, 
    skiprows=2, 
    usecols=[1, 10, 11],  
    names=["主要仕入先名", "総売上数", "総売上金額"]  
)

# Merge data based on 主要仕入先名
merged_df = df.merge(照会_df, on="主要仕入先名", how="left")

# Update 純売上数 and 純売上金額 with values from 照会_df
df["純売上数"] = merged_df["総売上数"]
df["純売上金額"] = merged_df["総売上金額"]

照会 = df[["純売上数", "純売上金額"]].values.tolist() 
#@ 
Excel.Cs_WriteContent(self, 返品表, "円グラフ (工場別)_rpa", {"column": "K", "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": 5, "startColumn": "A", "startRow": 1, "type": 0}, 照会, 0, skip_err=0, delay_before=0, delay_after=0) 
#@_折りたたみ"OEM" 
品質表 = Excel.Cs_OpenV3(self, "C:\\Python Learning\\Work\\品質不良集計場所\\品質不良集計場所 - RPA用\\品質ロボ見積用資料\\品質ロボ見積用資料\\品質報告\\品質不良率＆ロスコスト(2024）_8.xlsx", 1, "MicrosoftExcel", 1, 1, "", "", 0, 0, var_ret=0, skip_err=0, delay_before=0, delay_after=1000) 
#@_コードを挿入する "売上仕入実績表" 
# Read the data from {i}月 sheet
売上_df = pd.read_excel(返品表, sheet_name="売上仕入実績表", header=0, dtype={"当月売上数量":int,"当月売上金額":int}).fillna('')

# Calculate the totals
total_qty = 売上_df["当月売上数量"].sum()
total_amt = 売上_df["当月売上金額"].sum()
 
#@ 
#@_コードを挿入する "OEM" 
import pandas as pd

# Define column names
column_names = ["B", "C", "D", "E", "F"]  

OEM_df = pd.read_excel(返品表, sheet_name="OEM", header=None, names=column_names, dtype={'E': int, 'F': int})

# Fill missing values in column B
OEM_df["B"] = OEM_df["B"].fillna(method='ffill')

# Define groups to merge
merge_groups = {
    "エディオンOEM・ヤマダOEM": ["エディオンOEM", "ヤマダOEM"],
    "その他　OEM・カタログハウスOEM": ["その他　OEM", "カタログハウスOEM"]
}

# Merge values for defined groups
merged_rows = []
for new_name, old_names in merge_groups.items():
    merged_E = OEM_df.loc[OEM_df['B'].isin(old_names), 'E'].sum()
    merged_F = OEM_df.loc[OEM_df['B'].isin(old_names), 'F'].sum()
    merged_rows.append({'B': new_name, 'E': merged_E, 'F': merged_F})

# Remove merged rows from original DataFrame
OEM_df = OEM_df[~OEM_df['B'].isin(sum(merge_groups.values(), []))]

# Append merged rows to the DataFrame
OEM_df = pd.concat([OEM_df, pd.DataFrame(merged_rows)], ignore_index=True)

# Group by column B and sum values in E and F
OEM_df = OEM_df.groupby("B", as_index=False)[["E", "F"]].sum()

# Define the custom order
custom_order = [
    'フジ医療OEM', 
    'エディオンOEM・ヤマダOEM', 
    'FFLなど', 
    'フランスベッド\nOEM', 
    'ニトリOEM', 
    'ピップOEM', 
    'その他　OEM・カタログハウスOEM'
]

# Sort based on custom order
OEM_df['B'] = pd.Categorical(OEM_df['B'], categories=custom_order, ordered=True)
OEM_df = OEM_df.sort_values('B').reset_index(drop=True)

OEM_df.round({"F":-3})

OEM_df["F"] = OEM_df["F"].map(lambda x : x /1000)
OEM_df["F"] = OEM_df["F"].map(int)
OEM_df["F"] = OEM_df["F"].map("{:,d}".format)

# Convert DataFrame to list
E_list = OEM_df[["E"]].values.tolist()
F_list = OEM_df[["F"]].values.tolist() 
#@ 
総不良数 = Excel.Cs_ReadCell(self, 返品表, "p_table", "C", 2, 0, var_ret=0, skip_err=0, delay_before=0, delay_after=0) 
Excel.Cs_WriteCellV2(self, 品質表, search_year, col_month[月], 3, 総不良数, 0, skip_err=0, delay_before=0, delay_after=0) 
Excel.Cs_WriteCellV2(self, 品質表, search_year, col_month[月], 4, str(total_qty), 0, skip_err=0, delay_before=0, delay_after=0) 
金額 = Excel.Cs_ReadCell(self, 返品表, "p_table", "J", 1, 0, var_ret=0, skip_err=0, delay_before=0, delay_after=0) 
Excel.Cs_WriteCellV2(self, 品質表, search_year, col_month[月], 16, 金額, 0, skip_err=0, delay_before=0, delay_after=0) 
Excel.Cs_WriteCellV2(self, 品質表, search_year, col_month[月], 19, str(total_amt), 0, skip_err=0, delay_before=0, delay_after=0) 
Excel.Cs_WriteContent(self, 品質表, search_year, {"column": col_month[月], "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": 5, "startColumn": "A", "startRow": 1, "type": 0}, E_list, 0, skip_err=0, delay_before=0, delay_after=0) 
Excel.Cs_WriteContent(self, 品質表, search_year, {"column": col_month[月], "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": 20, "startColumn": "A", "startRow": 1, "type": 0}, F_list, 0, skip_err=0, delay_before=0, delay_after=0) 
#@ 
