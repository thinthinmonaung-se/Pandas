#@_コードを挿入する "必要ライブラリimport" 
import datetime
import os
import pandas as pd
import copy 
#@ 
filepath = File.GetFileList(self, configs["マクロ_excel_path"], ".xls", "$", 0, 1, var_ret=0, skip_err=0, delay_before=0, delay_after=0) 
返品表 = Excel.Cs_OpenV3(self, filepath[0], 1, "MicrosoftExcel", 1, 1, "", "", 0, 0, var_ret=0, skip_err=0, delay_before=0, delay_after=1000) 
filename = Basic.SetVariable(self, os.path.basename(返品表), var_ret=0) 
self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("作業ファイル：　", filename) 
#@_コードを挿入する "年月変換" 
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
    1: "K",  2: "L",  3: "M",  4: "B",  5: "C", 
    6: "D",  7: "E",  8: "F",  9: "G",  10: "H", 
    11: "I", 12: "J"
} 
#@ 
self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("実行データ年月：　", str(search_year) +"年　" + str(month[月]) ) 
#@_折りたたみ"フィルタリング前の作業" 
#@_コードを挿入する "df作成" 
import pandas as pd

self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("【開始】", "仕入実績表転記") 

# Read the data from {i}月 sheet
売上df = pd.read_excel(configs["売上仕入_excel_path"], sheet_name=configs["売上仕入_シート"], header=0, dtype=str)
売上df.index = range(2, len(売上df) + 2)

# Convert the DataFrame values to a list of lists
売上 = 売上df.values.tolist() 
#@ 
Excel.Cs_WriteContent(self, 返品表, configs["売上仕入実績表_シート"], {"column": configs["col_売上仕入"], "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": configs["row_売上仕入"], "startColumn": "A", "startRow": 1, "type": 0}, 売上, 1, skip_err=0, delay_before=0, delay_after=0) 
self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("【終了】", "仕入実績表転記") 
#@ 
#@_折りたたみ"返品データのフィルタリング" 
self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("【開始】", "返品データのフィルタリング ") 
#@_コードを挿入する "クレーム番号の整理" 
df = pd.read_excel(返品表, sheet_name=month[月], header=2,dtype=str).fillna('')
df.index = range(4, len(df) + 4)

#condition = (df["処理"] == "サンプル") & (df["検品後ｸﾚｰﾑ番号"] != "201")
#df.loc[condition, "検品後ｸﾚｰﾑ番号"] = "正常"
#ｸﾚｰﾑ番号 = df["検品後ｸﾚｰﾑ番号"].values.tolist() 
#@ 
# Excel.Cs_WriteContent(self, 返品表, month[月], {"column": configs["col_検品後ｸﾚｰﾑ番号"], "lastColumn": "B", "lastRow": 1, "option": 2, "range": "A1:B1", "row": configs["row_検品後ｸﾚｰﾑ番号"], "startColumn": "A", "startRow": 1, "type": 0}, ｸﾚｰﾑ番号, 1, skip_err=0, delay_before=0, delay_after=0) 
self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("①　", "正常に変更") 
#@_コードを挿入する "OEM商品削除" 
column_names = ["B", "C", "D", "E", "F"]  

# Read the data from {i}月 sheet
OEM_df = pd.read_excel(返品表, sheet_name=configs["OEM_シート"], header=None, names=column_names,skiprows=1, dtype={'E': int, 'F': int})

# Rename the column
OEM_df = OEM_df.rename(columns={"D": '商品ｺｰﾄﾞ'})

# Find rows where '商品ｺｰﾄﾞ' match in both DataFrames and drop them from excel
matching_codes = OEM_df['商品ｺｰﾄﾞ']
df = df[~df['商品ｺｰﾄﾞ'].isin(matching_codes)]
 
#@ 
self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("②　", "OEM商品削除") 
#@_コードを挿入する "倉庫移動の行削除" 
#To match the sample data, I commented out the condition from below conditions
df = df[~df['返品先'].str.contains("⇒", na=False)]
 
#@ 
self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("③　", "倉庫移動の行削除") 
#@_コードを挿入する "箱不良＆箱なしの分類" 
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
#@ 
self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("④", "箱不良＆箱なしの分類") 
#@_コードを挿入する "型番色除去" 
# Remove alphabets until digits are found in the '商品ｺｰﾄﾞ' column
df_with_706['商品ｺｰﾄﾞ'] = df_with_706['商品ｺｰﾄﾞ'].str.replace(r'(?<=\d)[A-Za-z]+$', '', regex=True)
df_without_706['商品ｺｰﾄﾞ'] = df_without_706['商品ｺｰﾄﾞ'].str.replace(r'(?<=\d)[A-Za-z]+$', '', regex=True) 
#@ 
self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("⑤", "型番色除去") 
self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("【終了】", "返品データのフィルタリング") 
#@ 
#@_折りたたみ"返品データ集計" 
self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("【開始】", "返品データ集計") 
#@_コードを挿入する "プロセス" 
def process_returns_data(df_to_process, 返品表, configs, df_type=""):
    total_数量 = 0
    全体販売台数_OEM = 0
    不良金額_OEM = 0
    全体販売金額_OEM = 0
    p2_不良数 = []
    p2_ロスコスト = []
    p2_不良数合計 = 0
    p2_ロスコスト合計 = 0
    円_data = []
    
    content_to_write = df_to_process[["倉庫略称", "伝票日付", "商品ｺｰﾄﾞ", "数量", "検品前ｸﾚｰﾑ番号", "検品後ｸﾚｰﾑ番号", "ｸﾚｰﾑ内容", "返品先", "処理"]].values.tolist()
    
    df_to_process["数量"] = pd.to_numeric(df_to_process["数量"], errors="coerce")
    
    data_df = df_to_process.groupby("商品ｺｰﾄﾞ", as_index=False).agg({'数量': 'sum'})
    total_数量 = data_df["数量"].sum()
    
    p_table_write = data_df[["商品ｺｰﾄﾞ", "数量"]].values.tolist()
    self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("①", "不良数集計") 

    売上仕入_df = pd.read_excel(返品表, sheet_name=configs["売上仕入実績表_シート"], header=0, dtype={"当月売上数量": int, "当月売上金額": int})
    売上仕入_df.index = range(2, len(売上仕入_df) + 2)
    売上仕入_df['商品コード'] = 売上仕入_df['商品コード'].str.replace(r'(?<=\d)[A-Za-z]+$', '', regex=True)

    全体販売台数_OEM = 売上仕入_df["当月売上数量"].sum()
    全体販売金額_OEM = 売上仕入_df["当月売上金額"].sum()

    売上仕入_df = 売上仕入_df.groupby("商品コード", as_index=False)['当月売上数量'].sum()

    data_df = pd.merge(data_df, 売上仕入_df, how='left', left_on='商品ｺｰﾄﾞ', right_on='商品コード')
    data_df = data_df.rename(columns={'数量': '不良数', '当月売上数量': '販売数'})
    
    data_df = data_df.sort_values(by="不良数", ascending=False)
    self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("②", "不良数集計データ並び替え")

    data_df["不良率"] = data_df.apply(
        lambda row: 0 if row["販売数"] == 0 else round((row["不良数"] / row["販売数"]) * 100, 2),
        axis=1
    )
    data_df["不良率"] = data_df["不良率"].astype(float).round(2)
    data_df["不良率"] = data_df["不良率"].map(lambda x: f"{x:.2f}%")
    self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("③", "不良率算出え")

    単価表_df = pd.read_excel(返品表, sheet_name=configs["単価_シート"], header=0, dtype={"当月在庫評価単価": int})
    単価表_df.index = range(2, len(単価表_df) + 2)

    max_valuation_df = 単価表_df.loc[単価表_df.groupby("商品コード")["当月在庫評価単価"].idxmax()]
    data_df = pd.merge(data_df, max_valuation_df, how='left', left_on='商品ｺｰﾄﾞ', right_on='商品コード')
    data_df['ロスコスト'] = data_df['当月在庫評価単価'] * data_df['不良数']
    self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("④", "ロストコスト算出")

    工場_df = pd.read_excel(返品表, sheet_name=configs["製造工場_シート"], header=0, dtype=str)
    工場_df.index = range(2, len(工場_df) + 2)
    工場_df.rename(columns={'Unnamed: 0': '商品ｺｰﾄﾞ'}, inplace=True)
    self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("⑤", "製造工場反映")

    def get_first_manufacturer(code):
        matching_rows = 工場_df[工場_df['商品ｺｰﾄﾞ'].str.contains(code, na=False)]
        return matching_rows.iloc[0]['工場名'] if not matching_rows.empty else '不明'

    data_df['製造工場'] = data_df['商品ｺｰﾄﾞ'].apply(lambda x: get_first_manufacturer(str(x)))
    data_df["ロスコスト"] = data_df["ロスコスト"].fillna(0)

    if df_type == "df_with_706":
        不良金額_OEM = data_df["ロスコスト"].sum()
        円_data = data_df[['商品ｺｰﾄﾞ', '不良数']].values.tolist()
        
    if df_type == "df_without_706":
        p2_不良数合計_df = data_df.groupby('製造工場')['不良数'].sum().reset_index(name='不良数合計').sort_values('不良数合計', ascending=False)
        p2_不良数 = p2_不良数合計_df.values.tolist()
        p2_不良数合計 = p2_不良数合計_df['不良数合計'].sum()
    
        p2_ロスコス_df = data_df.groupby('製造工場')['ロスコスト'].sum().reset_index(name='ロスコスト合計').sort_values('ロスコスト合計', ascending=False)
        p2_ロスコスト = p2_ロスコス_df.values.tolist()
        p2_ロスコスト合計 = p2_ロスコス_df['ロスコスト合計'].sum()

        
    data_to_write = data_df[['商品ｺｰﾄﾞ', '不良数', '販売数', '不良率', 'ロスコスト', '製造工場']].values.tolist()
    sorted_data_write = sorted(copy.deepcopy(data_to_write), key=lambda x: float(x[3].strip('%')), reverse=True)
    self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("⑥", "不良率で並び替え")
    self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("⑦", "総不良数算出")
    self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("⑧", "総ロストコスト算出")
    
    if df_type == "df_without_706":
        self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("【終了】 ", "返品データ集計")

    return content_to_write, p_table_write, data_to_write, sorted_data_write, total_数量, 全体販売台数_OEM, 不良金額_OEM, 全体販売金額_OEM, 円_data, p2_不良数, p2_ロスコスト, p2_不良数合計, p2_ロスコスト合計

content_to_write, p_table_write, data_to_write, sorted_data_write,total_数量, 全体販売台数_OEM, 不良金額_OEM, 全体販売金額_OEM,円_data,p2_不良数, p2_ロスコスト, p2_不良数合計, p2_ロスコスト合計 =process_returns_data(df_with_706, 返品表, configs, df_type="df_with_706")
content_to_write1, p_table_write1, data_to_write1, sorted_data_write1,total_数量1, 全体販売台数_OEM1, 不良金額_OEM1, 全体販売金額_OEM1,円_data1,p2_不良数1, p2_ロスコスト1, p2_不良数合計1, p2_ロスコスト合計1 = process_returns_data(df_without_706, 返品表, configs, df_type="df_without_706")

data_array = [
    (content_to_write, p_table_write, data_to_write, sorted_data_write,total_数量, 全体販売台数_OEM, 不良金額_OEM, 全体販売金額_OEM,円_data,p2_不良数, p2_ロスコスト, p2_不良数合計, p2_ロスコスト合計),
    (content_to_write1, p_table_write1, data_to_write1, sorted_data_write1,total_数量1, 全体販売台数_OEM1, 不良金額_OEM1, 全体販売金額_OEM1,円_data1,p2_不良数1, p2_ロスコスト1, p2_不良数合計1, p2_ロスコスト合計1)
]

品質課題_sheet_names = [configs["品質課題_シート"], configs["品質課題_1_シート"]]
p_table_sheet_names = [configs["p_table_シート"], configs["p_table1_シート"]]

# OEMデータ作業
OEM_df["B"] = OEM_df["B"].fillna(method='ffill')

merge_groups = {
    "エディオンOEM・ヤマダOEM": ["エディオンOEM", "ヤマダOEM"],
    "その他　OEM・カタログハウスOEM": ["その他　OEM", "カタログハウスOEM"]
}

merged_rows = []
for new_name, old_names in merge_groups.items():
    merged_E = OEM_df.loc[OEM_df['B'].isin(old_names), 'E'].sum(min_count=1)
    merged_F = OEM_df.loc[OEM_df['B'].isin(old_names), 'F'].sum(min_count=1)
    merged_rows.append({'B': new_name, 'E': merged_E, 'F': merged_F})

OEM_df = OEM_df[~OEM_df['B'].isin(sum(merge_groups.values(), []))]
OEM_df = pd.concat([OEM_df, pd.DataFrame(merged_rows)], ignore_index=True)
OEM_df = OEM_df.groupby("B", as_index=False)[["E", "F"]].sum()

rename_mapping = {
    "フジ医療OEM": "フジ医療器台数（フジ医療で検索）",
    "エディオンOEM・ヤマダOEM": "家電量販店OEM台数",
    "FFLなど": "アメリカ台数（アメリカで検索）",
    "フランスベッド\nOEM": "フランスベッド台数（フランスで検索）",
    "ニトリOEM": "ニトリ台数",
    "ピップOEM": "ピップ台数（ピップ　ルポゼ）",
    "その他　OEM・カタログハウスOEM": "その他"
}

OEM_df['B'] = OEM_df['B'].replace(rename_mapping)
OEM_df['B'] = pd.Categorical(OEM_df['B'], categories=rename_mapping.values(), ordered=True)
OEM_df = OEM_df.sort_values('B').reset_index(drop=True)
OEM_df["F"] = OEM_df["F"].round(-3)
E_list = OEM_df[["E"]].values.tolist()
F_list = OEM_df[["F"]].values.tolist()

#円グラフ (工場別)反映
df = pd.read_excel(
    返品表, 
    sheet_name=configs["円グラフ (工場別)_シート"], 
    header=None, 
    skiprows=4, 
    nrows=33, 
    usecols=[9, 10, 11],  
    names=["主要仕入先名", "純売上数", "純売上金額"]  
)

照会_df = pd.read_excel(
    configs["照会_excel_path"], 
    sheet_name=configs["照会_シート"], 
    header=None, 
    skiprows=2, 
    usecols=[1, 10, 11],  
    names=["主要仕入先名", "総売上数", "総売上金額"]  
)

merged_df = df.merge(照会_df, on="主要仕入先名", how="left")
df["純売上数"] = merged_df["総売上数"]
df["純売上金額"] = merged_df["総売上金額"]
照会 = df[["純売上数", "純売上金額"]].values.tolist() 
#@ 
#@ 
#@_折りたたみ"データの書き込み" 
for i in range(len(data_array)): 
    #@_コードを挿入する "\"i\" 変数設定" 
    content_to_write, p_table_write, data_to_write, sorted_data_write,total_数量, 全体販売台数_OEM, 不良金額_OEM, 全体販売金額_OEM,円_data,p2_不良数, p2_ロスコスト, p2_不良数合計, p2_ロスコスト合計 = data_array[i]
    
    for item in sorted_data_write:
        del item[-1]
    
    p_table_sheet = p_table_sheet_names[i]
    品質課題_sheet = 品質課題_sheet_names[i] 
    #@ 
    Excel.Cs_WriteContent(self, 返品表, 品質課題_sheet, {"column": configs["col_品質課題"], "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": configs["row_品質課題"], "startColumn": "A", "startRow": 1, "type": 0}, content_to_write, 1, skip_err=0, delay_before=0, delay_after=0) 
    Excel.Cs_WriteContent(self, 返品表, p_table_sheet, {"column": configs["col_p_table"], "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": configs["row_p_table"], "startColumn": "A", "startRow": 1, "type": 0}, p_table_write, 1, skip_err=0, delay_before=0, delay_after=0) 
    pA_row = Excel.Cs_GetLastEmptyCell(self, 返品表, p_table_sheet, "指定一列", configs["col_p_table"], 1, var_ret=0, skip_err=0, delay_before=0, delay_after=0) 
    Excel.Cs_WriteRowV2(self, 返品表, p_table_sheet, pA_row , configs["col_p_table"], ["総計",str(total_数量)], 1, skip_err=0, delay_before=0, delay_after=0) 
    Excel.Cs_WriteCellV2(self, 返品表, p_table_sheet, configs["col_total_数量"], configs["row_数量"], total_数量, 1, skip_err=0, delay_before=0, delay_after=0) 
    Excel.Cs_WriteContent(self, 返品表, p_table_sheet, {"column": configs["col_orderby_不良数"], "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": configs["row_orderby_不良数"], "startColumn": "A", "startRow": 1, "type": 0}, data_to_write, 1, skip_err=0, delay_before=0, delay_after=0) 
    Excel.Cs_WriteContent(self, 返品表, p_table_sheet, {"column": configs["col_orderby_不良率"], "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": configs["row_orderby_不良率"], "startColumn": "A", "startRow": 1, "type": 0}, sorted_data_write, 0, skip_err=0, delay_before=0, delay_after=0) 
    if (i == 0): 
        #@_折りたたみ"p_tableデータ使用" 
        品質表 = Excel.Cs_OpenV3(self, configs["品質_excel_path"], 1, "MicrosoftExcel", 1, 1, "", "", 0, 0, var_ret=0, skip_err=0, delay_before=0, delay_after=1000) 
        シート名 = Excel.Cs_GetName(self, 品質表, "全部", var_ret=0, skip_err=0, delay_before=0, delay_after=0) 
        if (str(search_year) not in シート名): 
            Excel.Cs_CopyV2(self, 品質表, configs["sample_シート"], 0, None, str(search_year), 0, 1, skip_err=0, delay_before=0, delay_after=0) 
            Excel.Cs_Move(self, 品質表, str(search_year), "按位置", 1, "", "之前", "第一位", 1, skip_err=0, delay_before=0, delay_after=0) 
        Excel.Cs_WriteCellV2(self, 品質表, str(search_year), col_month[月], configs["row_total_数量"], total_数量, 0, skip_err=0, delay_before=0, delay_after=0) 
        Excel.Cs_WriteCellV2(self, 品質表, str(search_year), col_month[月], configs["row_total_qty"], 全体販売台数_OEM, 0, skip_err=0, delay_before=0, delay_after=0) 
        Excel.Cs_WriteCellV2(self, 品質表, str(search_year), col_month[月], configs["row_total_ロスコスト"], 不良金額_OEM, 0, skip_err=0, delay_before=0, delay_after=0) 
        Excel.Cs_WriteCellV2(self, 品質表, str(search_year), col_month[月], configs["row_total_amt"], 全体販売金額_OEM, 0, skip_err=0, delay_before=0, delay_after=0) 
        Excel.Cs_WriteContent(self, 品質表, str(search_year), {"column": col_month[月], "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": configs["row_OEM_E"], "startColumn": "A", "startRow": 1, "type": 0}, E_list, 0, skip_err=0, delay_before=0, delay_after=0) 
        Excel.Cs_WriteContent(self, 品質表, str(search_year), {"column": col_month[月], "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": configs["row_OEM_F"], "startColumn": "A", "startRow": 1, "type": 0}, F_list, 0, skip_err=0, delay_before=0, delay_after=0) 
        #@ 
        Excel.Cs_WriteContent(self, 返品表, configs["円グラフ_シート"], {"column": configs["col_円グラフ"], "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": configs["row_円グラフ"], "startColumn": "A", "startRow": 1, "type": 0}, 円_data, 0, skip_err=0, delay_before=0, delay_after=0) 
        self.TASK_COMPONENT_Gn6grr11739335179514_1_0_4("作業名：", "返品数トップ10グラフ更新") 
    else: 
        #@_折りたたみ"p_table1データ使用" 
        Excel.Cs_WriteContent(self, 返品表, configs["p_table2_シート"], {"column": configs["col_p_table2"], "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": configs["row_p_table2_不良数"], "startColumn": "A", "startRow": 1, "type": 0}, p2_不良数, 1, skip_err=0, delay_before=0, delay_after=0) 
        A_end_row = Excel.Cs_GetLastEmptyCell(self, 返品表, configs["p_table2_シート"], "指定一列", configs["col_p_table2"], 1, var_ret=0, skip_err=0, delay_before=0, delay_after=0) 
        Excel.Cs_WriteRowV2(self, 返品表, configs["p_table2_シート"], A_end_row + 1, configs["col_p_table2"], ["総計",str(p2_不良数合計)], 1, skip_err=0, delay_before=0, delay_after=0) 
        Excel.Cs_WriteRowV2(self, 返品表, configs["p_table2_シート"], configs["row_title_ロスコスト"], configs["col_p_table2"], ["行ラベル","合計 / ロスコスト"], 1, skip_err=0, delay_before=0, delay_after=0) 
        Excel.Cs_WriteContent(self, 返品表, configs["p_table2_シート"], {"column": configs["col_p_table2"], "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": configs["row_p_table2_ロスコスト"], "startColumn": "A", "startRow": 1, "type": 0}, p2_ロスコスト, 1, skip_err=0, delay_before=0, delay_after=0) 
        A1_end_row = Excel.Cs_GetLastEmptyCell(self, 返品表, configs["p_table2_シート"], "指定一列", configs["col_p_table2"], 1, var_ret=0, skip_err=0, delay_before=0, delay_after=0) 
        Excel.Cs_WriteRowV2(self, 返品表, configs["p_table2_シート"], A1_end_row + 1, configs["col_p_table2"], ["総計",str(p2_ロスコスト合計)], 1, skip_err=0, delay_before=0, delay_after=0) 
        #@ 
Excel.Cs_WriteContent(self, 返品表, configs["円グラフ (工場別)_シート"], {"column": configs["col_照会"], "lastColumn": "B", "lastRow": 1, "option": 3, "range": "A1:B1", "row": configs["row_照会"], "startColumn": "A", "startRow": 1, "type": 0}, 照会, 0, skip_err=0, delay_before=0, delay_after=0) 
