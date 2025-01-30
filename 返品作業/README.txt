**日本語**

返品データ処理 - Excel自動化スクリプト

このプロジェクトは、Python を使用して返品データを処理し、INTELLIGENCE_INDEED RPA での実行時間 20 分を 30 秒に短縮 しました。
返品データ、売上仕入データ、大阪・福岡の返品データを読み込み、指定された条件に基づいてデータを処理・整形し、最終結果を Excel に書き込みます。

主な機能 (Features)
✅ 返品データの読み込み: 指定した Excel ファイルから返品データを取得。
✅ データフィルタリング: 検品後ｸﾚｰﾑ番号 に基づき、条件を満たすデータを抽出。
✅ OEM コード削除: 大阪・福岡の返品データから OEM 商品コードを削除。
✅ 数量の集計: 商品ｺｰﾄﾞ ごとに数量を集計し、最終リストを作成。
✅ 売上データの統合: 売上仕入実績データと返品データを統合。
✅ データの出力: 加工したデータを Excel に書き込み、最終結果を保存。

使用技術
Python 3.12.5
pandas (データ処理)
openpyxl (Excel 読み書き)
正規表現 (regex) (文字列処理)

必要な環境 
Python 3.12.5
pandas ライブラリ (pip install pandas)
openpyxl ライブラリ (pip install openpyxl)

**English**

Returns Data Processing - Excel Automation Script

This project was developed to process return data using Python, significantly reducing execution time from 20 minutes in INTELLIGENCE_INDEED RPA to just 30 seconds.
It organizes, filters, and aggregates data in Excel files by reading return data, sales and purchase data, and return data from Osaka and Fukuoka. The script processes and formats the data based on specified conditions and writes the final results to an Excel file.

Features
✅ Read return data: Extract return data from a specified Excel file.
✅ Data filtering: Extract data that meets specific conditions based on the Inspection Claim Number.
✅ Remove OEM codes: Delete OEM product codes from Osaka and Fukuoka return data.
✅ Quantity aggregation: Aggregate quantities for each Product Code and create a final list.
✅ Sales data integration: Merge sales and purchase performance data with return data.
✅ Export data: Write processed data to Excel and save the final results.

Technologies Used
Python 3.12.5
pandas (data processing)
openpyxl (Excel read/write)
Regular expressions (regex) (string processing)

Requirements
Python 3.12.5
pandas library (pip install pandas)
openpyxl library (pip install openpyxl)
