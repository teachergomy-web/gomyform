兒美成績系統 - 更新說明
========================

【資料夾結構】
GomyGrade.exe       ← 主程式（只需打包一次）
pages.py            ← 可熱更新
data_manager.py     ← 可熱更新
ui_utils.py         ← 可熱更新
docx_exporter.py    ← 可熱更新
excel_importer.py   ← 可熱更新
一鍵更新.bat         ← 更新腳本
成績單模板.docx      ← 模板
升級考模板.docx      ← 模板
data/
  grades.json       ← 班級資料（不會被更新覆蓋）

【如何更新】
1. 確保電腦有網路連線
2. 關閉程式
3. 執行「一鍵更新.bat」
4. 重新開啟程式

【注意事項】
- grades.json 不會被更新覆蓋，資料安全
- 如果更新後程式有問題，可以向管理員索取舊版 .py 檔案
