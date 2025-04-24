# jet-vba

本專案旨在使用 Excel VBA 開發日記帳分錄測試 (Journal Entry Testing, JET) 工具。

## 架構

本專案遵循受 MVC 原則啟發的分層架構，以促進關注點分離和可維護性：

*   **視圖 (`vMain.frm`, `vMapping.frm`):** 用於互動的使用者介面表單。
*   **控制器 (`cApplication.cls`, `cMapping.cls`):** 處理 UI 事件，協調應用程式流程，並將任務委派給服務層。
*   **服務層 (`ImportService.cls`, `PreviewService.cls`, `GLService.cls`, `TBService.cls`, `MappingService.cls`):** 封裝特定功能的業務邏輯（匯入、預覽、資料處理、映射）。
*   **資料存取層 (DAL) (`AccessDAL.cls`):** 管理與 Microsoft Access 資料庫 (`.accdb`) 的所有互動，使用 ADODB 和 ADOX 並採用後期綁定。
*   **公用程式 (`mod_Utility.bas`):** 包含通用輔助函數（例如：檔案選擇、CSV 編碼偵測）。

## 主要功能

*   **CSV 匯入:** 將總帳 (GL) 和試算表 (TB) 資料從 CSV 檔案匯入 Access 資料庫 (`default.accdb`)。
    *   處理不同的 CSV 編碼（偵測 UTF-8 BOM，若偵測失敗則預設為 950）。
    *   匯入時自動刪除並重新建立目標資料表 (`GL`, `TB`)。
*   **資料庫管理:**
    *   首次嘗試連接時，如果 `poc` 目錄中不存在 `default.accdb` Access 資料庫檔案，則自動建立該檔案。
*   **資料預覽:** 將選定 Access 資料表的資料載入專用的 Excel 工作表（例如：`GL_Preview`, `TB_Preview`）供使用者檢視。
    *   將預覽限制為可設定的列數 (`MAX_ROWS_TO_SHOW`)。
*   **動態資料表清單:** 點擊下拉按鈕時，使用 Access 資料庫中可用的使用者資料表填充 `vMain` 上的 ComboBox (`ListTable`)。
*   **欄位映射:** 用於將來源欄位映射到目標欄位的基礎架構（實作細節在 `cMapping`, `vMapping`, `MappingService` 中）。
*   **資料處理:** 用於處理匯入資料的服務 (`GLService`, `TBService`)（細節待進一步定義/實作）。

## 目前狀態 (截至 2025年4月24日)

*   核心架構（視圖、控制器、服務、DAL）已建立。
*   CSV 匯入、自動資料庫建立和基本資料預覽功能已實作。
*   `cApplication.cls` 的重構已大致完成，以符合控制器職責。
*   **目前焦點/問題:** 解決 `vMain` 上的 `ListTable` ComboBox 在點擊下拉按鈕時無法顯示其下拉列表的問題，儘管底層項目已透過 `GetTableNames` 程序正確更新。

## 設定與使用

1.  在 Microsoft Excel 中開啟 `poc/JET.xlsm` 檔案。
2.  如果出現提示，請啟用巨集。
3.  主介面 (`vMain`) 應該會出現。
4.  使用按鈕匯入 GL/TB CSV 檔案（確保範例 CSV 位於 `poc/data` 資料夾中）。
5.  點擊資料表清單 ComboBox 上的下拉箭頭以刷新並檢視可用的資料表（目前遇到下拉顯示問題）。
6.  選擇一個資料表（如果可能）並點擊「預覽」以在 Excel 中檢視資料。
7.  如果 `poc` 資料夾中不存在 `default.accdb` 資料庫，它將被自動建立。

## 技術

*   Microsoft Excel VBA
*   Microsoft ActiveX Data Objects (ADODB) - 後期綁定
*   Microsoft ADO Ext. for DDL and Security (ADOX) - 後期綁定
*   Microsoft Scripting Runtime (FileSystemObject) - 後期綁定