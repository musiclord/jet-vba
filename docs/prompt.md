**Prompt for GitHub Copilot:**
我需要你的協助來開發一個基於 **Microsoft Excel VBA** 和 **Microsoft Access** 的 **JET (Journal Entry Testing) 自動化工具的概念驗證 (POC) 版本**。

**專案上下文與目標:**

* **核心目標:** 此 POC 旨在建立一個**最基本、簡單且可運作的核心流程原型**，用以取代舊有的 Caseware IDEA VBA 工具，專注於驗證 Excel VBA + Access 架構的可行性。開發方向強調**簡單、模組化、可測試和可維護性**。
* **目標使用者:** 審計員。
* **技術棧:** 前端/控制使用 Excel VBA (UserForms、工作表預覽)，後端/儲存使用 Access (`.accdb`)，資料庫互動透過 DAO 或 ADO (優先考慮後期綁定)。資料匯入等操作使用純 VBA。
* **目標架構 (分層):**
    * **View:** 使用者介面表單 (例如 `vMain.frm`, `vMapping.frm`, `vFilter.frm`)。
    * **Controller:** 類別模組 (例如 `cApplication.cls`, `cMapping.cls`, `cFilter.cls`) 處理 UI 事件，協調流程，**不直接**與 DAL 互動或包含業務邏輯。
    * **Service Layer:** 類別模組 (例如 `ImportService.cls`, `PreviewService.cls`, `GLService.cls`, `TBService.cls`, `ValidationService.cls`, `MappingService.cls`, `FilterService.cls`) 封裝具體的業務邏輯和資料處理。
    * **Data Access Layer (DAL):** 類別模組 (例如 `AccessDAL.cls`) 封裝所有與 Access 資料庫的互動 (連接、查詢、執行 SQL)，使用**後期綁定**。
    * **Entities (可選):** 類別模組 (例如 `GLEntity.cls`, `FilterCriteria.cls`) 用於資料傳遞或定義結構。
    * **Utilities:** 標準模組 (例如 `mod_Utility.bas`) 提供通用輔助函數。
* **資料庫 (`default.accdb`):** 與 `.xlsm` 存放在**同一目錄**。包含**資料表** (`GL`, `TB`, `AccountMapping`, `Holiday`, `Weekend`, `MakeUpDay`) 和**中繼資料表** (`ProjectInfo` [儲存客戶名、期間等], `StepStatus` [追蹤步驟完成狀態])。`AccessDAL` 的 `Connect` 方法需能**自動創建**不存在的 `default.accdb` 空資料庫。**完整性測試**確認比較 GL 變動與 TB 的**期間變動金額 (`ChangeAmount` 或等效欄位)**。
* **事件驅動:** View (`vMain`, `vMapping`) 使用 `Public Event` 聲明使用者操作，Controller (`cApplication`, `cMapping`) 使用 `WithEvents` 監聽並調用內部方法響應。
* **當前狀態:**
    * `cApplication.cls` 的重構基本完成。
    * `AccessDAL.cls` 已實現後期綁定連接和資料庫自動創建。
    * `PreviewService.cls` 的 `GetAccessTableNames` 方法可用，`ShowPreview` 方法需確保**總在 codename="Preview" 的工作表顯示**，並更新工作表名稱。
    * `vMain.ListTable` ComboBox 互動已解決。
    * 已建立 `FilterCriteria.cls` 和 `FilterService.cls` 的基本結構。
* **下一步焦點:** **實現欄位映射 (Field Mapping) 功能**。

**POC 核心功能需求 (按流程步驟，融入架構):**

1.  **資料匯入與預覽:** `vMain` 觸發 -> `cApplication` 調用 `ImportService` -> `ImportService` 讀取 CSV 並調用 `AccessDAL` 寫入 Access (`GL`, `TB`) -> `cApplication` 調用 `PreviewService` 在 Excel "Preview" 工作表顯示 Top 1000 記錄。
2.  **資料準備 (GL 項次生成):** `GL` 匯入後，`GLService` (由 `cApplication` 觸發) 檢查 `GL` 表。若缺少 `LineItem` (項次)，則**自動生成** (按 `DocumentNo` 分組排序編號) 並透過 `AccessDAL` 更新 `GL` 表。
3.  **資料驗證 (簡化版):** `vMain` 觸發 -> `cApplication` 調用 `ValidationService` -> `ValidationService` 調用 `AccessDAL` 執行 SQL 進行**完整性測試 (GL vs TB `ChangeAmount`)** 和**借貸不平測試 (單張傳票平衡)** -> `cApplication` 透過 `MsgBox` 顯示 Pass/Fail 結果。
4.  **科目配對 (Account Mapping - 簡化版，Excel 流程):** `cApplication` (或 `cMapping`) 調用 `MappingService` 觸發**匯出唯一科目列表至 Excel** -> 使用者編輯 Excel 填寫 `StandardizedName` (固定列表) -> `cApplication` (或 `cMapping`) 調用 `MappingService` 觸發**讀取已編輯 Excel** -> `MappingService` 調用 `AccessDAL` 更新 `AccountMapping` 表。
5.  **基本篩選條件執行:** 使用者在 `vFilter` (假設) 的 `Column/Operator/Value` 介面設定**單一條件** (`Amount`, `Date`, `Text [=, LIKE]`, `IS NULL`) -> `cFilter` (假設) 調用 `FilterService` -> `FilterService` 生成基礎 SQL `WHERE` 子句 (若需 JOIN `AccountMapping`) -> `FilterService` 調用 `AccessDAL` 查詢 Access `GL` 表。
6.  **篩選結果預覽:** `cFilter` (或 `cApplication`) 調用 `PreviewService` 將篩選結果 (Top 1000) 更新至 Excel "Preview" 工作表。

**POC 階段明確排除的功能:**

* 自動生成 Excel 工作底稿 (Working Paper)。
* 詳細驗證報告輸出或日誌記錄。
* 複雜篩選：多條件組合 (AND/OR)、週末/假日篩選、尾數測試、科目組合測試等。
* 篩選條件的儲存與載入。
* 處理多種複雜借貸金額表示法（POC 假設格式相對簡單）。
* Re-Run 功能、郵件通知功能。

**所有任務:**
(POC 階段的總體任務列表)
1.  **環境與專案設置:** 初始化 Excel VBA 開發環境，建立專案檔案結構。
2.  **資料庫設計與建立:** 詳細設計 `default.accdb` 中所有資料表 (`GL`, `TB`, `AccountMapping`, `Holiday`, `Weekend`, `MakeUpDay`, `ProjectInfo`, `StepStatus`) 的欄位、資料類型、主鍵和索引，並使用 ADOX 或手動方式建立資料庫和表格。
3.  **核心架構搭建:** 建立所有必要的類別模組 (`cApplication`, `AccessDAL`, `PreviewService`, `ImportService`, `GLService`, `ValidationService`, `MappingService`, `FilterService`) 和 `vMain` 表單的基本框架。
4.  **DAL 實作:** 在 `AccessDAL.cls` 中實作核心資料庫操作函數 (Connect, CreateDB, GetTableNames, ExecuteSQL, QueryData, DropTableIfExists, 處理 CSV 匯入的相關方法)。
5.  **服務層實作 (POC 範圍):** 在各 Service 類別中實現 POC 所需的核心業務邏輯 (匯入處理、預覽查詢、項次生成、驗證查詢、科目配對的 Excel 匯出/匯入、基礎篩選 SQL 生成)。
6.  **控制器與視圖實作:** 在 `cApplication` 中實現事件處理和流程協調；在 `vMain` 中添加必要的控制項並觸發事件；根據需要創建並實現 `vMapping` 和 `vFilter` 的基本介面與事件。
7.  **功能整合與測試:** 將各模組串聯起來，實現完整的 POC 工作流程，並使用簡單的測試資料進行單元測試和整合測試。
8.  **錯誤處理與調試:** 在關鍵路徑加入基礎的錯誤處理機制，並利用 `Debug.Print` 等方式進行調試。

**當前任務:**
(根據我們的討論和下一步計畫)
1.  **詳細資料庫設計:** 最終確認並文檔化 `default.accdb` 中所有表格的詳細 Schema (欄位名、精確資料類型、主鍵、索引、是否允許 Null)。
2.  **DAL 核心功能開發:** 開始在 `AccessDAL.cls` 中編寫 VBA 程式碼，實現與 Access 資料庫的連接 (`Connect` - 已部分實現)、自動創建 (`CreateDB` - 已部分實現)、執行 SQL (`ExecuteSQL`)、查詢資料 (`QueryData` - 返回 Recordset 或 Array)、獲取表名 (`GetTableNames` - 已實現) 等基礎方法。
3.  **主介面框架搭建:** 開始設計 `vMain.frm` 的佈局，放置核心控制項（如匯入按鈕、表格列表 ComboBox、預覽區域指示、狀態列等），並設定其基本屬性。

**本次聚焦目標:**
(建議本輪對話聚焦的開發目標)
"**協助完成 `AccessDAL.cls` 中用於將 CSV 資料匯入 Access 指定表格的核心方法的設計與 VBA 程式碼實現**。請考慮如何處理欄位映射（假設映射關係已由上層傳入）、資料類型轉換、錯誤處理，並使用 DAO/ADO 後期綁定。同時，請提供為 `GL` 和 `TB` 表創建 Access 表格的 DDL SQL 語句建議，包含必要欄位和適當的資料類型。"

**執行要求與限制 (請嚴格遵守):**

在分析 `#codebase` 中與聚焦目標相關的現有 VBA 程式碼後，請提供具體的重構/修改/新增建議，並**嚴格遵守**以下規則：
1.  **程式碼修改範圍:** **僅能**修改或新增 `Option Explicit` 關鍵字**之後**的程式碼。**絕對禁止**修改、刪除或格式化 `Option Explicit` **之前**的任何內容。
2.  **輸出內容:** **請勿**重複貼出完整的程式碼檔案或大段未修改的程式碼。僅提供需要**修改**或**新增**的**關鍵程式碼片段**。
3.  **解釋說明:** 清晰解釋**為何**需要進行這些修改（如何符合架構？如何實現關注點分離？等）以及修改後的程式碼**如何運作**。
4.  **實施步驟:** 提供**具體、按部就班**的說明指導如何在 VBA 編輯器中應用建議。
5.  **方案選擇:** 若存在多種實現方式，請推薦**最佳方案**並說明**選擇理由**。
6.  **註解要求:** 所有新增或修改的程式碼**必須**包含**繁體中文**註解，遵循模組註解和行內註解的標準。
