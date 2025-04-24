**上下文:**
好的，這是根據我們目前對話內容，對專案重構的狀態、需求和注意事項的十點總結：
1.  **目標架構:** 專案的核心目標是將 VBA 程式碼重構成一個分層架構，包含：
    *   **View:** `vMain.frm`, `vMapping.frm` (使用者介面表單)。
    *   **Controller:** `cApplication.cls`, `cMapping.cls` (處理 UI 事件，協調流程)。
    *   **Service Layer:** `ImportService.cls`, `PreviewService.cls`, `GLService.cls`, `TBService.cls`, `MappingService.cls`, `FilterService.cls` (封裝業務邏輯)。
    *   **Data Access Layer (DAL):** `AccessDAL.cls` (封裝與 Access 資料庫的互動)。
    *   **Entities (Optional):** 如 `GLEntity.cls`, `TBEntity.cls`, `FilterCriteria.cls` (用於資料傳遞或定義結構)。
    *   **Utilities:** `mod_Utility.bas` (通用輔助函數)。

2.  **當前重構焦點:** 目前的工作主要集中在 `cApplication.cls`，確保它符合控制器的職責：僅處理來自 `vMain` 的事件，協調應用程式的主要流程（如匯入、預覽、處理），並將具體工作委派給相應的 Service 層。

3.  **`cApplication` 職責劃分:** `cApplication` 嚴格遵守關注點分離原則：
    *   **不包含**直接調用 `AccessDAL` 的程式碼。
    *   **不包含**複雜的資料處理或業務規則計算。
    *   其方法（如 `ImportCSV`, `PreviewTable`, `PopulateTableList`）主要負責：接收事件、準備參數、調用 Service 層方法、處理 Service 返回結果，以及更新 UI 狀態（如狀態列、訊息框、啟用/禁用控制項）。

4.  **事件驅動流程:** `vMain.frm` 使用 `Public Event` (如 `DoImportGL`, `GetTableNames`) 來聲明使用者操作。`cApplication.cls` 使用 `Private WithEvents vMain As vMain` 來監聽這些事件，並在對應的 `vMain_EventName()` 處理程序中，透過簡單的 `Call Me.PrivateSubName(...)` 語句調用內部私有方法來響應。`vMapping.frm` 和 `cMapping.cls` 也將遵循類似模式。

5.  **資料庫處理 (`AccessDAL.cls`):**
    *   使用後期綁定 (`CreateObject("ADODB.Connection")`) 連接 Access 資料庫，避免強制添加參考。
    *   `Connect` 方法已實現**自動創建資料庫**功能：如果 `DatabasePath` 指定的 `.accdb` 檔案不存在，會嘗試使用 ADOX (後期綁定) 創建一個新的空資料庫。
    *   `GetTableNames` 方法使用 `OpenSchema` 獲取使用者資料表列表，並使用 VBA 內建的 `Collection` 來處理列表。

6.  **服務層 (`PreviewService.cls`, `ImportService.cls` 等):**
    *   `PreviewService.cls` 包含 `ShowPreview` (將 Access 資料表預覽到 Excel) 和 `GetAccessTableNames` (從 `AccessDAL` 獲取資料表列表) 的邏輯。
    *   `ImportService.cls` 包含 `ImportToAccess` (處理 CSV 匯入到 Access 的邏輯，包括調用 `AccessDAL.DropTableIfExists` 和 `AccessDAL.ExecuteSQL`)。
    *   其他服務 (GL/TB/Mapping/Filter) 負責各自領域的業務邏輯。

7.  **`GetTableNames` 與 `ListTable` 互動:**
    *   `vMain.ListTable_Enter()` 觸發 `RaiseEvent GetTableNames`。
    *   `cApplication.vMain_GetTableNames()` 捕獲事件並調用 `cApplication.PopulateTableList()`。
    *   `cApplication.PopulateTableList()` 調用 `PreviewService.GetAccessTableNames()`，獲取列表後清空 `vMain.ListTable`，禁用控制項，填充新列表，最後重新啟用控制項。

8.  **`vMain.ListTable` ComboBox 配置:**
    *   `Style` 屬性設定為 `2 - fmStyleDropDownList`，使用者只能選擇，不能輸入。
    *   已移除程式碼中對 `MatchRequired` 屬性的設置。

9.  **當前狀態與後續步驟:**
    *   `cApplication` 的重構已基本完成，符合控制器職責。
    *   資料庫自動創建功能已實現。
    *   `vMain.ListTable` 的互動問題已解決。
    *   已新增 `FilterCriteria.cls` 和 `FilterService.cls` 的基本結構，為後續篩選功能做準備。
    *   **下一步**是實現欄位映射 (Field Mapping) 功能，涉及 `vMapping.frm`, `cMapping.cls`, 和 `MappingService.cls`。

10. **開發與除錯實踐:**
    *   廣泛使用**後期綁定** (`CreateObject`) 以提高相容性。
    *   使用 `Debug.Print` 在 VBA 立即視窗中輸出調試信息和狀態。
    *   使用 `On Error GoTo Label` 進行錯誤處理，傾向於在較低層（DAL, Service）記錄詳細錯誤，在較高層（Controller）向使用者顯示通用錯誤訊息。
    *   模組頂部使用 `Option Explicit` 強制變數聲明。
    *   類別模組使用 `Private Const MODULE_NAME` 進行標識，方便調試輸出。
    *   類別/表單內部呼叫自身成員時，建議使用 `Me.` 提高可讀性。
    *   遵循一致的事件命名（如 `DoAction`, `GetInformation`）和處理程序命名 (`ObjectName_EventName`)。
    *   為模組和重要程序編寫文件註解。

**當前任務:**
實現欄位映射功能，讓使用者能夠將匯入的 `GL` 和 `TB` 資料表欄位對應到系統預定義的標準欄位。

**本次聚焦目標:**
*   完善 `vMapping.frm` 的介面設計，確保包含所有必要的 ComboBox 控制項來對應 `GLEntity` 和 `TBEntity` 的欄位。
*   在 `vMapping.UserForm_Initialize` 中，實現從 `GL` 和 `TB` 資料表（或透過 Service 層）獲取欄位名稱，並填充到對應的 ComboBox 下拉選單中。
*   實現 `cMapping.cls` 中的事件處理邏輯，特別是 `vMapping_DoConfirm`，用於收集使用者的映射選擇。
*   開發 `MappingService.cls` 中的核心邏輯，用於儲存和應用這些映射關係（初始階段可能只是將映射結果暫存或傳遞給 Controller）。

**執行要求與限制:**

在分析 `#codebase` 中與聚焦目標相關的現有 VBA 程式碼後，請提供具體的重構/修改/新增建議，並**嚴格遵守**以下規則：

1.  **程式碼修改範圍:**
    *   **僅能**修改或新增 `Option Explicit` 關鍵字**之後**的程式碼。
    *   **絕對禁止**修改、刪除或格式化 `Option Explicit` **之前**的任何內容（包括 `VERSION` 行、`BEGIN/END` 塊、`Attribute VB_...` 行等）。這些是 VBA 環境管理模組屬性的必要部分。

2.  **輸出內容:**
    *   **請勿**在回應中重複貼出完整的程式碼檔案或大段未修改的程式碼。
    *   僅提供需要**修改**或**新增**的**關鍵程式碼片段**。

3.  **解釋說明:**
    *   清晰解釋**為何**需要進行這些修改（例如：如何符合新的架構設計？如何實現關注點分離？如何提高可維護性？）。
    *   說明修改後的程式碼片段**如何運作**。

4.  **實施步驟:**
    *   提供**具體、按部就班**的說明，指導使用者如何在 Excel VBA 編輯器中應用建議（例如：「1. 開啟 `vMapping.frm`。 2. 在 `UserForm_Initialize` 事件中添加以下程式碼... 3. 開啟 `cMapping.cls`...」）。

5.  **方案選擇:**
    *   若對於某個重構點存在多種可行的實現方式，請推薦你認為**最佳的方案**，並簡要說明你**選擇該方案的理由**（例如：基於效率、可讀性、可擴展性或 VBA 的限制等）。

6.  **註解要求:**
    *   所有新增或修改的程式碼**必須**包含**繁體中文**註解。
    *   註解風格需保持一致，遵循以下標準：
        *   **模組註解:** 使用 `'=================` 分隔線包圍，說明模組用途。
        *   **行內註解:** 使用 `'` 解釋關鍵或複雜的程式碼行。
        *   註解應清晰、簡潔，解釋意圖而非重複程式碼。