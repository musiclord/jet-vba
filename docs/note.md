# 待開發的程序:
- 匯入檔案
    - 確保 GL 和 TB 的資料表正確匯入，並在 Access 資料庫中存為對應的資料表
- 驗證資料
    - 完整性測試
        - 定義完整性測試的邏輯和流程
        - 列出該測試的需求和資料
        - 實現完整性測試
    - 借貸不平
        - 定義借貸不平測試的邏輯和流程
        - 列出該測試的需求和資料
        - 實現借貸不平測試
    - 科目配對
        - 定義科目配對的邏輯和流程
        - 列出該功能的需求和資料
        - 實現科目配對的功能
        - 實現科目配對的介面
- 篩選條件
    - 實現

# 類別描述
- **mod_Utility.bas**
    - 通用功能工具，作為開放函數讓所有類別存取使用。

- **vMain.frm**
    - 主程式介面，讓使用者執行 :1.匯入檔案,2.驗證資料,3.篩選條件,4.輸出報告；並且設計預覽功能，選擇並於工作表檢視資料表的前1000筆資料。

- **vMapping.frm**
    - 映射欄位介面，讓使用者將匯入的資料，藉由操作下拉式選單來映射至正確對應的欄位名稱。

- **cApplication.cls**
    - 該VBA應用程序的主要控制器，控制 `vMain` 介面的程序，並呼叫對應函數。

- **cMapping.cls**
    - 映射欄位的控制器，控制 `vMapping` 介面的程序。

- **AccessDAL.cls**
    - 負責所有與 Microsoft Access 資料庫的互動，封裝 ADO/ADOX 連線，執行 SQL 語句 ( *如 SELECT, INSERT, UPDATE, DELETE* ) 以及操作資料庫物件 ( *如檢查表格是否存在、建立表格、新增欄位等* ) 的底層細節。

- **GLEntity.cls**
    - `GL` (General Ledger) 資料實體類別，定義 `GL` 資料結構及驗證。

- **TBEntity.cls**
    - `TB` (Trial Balance) 資料實體類別，定義 `TB` 資料結構及驗證。

- **ImportService.cls**
    - 處理檔案匯入的服務層，負責從本機匯入資料檔案(如CSV、Excel)至Access資料庫，與AccessDAL協作完成資料存取操作。

- **MappingService.cls**
    - 處理欄位映射 (Field Mapping) 邏輯的服務層，用於標準化欄位名稱，儲存與管理映射關係。

- **PreviewService.cls**
    - 處理預覽資料表的服務層，負責將資料庫中指定資料表的前1000筆資料，載入至指定的工作表。

# 流程描述
- 匯入檔案
    - 匯入 `GL.csv` 至 Access 為資料表 (table) `GL`
    - 匯入 `TB.csv` 至 Access 為資料表 `TB`
    - 預覽資料表 `GL` 於工作表 (worksheet) `GL`
    - 預覽資料表 `TB` 於工作表 `TB`
    - 對資料表 `GL` 和 `TB` 以下拉式選單配對欄位 (Field Mapping)
    - 對資料表 `GL` 增加欄位 [文件項次] 做資料增強
    - 標準化資料表 `GL` 和 `TB` 為 `GL#` 和 `TB#` 
- 驗證資料
    - 完整性驗證 (Completeness)
    - 借貸不平驗證 (Document Not Balance)
    - 資料元素攸關驗證 (Relevant Data Elements)
    - 科目配對 (Account Mapping)
- 篩選條件
    - 設定基本篩選條件 (Criteria Selection)
    - 組合篩選條件設為一組篩選配置
    - 依照篩選配置進行篩選作業
    - 預覽篩選結果於工作表
- 輸出報告
    - 將驗證結果 `#Completeness` 填至範本工作表 `Validation.xlsx`
    - 將篩選結果 `#Filtered` 填至範本工作表 `Filtered_Result.xlsx`

# 程序描述


