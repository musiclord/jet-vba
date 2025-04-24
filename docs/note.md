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
    - 操作下拉式選單配對欄位 (Field Mapping)
    - 增加欄位 [文件項次] 做資料增強
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
	1.	匯入GL.csv至Access資料庫
		1.1	檢查根目錄是否存在同名的資料庫
		1.2	檢查資料庫是否存在同名的資料表
		1.3	以正確的編碼及資料型別匯入資料表
	2. 讀取Access資料庫，將新匯入(或已存在同名)的資料表載入至工作表
		2.1	檢查是否存在同名的工作表
		2.2	覆寫舊資料，匯入新資料
		2.3	將工作表作為查詢結果的View，以1000筆資料為限
	3. 使用者根據vMapping的下拉式選單，將 mGLEntity 預先定義的欄位名稱，配對至匯入的GL資料表
		3.1 在 mGLEntity 中記錄資料表的欄位名稱
		3.2 將使用者重新配對的欄位記錄在 mGLEntity 並連結舊欄位名稱，使得後續ETL可以正確的操映射後的欄位
		3.3 c/ c
	4. 做資料增強，將GL資料表添加"LineItem"欄位，用來區別重複相同傳票號碼的不同分錄
		4.1 呼叫 mGLEntity 來以定義好的資料結構，呼叫 cAccess 來操作具體的查詢語句
		4.2 使用SQL新增欄位 "LineItem" 對每個相同 "傳票號碼" 的分錄資料進行逐一增加的順序號
		4.3 資料庫的GL資料表更新了
	5. 檢視更新後的資料表
		5.1 使用者點擊 vMain 的 ListTable 選單來選取想預覽的資料表
		5.2 使用者點擊 vMain 的 ButtonPreview 來觸發 DoPreview事件
		5.3 創建或更新(若已存在)工作表 "Preview"
		5.4 將要預覽資料表的查詢結果載入至工作表


在 vMain.frm 中有按鈕 ButtonPreview 在點擊後會觸發 DoPreview 事件並由 cApplication.cls 的 vMain_DoPreview 處理該事件，在流程上是在查詢資料庫並傳回至工作表，因此他應該能重複使用，將 mImport.LoadToExcel 改寫，使用與 vMain_DoPreview 相同的運作邏輯。請問若要新增一個函數來處理上述功能，他應該歸屬於哪個類別? 如何調用?

為了完成上述內容，請先幫我設計 #file:cAccess.cls  並告訴我其他使用資料庫的函數與類別要如何調用? 另說明有哪些部份需要更動? 例如 #file:mImport.cls  的方法、或是 #file:mGLService.cls  尚未完成的 AddLineItem() 邏輯、以及MVC架構中的 #file:mGLEntity.cls  等。在思考的過程中，請以你認為最優化且簡潔的代碼為主，避免過度設計複雜的邏輯，提升代碼易讀性，並且你需要以當前專案的全局來分析代碼風格是否符合Excel的VBA。