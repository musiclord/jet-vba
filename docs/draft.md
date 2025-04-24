purpose, objectvie: Solution for what? 集中化分工 提升效益 
better solution(優於現況)? Time cost, user experience?
compete? KCT? or else, target aimed for "2025 JET 所內目標"
idea limitation? 自研 core target for?
application -> annoymys or virtual case fro demonstration.
Target to design a prototype, extend to visions. blah blah blah.



根據 0321 & 0328 描述之情境完善以下需求
- Import Data (CSV, TXT)
	- Import to worksheet from PowerQuery
	- Worksheet format
- Column Selection
	- Dropdown List for columns
	- 
- Account Mapping (track to histroy, need state model)
	- Difference, records,
- Validate data Completeness
	- 
- Criteria Matching (5-7 limit fields)
	- Limiation resources, simplest logics ( such as dates, account filtering)
- Export WP (Most simplified format)
	- Worksheet content format, 
	
	
確認 JET 完整流程
JET 資料量 範圍? 格式?

我正在處理Journal Entry Testing Tool(JET)的案件，要將原先的工具遷移至新的解決方案，但目前仍在探索階段；原先的方案是使用Caseware Idea並在此之上開發VBA作為自定義工具，配合使用簡易的Windows Forms作為介面。現在由於授權過期準備要淘汰Idea，因此在評估如何在不依賴Idea的環境下開發JET，目前有兩個方向:
1. 依賴於 .NET 的 Windows Forms Application
2. 依賴於 Office應用程式的 VBA 巨集
由於目前資源大多不能部屬複雜且過多依賴的環境，因此會偏向於VBA巨集，但是受限於底層支援，Excel不支援操作超過一百萬筆資料的內容，而有時候general ledger會超過這個數字，而一般JET在做的事情是以下:
1. 匯入檔案(.csv, .txt, .xlsx)
2. 資料映射，因為客戶提供的資料(例如general ledger或trial balance等)會根據不同系統與公司而有不同的命名，例如"傳票編號"可能在資料欄位命名為"DocumentNo."或其他名稱，為了確保後續ETL的過程正確，會需要制定一個標準化的資料，也就是先經過欄位映射(column mapping)或其他你認為更適合的名稱來描述這個流程。
3. 


pbi每個步驟會記錄於腳本
例如在power query內的操作步驟，會像idea那樣紀錄
欄位配對問題，交由GA處理，因此仍需考慮自定義欄位

office script(typescript)可以處理操作步驟紀錄與還原嗎?
處理JET時有哪些RPA流程，例如導入資料如果相同，是否可以RPA前置作業?


製作definition表，紀錄每個標準化名稱的定義，例如:
傳票編號:標識和追?每一筆會計交易的唯一識別碼

VBA 是 VB 的子集，嵌入在 Microsoft Office 應用程式（如 Excel、Word、Access）中，用於自動化這些應用程式的任務。


C#能處理大量資料嗎?
如何儲存資料表? 如何定義ERM?


ASP.NET Core + WebAssembly + Blazor + Tailwind CSS

當前研究台大課程 JE Tool 的探索 應該是什麼 timesheet 


JET 資料 可能有 資料型別 或欄位不一樣 例如借貸正負數 以不同形式表示
有需要優化的嗎

五個步驟 要儲存狀態 防止步驟中斷 得重新開始


步驟一 降低import 門檻 例如欄位配對
步驟二 mapping科目 將同性質欄位 作為同group或類別 因此篩選時 會以同性質類別 做key
步驟三 正篩 負篩(except, exclude) 日期資料 例如 選出星期六日並排除指定日期


aware of enegagement number, 將各別專案獨立出環境，例如local DB，需各別設計資料表 (避免資料挪用)

datagridview可以預覽資料表內容


criteria: 

完整性 紀錄於帳本上 帳本上須於發票上

2025-3-26
je testing optimiz
sop select by each year, for exmaple 5 criteria, then define as preset has those 5 criteria
customize criteria already in current version jet

date keyin by use input (which method) --> Transform to valid adta format

category is it neeceesary? --> optional column, 

isManual --> excel predifined --> load in for Step1_Check_User_Define_Manual as one column --> 

Criteria --> contains (logic needs to be re-defined) --> 

DropDownList -->Catch Event by data(new item) --> Handle event if catch --> 

question 12 useinput by string --> avoid string length out f forms --> let worksheet as default view and extract as data input.

Document, LineItem--> IF-ELSE to check if data containes necessary  --> for example if LineItem not in GL, then created LineItem group by DocumentNumber

Draw the complete flow chart of "validation" procedure. 

單站式傳票?

AND/OR criteria --> Re-design a QUeryBuilder that can modify not only ONE but MANY criteria at ONCE.

要操作的主體: GL, TB, 
要操作的欄位: 科目, 供應商, 匯率, 
要操作的邏輯: <= , >= , && , + , - , * , / ,