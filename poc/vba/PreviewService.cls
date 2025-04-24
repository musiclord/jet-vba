VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PreviewService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'Model: Preview Service
'Description: Load maximum 1000 records of selected table to worksheet, let user preview data without opening Access.

Private Const MODULE_NAME As String = "PreviewService"
Private Const adDate As Long = 7 ' ADODB Date type constant


Public Function GetAccessTableNames(ByVal accessDbPath As String) As Variant
    ' Service: Retrieves a list of table names from the specified Access database.
    Const METHOD_NAME As String = "GetAccessTableNames"
    Dim dal As AccessDAL
    Dim tableNames As Variant

    On Error GoTo ErrorHandler
    GetAccessTableNames = Empty ' Default return value

    Set dal = New AccessDAL
    dal.DatabasePath = accessDbPath
    tableNames = dal.GetTableNames ' Delegate to DAL

    GetAccessTableNames = tableNames ' Return the variant array

CleanUp:
    Set dal = Nothing
    Exit Function

ErrorHandler:
    Debug.Print "--- " & MODULE_NAME & "." & METHOD_NAME & " Error ---"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Description: " & Err.Description
    ' Let the caller (cApplication) handle the error message to the user
    ' Return Empty to indicate failure
    GetAccessTableNames = Empty
    GoTo CleanUp
End Function

Private Function FindOrCreatePreviewSheet() As Worksheet
    ' 輔助函數：尋找 CodeName 為 "Preview" 的工作表，若不存在則創建並嘗試設定 CodeName。
    Const PREVIEW_CODENAME As String = "Preview"
    Dim ws As Worksheet
    Dim comp As Object ' VBComponent

    On Error Resume Next ' 暫時忽略錯誤，以便檢查 CodeName
    Set FindOrCreatePreviewSheet = Nothing ' 預設傳回值

    ' 1. 嘗試透過 CodeName 尋找現有工作表
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.CodeName, PREVIEW_CODENAME, vbTextCompare) = 0 Then
            Set FindOrCreatePreviewSheet = ws
            Exit Function ' 找到即返回
        End If
    Next ws
    On Error GoTo 0 ' 恢復預設錯誤處理

    ' 2. 如果找不到，則創建新工作表
    Debug.Print "Worksheet with CodeName '" & PREVIEW_CODENAME & "' not found. Creating a new one."
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))

    ' 3. 嘗試設定新工作表的 CodeName (需要啟用 "信任 VBA 專案物件模型存取")
    On Error Resume Next ' 捕捉設定 CodeName 可能發生的錯誤
    Set comp = ThisWorkbook.VBProject.VBComponents(ws.CodeName)
    comp.Name = PREVIEW_CODENAME
    If Err.Number <> 0 Then
        Debug.Print "Warning: Could not set CodeName for the new preview sheet '" & ws.Name & "' to '" & PREVIEW_CODENAME & "'." & _
                    " Please ensure 'Trust access to the VBA project object model' is enabled in Excel Options," & _
                    " or manually set the CodeName in the VBE Properties window."
        Err.Clear
        ' 即使無法設定 CodeName，仍然傳回新建立的工作表供本次使用
    Else
        Debug.Print "Successfully set CodeName for sheet '" & ws.Name & "' to '" & PREVIEW_CODENAME & "'."
    End If
    On Error GoTo 0 ' 恢復預設錯誤處理

    Set FindOrCreatePreviewSheet = ws ' 傳回新建立的工作表

End Function

Public Function ShowPreview(ByVal accessDbPath As String, ByVal tableName As String, ByVal worksheetName As String, ByVal maxRowsToShow As Long) As Boolean
    Const METHOD_NAME As String = "ShowPreview"
    Dim dal As AccessDAL
    Dim rs As Object   ' ADODB.Recordset
    Dim ws As Worksheet
    Dim sql As String
    Dim i As Long
    ' Dim targetSheetExists As Boolean ' 不再需要
    Dim fieldCount As Long
    Dim recordCount As Long
    Dim previewSheetName As String

    On Error GoTo ErrorHandler
    ShowPreview = False ' Default to failure

    ' --- Excel Worksheet Handling ---
    Application.ScreenUpdating = False ' Prevent screen flicker

    ' 1. 獲取或創建 CodeName 為 "Preview" 的工作表
    Set ws = FindOrCreatePreviewSheet()
    If ws Is Nothing Then
        ' 如果 FindOrCreatePreviewSheet 失敗 (極少數情況)，則引發錯誤
        Err.Raise vbObjectError + 518, MODULE_NAME & "." & METHOD_NAME, "Failed to find or create the 'Preview' worksheet."
    End If

    ' 2. 清空工作表並設定顯示名稱
    ws.Cells.Clear
    previewSheetName = "#" & tableName ' 設定目標顯示名稱
    On Error Resume Next ' 處理可能的名稱衝突或無效字元
    ws.Name = previewSheetName
    If Err.Number <> 0 Then
        Debug.Print "Warning: Could not rename worksheet to '" & previewSheetName & "'. Error: " & Err.Description
        Err.Clear
        ' 即使重命名失敗，仍繼續執行
    End If
    On Error GoTo ErrorHandler ' 恢復錯誤處理
    ws.Activate ' 啟用工作表以供使用者查看
    Debug.Print "Using worksheet '" & ws.Name & "' (CodeName: " & ws.CodeName & ") for preview."
    ' --- End Excel Worksheet Handling ---

    ' 3. 建立 DAL 並取得資料 (原步驟 2)
    Set dal = New AccessDAL
    dal.DatabasePath = accessDbPath

    ' 4. 建立 SQL 語句 (原步驟 3)
    If maxRowsToShow > 0 Then
        sql = "SELECT TOP " & maxRowsToShow & " * FROM [" & tableName & "]"
    Else
        sql = "SELECT * FROM [" & tableName & "]"
    End If

    ' 5. 從 DAL 取得 Recordset (原步驟 4)
    Set rs = dal.GetRecordset(sql, disconnectRecordset:=True)

    ' 6. 檢查 Recordset 是否有效 (原步驟 5)
    If rs Is Nothing Then
        Err.Raise vbObjectError + 517, MODULE_NAME & "." & METHOD_NAME, "Failed to retrieve data from Access table '" & tableName & "'."
    End If

    ' --- 將資料寫入 Excel --- (原步驟 6-9)
    If Not (rs.EOF And rs.BOF) Then ' 檢查是否有記錄
        rs.MoveFirst
        fieldCount = rs.Fields.Count
        recordCount = rs.recordCount ' 從離線 Recordset 取得 RecordCount

        ' 寫入標頭
        For i = 0 To fieldCount - 1
            ws.Cells(1, i + 1).value = rs.Fields(i).Name
        Next i

        ' 複製資料
        If recordCount > 0 Then
            ws.Range("A2").CopyFromRecordset rs, recordCount
        End If

        ' 格式化日期 (可選)
        For i = 0 To fieldCount - 1
            If rs.Fields(i).Type = adDate Then
                If recordCount > 0 Then
                     ws.Range("A2").Offset(0, i).Resize(recordCount, 1).NumberFormat = "yyyy/mm/dd"
                End If
            End If
        Next i

        ' 自動調整欄寬
        ws.Columns.AutoFit
    Else
        ' 處理空資料表：如果可能，寫入標頭
        If Not rs Is Nothing Then
             If rs.Fields.Count > 0 Then
                 For i = 0 To rs.Fields.Count - 1
                    ws.Cells(1, i + 1).value = rs.Fields(i).Name
                Next i
                ws.Columns.AutoFit
             Else
                 ws.Cells(1, 1).value = "Table [" & tableName & "] exists but has no columns or could not be read."
             End If
        Else
             ws.Cells(1, 1).value = "Could not read table [" & tableName & "]."
        End If
        ' 移除 MsgBox，讓 Controller 決定是否顯示訊息
        ' MsgBox "No data to load from Access table [" & tableName & "].", vbInformation
        Debug.Print "No data to load from Access table [" & tableName & "]."
    End If
    ' --- End Write data to Excel ---

    ' 7. 成功 (原步驟 10)
    ShowPreview = True

CleanUp:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    Set rs = Nothing
    Set dal = Nothing ' DAL 的 Terminate 會處理連線中斷
    Set ws = Nothing
    Application.ScreenUpdating = True ' 恢復螢幕更新
    Exit Function

ErrorHandler:
    ShowPreview = False
    Debug.Print "--- " & MODULE_NAME & "." & METHOD_NAME & " Error ---"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Source: " & Err.Source
    Debug.Print "Error Description: " & Err.Description
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ' 將錯誤引發給呼叫者 (cApplication)
    Err.Raise Err.Number, MODULE_NAME & "." & METHOD_NAME, Err.Description
    GoTo CleanUp
End Function
