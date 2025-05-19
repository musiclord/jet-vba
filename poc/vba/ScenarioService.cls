VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ScenarioService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE As String = "ScenarioService"
Private m_DbPath As String
Private m_dal As AccessDAL
Private m_JE1 As String
Private m_JE2 As String
Private m_JE3 As String

Public Sub Initialize(ByVal DbPath As String, ByVal je1 As String, ByVal je2 As String, ByVal je3 As String)
    Const METHOD As String = "Initialize"
    
    On Error GoTo ErrorHandler
    
    Set m_dal = New AccessDAL
    m_DbPath = DbPath
    m_dal.DatabasePath = m_DbPath
    m_JE1 = je1
    m_JE2 = je2
    m_JE3 = je3
    Debug.Print MODULE & "." & METHOD & " 初始化成功。使用資料庫: " & m_DbPath
    Exit Sub
ErrorHandler:
    Debug.Print "錯誤於 " & MODULE & "." & METHOD & ": " & Err.Description
    Err.Raise Err.Number, MODULE & "." & METHOD, Err.Description ' 將錯誤向上拋給呼叫者
End Sub

Public Function ExecuteScenario1()
    ' 短時間內高頻交易
    Const METHOD As String = "ExecuteScenario1"
    Dim success As Boolean
    Dim JE As String
    JE = m_JE1
    
    If Not m_dal.Connect Then
        Debug.Print MODULE & "." & METHOD & " - 錯誤: 無法連接到案例資料庫: " & m_dal.DatabasePath
        Exit Function
    End If '若無例外則連接成功
    
    ' Execute procedure... start
    
    ' Execute procedure... end
    
    MsgBox "Scenario 1 執行完畢。", vbInformation
    succcess = True
    m_dal.Disconnect
    ExecuteScenario1 = success
    
End Function

Public Function ExecuteScenario2()
    ' 單筆大額款項拆分成多筆小金額款項
    Const METHOD As String = "ExecuteScenario2"
    Dim success As Boolean
    Dim JE As String
    JE = m_JE2
    
    If Not m_dal.Connect Then
        Debug.Print MODULE & "." & METHOD & " - 錯誤: 無法連接到案例資料庫: " & m_dal.DatabasePath
        Exit Function
    End If '若無例外則連接成功
    
    ' Execute procedure... start
    
    ' Execute procedure... end
    
    MsgBox "Scenario 2 執行完畢。", vbInformation
    succcess = True
    m_dal.Disconnect
    ExecuteScenario1 = success
    
End Function

Public Function ExecuteScenario3()
    ' 涉及外幣交易
    Const METHOD As String = "ExecuteScenario3"
    Dim success As Boolean
    Dim JE As String
    JE = m_JE3
    
    If Not m_dal.Connect Then
        Debug.Print MODULE & "." & METHOD & " - 錯誤: 無法連接到案例資料庫: " & m_dal.DatabasePath
        Exit Function
    End If '若無例外則連接成功
    
    ' Execute procedure... start
    
    ' Execute procedure... end
    
    MsgBox "Scenario 3 執行完畢。", vbInformation
    succcess = True
    m_dal.Disconnect
    ExecuteScenario1 = success
    
End Function
    
