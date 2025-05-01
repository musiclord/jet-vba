VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Z_vMain 
   Caption         =   "Main"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5550
   OleObjectBlob   =   "Z_vMain.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "Z_vMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
'View: Main
'主程式介面，讓使用者執行 :1.匯入檔案,2.驗證資料,3.篩選條件,4.輸出報告
'並設計預覽功能，選擇並於工作表檢視資料表的前1000筆資料。

Public Event DoExit()
Public Event DoImportGL()
Public Event DoImportTB()
Public Event DoPreview()
Public Event OpenMapping()
Public Event GetTableNames()

Private Sub UserForm_Initialize()
    Me.Show (vbModeless)
End Sub

Private Sub ButtonExit_Click()
    RaiseEvent DoExit
End Sub

Private Sub ButtonImportGL_Click()
    RaiseEvent DoImportGL
End Sub

Private Sub ButtonImportTB_Click()
    RaiseEvent DoImportTB
End Sub

Private Sub ButtonMap_Click()
    RaiseEvent OpenMapping
End Sub

Private Sub ButtonPreview_Click()
    ' 將 ListTable選中的資料表載入前1000筆資料至預覽工作表
    RaiseEvent DoPreview
End Sub

Private Sub ListTable_Enter()
    ' 當 ComboBox 獲得焦點時觸發事件
    RaiseEvent GetTableNames
End Sub
