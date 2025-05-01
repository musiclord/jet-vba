VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} vGLConfig 
   Caption         =   "GL Configuration"
   ClientHeight    =   8415.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11760
   OleObjectBlob   =   "vGLConfig.frx":0000
   StartUpPosition =   1  '所屬視窗中央
End
Attribute VB_Name = "vGLConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'===== 表單資訊 =====
'名稱: vGLConfig
'描述: 設定總帳 (GL) 資料匯入的欄位對應和相關選項。
'
'===== 表單控制項 =====
'下拉式選單 (ComboBox):
'  - cboDocumentNo: 對應來源檔中的「傳票號碼」欄位
'  - cboLineItem: 對應來源檔中的「傳票項次」欄位
'  - cbotPostDate: 對應來源檔中的「總帳日期」欄位
'  - cboAccountNo: 對應來源檔中的「會計科目編號」欄位
'  - cboAccountName: 對應來源檔中的「會計科目名稱」欄位
'  - cboDescription: 對應來源檔中的「傳票摘要」欄位
'  - cboMethod: 選擇總帳金額的處理方式 (例如：單一金額欄位、借貸分開欄位等)
'  - cboAmount: 對應來源檔中的「傳票金額」欄位 (若 cboMethod 選擇單一金額)
'  - cboDebit: 對應來源檔中的「借方金額」欄位 (若 cboMethod 選擇借貸分開)
'  - cboCredit: 對應來源檔中的「貸方金額」欄位 (若 cboMethod 選擇借貸分開)
'  - cboDrCr: 對應來源檔中的「借貸類別」欄位 (若 cboMethod 選擇借貸標記)
'  - cboDebitFlag: 輸入代表借方的標記值 (若 cboMethod 選擇借貸標記)
'  - cboPostingStatus: 對應來源檔中的「過帳狀態」欄位
'  - cboPostedCode: 輸入代表已過帳的狀態代碼
'  - cboPosted: 選擇是否僅篩選已過帳的資料 (通常是 CheckBox 或 OptionButton，若為 ComboBox 則可能提供 Yes/No 選項)
'
'命令按鈕 (CommandButton):
'  - btnImport: 觸發匯入 GL 資料的事件 (RaiseEvent DoImport)
'  - btnPreview: 觸發預覽 GL 資料表的事件 (RaiseEvent DoPreview)
'  - btnConfirm: 確認設定並觸發 DoConfirm 事件 (RaiseEvent DoConfirm)
'  - btnExit: 關閉表單並觸發 DoExit 事件 (RaiseEvent DoExit)

Public Event DoImport()
Public Event DoPreview()
Public Event DoConfirm()
Public Event DoExit()

Private Sub btnConfirm_Click()
    RaiseEvent DoConfirm
End Sub

Private Sub btnExit_Click()
    RaiseEvent DoExit
End Sub

Private Sub btnImport_Click()
    RaiseEvent DoImport
End Sub

Private Sub btnPreview_Click()
    RaiseEvent DoPreview
End Sub

Private Sub cboCredit_Change()

End Sub

Private Sub lblDocumentNo_Click()

End Sub

Private Sub UserForm_Initialize()

End Sub

