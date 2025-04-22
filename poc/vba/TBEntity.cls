VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TBEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'Model: TB entity
'Description: Data structure and interface for GL.

'Necessary
Private strDocumentNumber As String     '傳票號碼
Private strAccountNumber As String      '會計科目代號
Private strAccountName As String        '會計科目名稱
Private dblOpeningBalance As Double     '期初金額
Private dblClosingBalance As Double     '期末金額
Private dblNetChange As Double          '本期變動
Private intLineItem As Integer          '項次編號

Public Property Get DocumentNumber() As String
    DocumentNumber = strDocumentNumber
End Property
Public Property Let DocumentNumber(ByVal value As String)
    strDocumentNumber = value
End Property

Public Property Get AccountNumber() As String
    AccountNumber = strAccountNumber
End Property
Public Property Let AccountNumber(ByVal value As String)
    strAccountNumber = value
End Property

Public Property Get AccountName() As String
    AccountName = strAccountName
End Property
Public Property Let AccountName(ByVal value As String)
    strAccountName = value
End Property

Public Property Get OpeningBalance() As Double
    OpeningBalance = dblOpeningBalance
End Property
Public Property Let OpeningBalance(ByVal value As Double)
    dblOpeningBalance = value
End Property

Public Property Get ClosingBalance() As Double
    ClosingBalance = dblClosingBalance
End Property
Public Property Let ClosingBalance(ByVal value As Double)
    dblClosingBalance = value
End Property

Public Property Get NetChange() As Double
    NetChange = dblNetChange
End Property
Public Property Let NetChange(ByVal value As Double)
    dblNetChange = value
End Property

Public Property Get LineItem() As Integer
    LineItem = intLineItem
End Property
Public Property Let LineItem(ByVal value As Integer)
    intLineItem = value
End Property
