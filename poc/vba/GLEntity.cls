VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GLEntity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Model: GL Entity
'Description: Data structure for GL.
Option Explicit

'Necessary
Private strDocumentNumber As String     '傳票號碼
Private strAccountNumber As String      '會計科目編號
Private strAccountName As String        '會計科目名稱
Private strDescription As String        '傳票摘要
Private intLineItem As Integer          '傳票文件項次
Private dblAmount As Double             '傳票金額
Private dtPostDate As Date              '總帳日期

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

Public Property Get Description() As String
    Description = strDescription
End Property
Public Property Let Description(ByVal value As String)
    strDescription = value
End Property

Public Property Get LineItem() As Integer
    LineItem = intLineItem
End Property
Public Property Let LineItem(ByVal value As Integer)
    intLineItem = value
End Property

Public Property Get Amount() As Double
    Amount = dblAmount
End Property
Public Property Let Amount(ByVal value As Double)
    dblAmount = value
End Property

Public Property Get PostDate() As Date
    PostDate = dtPostDate
End Property
Public Property Let PostDate(ByVal value As Date)
    dtPostDate = value
End Property

'Additional properties can be added here as needed
' to represent other fields in the GL entity.

