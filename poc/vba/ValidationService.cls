VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValidationService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME As String = "ValidationService"
Private m_DAL As AccessDAL ' 用於儲存 DAL 實例 (若使用依賴注入)

' 方法一：透過 SetDAL 方法注入 DAL
Public Sub SetDAL(ByVal dalInstance As AccessDAL)
    Set m_DAL = dalInstance
End Sub

' 方法二：在需要時內部創建 (較不推薦，耦合度較高)
' Private Function GetDAL() As AccessDAL
'    Set GetDAL = New AccessDAL
'    GetDAL.DatabasePath = ThisWorkbook.Path & "\default.accdb" ' 需要設定路徑
' End Function
