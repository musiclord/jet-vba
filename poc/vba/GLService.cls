VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GLService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'Model: GL Service
'Description: All GL related procedure, such as AddLineItem()

Public Sub AddLineItem()
    ' 對資料表 GL 新增欄位 [LineItem(傳票文件項次)]
    ' 用作資料增強，將欄位映射後的資料表 GL 對 [DocumentNumber] 分組後設置遞增的數字
End Sub
