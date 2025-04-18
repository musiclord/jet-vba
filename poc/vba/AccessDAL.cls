VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessDAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Model: Access DAL
'Description: Data access layer that handles interaction with Access database using ADODB.
Option Explicit

Private Const PROVIDER As String = "Microsoft.ACE.OLEDB.12.0"
Private conn As Object ' ADODB.Connection
Private dbPath As String

