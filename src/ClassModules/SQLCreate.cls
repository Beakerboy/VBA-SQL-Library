VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements iSQLQuery

Private oSQL As SQLQuery

Private Sub Class_Initialize()
    Set oSQL = New SQLQuery
End Sub

Public Sub AddArgument(sName As String, vValue)
    oSQL.AddArgument sName, vValue
End Sub

Public Function iSQLQuery_ToString() As String
    Dim return_string As String
    iSQLQuery_ToString = oSQL.ReplaceArguments(return_string)
End Function
