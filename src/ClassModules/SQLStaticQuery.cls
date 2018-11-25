VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLStaticQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements iSQLQuery

Private sQuery As String
Private oSQL As SQLQuery

Private Sub Class_Initialize()
    Set oSQL = New SQLQuery
End Sub

Public Property Let Query(sValue)
    sQuery = sValue
End Property

Public Function iSQLQuery_ToString() As String
    iSQLQuery_ToString = oSQL.ReplaceArguments(sQuery)
End Function

Public Sub AddArgument(sName As String, vValue)
    oSQL.AddArgument sName, vValue
End Sub

Public Sub ClearArguments()
    oSQL.ClearArguments
End Sub


