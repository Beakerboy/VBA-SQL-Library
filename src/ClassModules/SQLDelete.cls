VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements iSQLQuery

Private sTable As String
Private oSQL As SQLQuery

Private Sub Class_Initialize()
    Set oSQL = New SQLQuery
End Sub

Property Let Table(sValue As String)
    sTable = sValue
End Property

Public Sub AddArgument(sName As String, vValue)
    oSQL.AddArgument sName, vValue
End Sub

Public Sub AddWhere(Field, Value, Optional op As String = "=", Optional GroupType As String = "AND")
    oSQL.AddWhere Field, Value, op, GroupType
End Sub

Public Sub AddWhereGroup(NewWhereGroup As SQLWhereGroup, Optional GroupType As String = "AND")
    oSQL.AddWhereGroup NewWhereGroup, GroupType
End Sub

Public Function iSQLQuery_ToString() As String
    Dim return_string As String
    return_string = "DELETE FROM " & sTable & oSQL.WhereString
    iSQLQuery_ToString = oSQL.ReplaceArguments(return_string)
End Function
