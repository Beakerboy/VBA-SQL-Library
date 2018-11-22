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
Private oWhere As SQLCondition              'Can have either a SQLWhere or a SQLWhereGroup
Private oWhereGroup As SQLWhereGroup

Property Let Table(sValue As String)
    sTable = sValue
End Property

'Add a single WHERE clause to the SQL statement
' op is the operation
' value is the value
' EXAMPLE: addwhere 'id', 2
'   is equivalent to "WHERE id=2"
'
Public Sub AddWhere(Field, Value, Optional op As String = "=", Optional GroupType As String = "AND")
    Dim NewWhere As New SQLCondition
    If Not (oWhere Is Nothing) Then
        NewWhere.Create Field, Value, op
        Set oWhereGroup = New SQLWhereGroup
        oWhereGroup.SetGroup oWhere, NewWhere, GroupType
        
        'Clear SQLWhere since we are using SQLWhereGroup instead
        Set oWhere = Nothing
    ElseIf oWhereGroup Is Nothing Then
        Set oWhere = New SQLCondition
        oWhere.Create Field, Value, op
    Else
        NewWhere.Create Field, Value, op
        oWhereGroup.AddWhere NewWhere, GroupType
    End If
End Sub

Public Function iSQLQuery_ToString() As String
    Dim return_string As String
    return_string = "DELETE FROM " & sTable & WhereString
    iSQLQuery_ToString = return_string
End Function

Private Function WhereString() As String
    Dim sWhere As String
    If Not (oWhere Is Nothing And oWhereGroup Is Nothing) Then
        If Not (oWhere Is Nothing) Then
            sWhere = oWhere.toString
        Else
            sWhere = oWhereGroup.toString
        End If
        WhereString = " WHERE " & sWhere
    End If
End Function
