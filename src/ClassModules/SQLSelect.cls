VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
Implements iSQLQuery

Private vFields As Variant
Private sTable As String
Private bDistinct As Boolean
Private vGroupBy As Variant
Private oHaving As SQLWhere
Private oHavingGroup As SQLWhereGroup
Private aJoin() As String
Private aOrder() As String
Private oWhere As SQLWhere              'Can have either a SQLWhere or a SQLWhereGroup
Private oWhereGroup As SQLWhereGroup

Property Let Fields(vValue)
    vFields = vValue
End Property

Property Let GroupBy(vValue)
    vGroupBy = vValue
End Property

Private Sub Class_Initialize()
    bDistinct = False
    vGroupBy = Array()
End Sub

'Add a single WHERE clause to the SQL statement
' op is the operation
' value is the value
' EXAMPLE: addwhere 'id' '=', 2
'   is equivalent to "WHERE id=2"
'
Public Sub AddHaving(Field, op As String, Value, Optional GroupType As String = "AND")
    Dim NewHaving As New SQLWhere
    If Not (oHaving Is Nothing) Then
        NewHaving.Create Field, op, Value
        Set oHavingGroup = New SQLWhereGroup
        oHavingGroup.SetGroup oHaving, NewHaving, GroupType
        
        'Clear SQLWhere since we are using SQLWhereGroup instead
        Set oHaving = Nothing
    ElseIf oHavingGroup Is Nothing Then
        Set oHaving = New SQLWhere
        oHaving.Create Field, op, Value
    Else
        NewHaving.Create Field, op, Value
        oHavingGroup.AddWhere NewHaving, GroupType
    End If

End Sub


'Add a single WHERE clause to the SQL statement
' op is the operation
' value is the value
' EXAMPLE: addwhere 'id' '=', 2
'   is equivalent to "WHERE id=2"
'
Public Sub AddWhere(Field, op As String, Value, Optional GroupType As String = "AND")
    Dim NewWhere As New SQLWhere
    If Not (oWhere Is Nothing) Then
        NewWhere.Create Field, op, Value
        Set oWhereGroup = New SQLWhereGroup
        oWhereGroup.SetGroup oWhere, NewWhere, GroupType
        
        'Clear SQLWhere since we are using SQLWhereGroup instead
        Set oWhere = Nothing
    ElseIf oWhereGroup Is Nothing Then
        Set oWhere = New SQLWhere
        oWhere.Create Field, op, Value
    Else
        NewWhere.Create Field, op, Value
        oWhereGroup.AddWhere NewWhere, GroupType
    End If
End Sub

Public Sub AddWhereGroup(NewWhereGroup As SQLWhereGroup, Optional GroupType As String = "AND")
    If oWhere Is Nothing Then
        If oWhereGroup Is Nothing Then
            Set oWhereGroup = NewWhereGroup
        Else
            oWhereGroup.AddWhere NewWhereGroup, GroupType
        End If
    Else
        Set oWhereGroup = New SQLWhereGroup
        oWhereGroup.SetGroup oWhere, NewWhereGroup, GroupType
        Set oWhere = Nothing
    End If
End Sub

Property Let Table(sValue As String)
  sTable = sValue
End Property

Public Function iSQLQuery_ToString() As String
    Dim return_string As String
    If UBound(vFields) < 0 Then
        return_string = ""
     Else
        return_string = "SELECT " & FieldList & " FROM " & sTable
        If Not (oWhere Is Nothing And oWhereGroup Is Nothing) Then
            return_string = return_string & " WHERE " & WhereString
        End If
        If UBound(vGroupBy) > -1 Then
            return_string = return_string & " GROUP BY " & Join(vGroupBy, ", ")
        End If
        If Not (oHaving Is Nothing And oHavingGroup Is Nothing) Then
            return_string = return_string & " HAVING " & HavingString
        End If
    End If
    iSQLQuery_ToString = return_string
End Function

Private Function WhereString() As String
    If Not (oWhere Is Nothing) Then
        WhereString = oWhere.ToString
    Else
        WhereString = oWhereGroup.ToString
    End If
End Function

Private Function HavingString() As String
    If Not (oHaving Is Nothing) Then
        HavingString = oHaving.ToString
    Else
        HavingString = oHavingGroup.ToString
    End If
End Function

Private Function FieldList()
    FieldList = Join(vFields, ", ")
End Function

'Generate a query of the form
'  SELECT sField FROM sTableValue WHERE sProperty = vValue
Public Sub getByProperty(sTableValue As String, sField As String, sProperty As String, vValue)
    Set oWhere = New SQLWhere
    sTable = sTableValue
    vFields = Array(sField)
    oWhere.Create sProperty, "=", vValue
End Sub
