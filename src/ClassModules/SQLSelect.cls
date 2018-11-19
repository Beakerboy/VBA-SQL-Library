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
Private sUnion() As Variant
'Limit
'Offset

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

Public Sub AddJoin(sType, sTable, Optional sAlias = Null, Optional sCondition = Null)
    Dim MyJoin As New SQLJoin
    With MyJoin
        .Type = sType
        .Table = sTable
        .Alias = sAlias
        .Condition = sCondition
    End With
    'Add MyJoin to the Join Array
End Sub

Public Sub InnerJoin(sTable, sAlias, sCondition)
    AddJoin "INNER", sTable, sAlias, sCondition
End Sub

Public Sub LeftJoin(sTable, sAlias, sCondition)
    AddJoin "LEFT OUTER", sTable, sAlias, sCondition
End Sub

'Left Joins are prefered over Right. Please edit query to use a left join
Public Sub RightJoin(sTable, sAlias, sCondition)
    AddJoin "RIGHT OUTER", sTable, sAlias, sCondition
End Sub

'sType is either "", "ALL", or "DISTINCT"
Public Sub Union(oSelect As SQLSelect, Optional sType = "")
    Dim UnionArray(1) As Variant
    UnionArray(0) = sType
    UnionArray(1) = oSelect
    'Add UnionArray to aUnion
End Sub

Public Sub Distinct(Optional bValue = True)
    bDistinct = bValue
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
        return_string = "SELECT " & DistinctString & FieldList & " FROM " & _
            sTable & WhereString & GroupByString & HavingString
    End If
    iSQLQuery_ToString = return_string
End Function

Private Function DistinctString() As String
    Dim sDistinct As String
    sDistinct = ""
    If bDistinct Then
        sDistinct = "DISTINCT "
    End If
    DistinctString = sDistinct
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

Private Function GroupByString() As String
    If UBound(vGroupBy) > -1 Then
        GroupByString = " GROUP BY " & Join(vGroupBy, ", ")
    End If
End Function
Private Function HavingString() As String
    Dim sHaving As String
    If Not (oHaving Is Nothing And oHavingGroup Is Nothing) Then
        If Not (oHaving Is Nothing) Then
            sHaving = oHaving.toString
        Else
            sHaving = oHavingGroup.toString
        End If
        HavingString = " HAVING " & sHaving
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
