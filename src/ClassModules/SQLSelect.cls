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

Private vFields() As Variant
Private bDistinct As Boolean
Private vGroupBy As Variant
Private oHaving As SQLWhere
Private oHavingGroup As SQLWhereGroup
Private aJoin() As Variant              'An array of arrays. Each array is a 'from_item'
Private aOrder() As Variant
Private oWhere As SQLWhere              'Can have either a SQLWhere or a SQLWhereGroup
Private oWhereGroup As SQLWhereGroup
Private sUnion() As Variant
'Limit
'Offset

Property Let Fields(vValue)
    For Each element In vValue
        AddField element
    Next
End Property

Property Let GroupBy(vValue)
    vGroupBy = vValue
End Property

Private Sub Class_Initialize()
    ReDim aJoin(0)
    bDistinct = False
    vGroupBy = Array()
End Sub

Public Function AddField(sField, Optional sAlias As String = "")
    Dim ArrLen As Integer
    If Not IsArrayAllocated(vFields) Then
        ReDim vFields(0)
        ArrLen = -1
    Else
        ArrLen = UBound(vFields)
        ReDim Preserve vFields(0 To ArrLen + 1)
    End If
    vFields(ArrLen + 1) = Array(sField, sAlias)
End Function

Public Sub addTable(sName As String, Optional sAlias As String = "")
    aJoin(0) = Array("", sName, sAlias, "")
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

Public Sub AddJoin(sType As String, sTable As String, Optional sAlias As String = "", Optional sCondition As String = "")
    'Should Check that sType is either "INNER", "LEFT OUTER", or "RIGHT OUTER"
    Dim JoinLen As Integer
    JoinLen = UBound(aJoin)
    ReDim Preserve aJoin(0 To JoinLen + 1)
    aJoin(JoinLen + 1) = Array(sType, sTable, sAlias, sCondition)
End Sub

Public Sub InnerJoin(sTable As String, sAlias As String, sCondition As String)
    AddJoin "INNER", sTable, sAlias, sCondition
End Sub

Public Sub LeftJoin(sTable, sAlias, Optional sCondition As String = "")
    AddJoin "LEFT OUTER", sTable, sAlias, sCondition
End Sub

'Left Joins are prefered over Right. Please edit query to use a left join
Public Sub RightJoin(sTable, sAlias, Optional sCondition As String = "")
    AddJoin "RIGHT OUTER", sTable, sAlias, sCondition
End Sub

'sType is either "", "ALL", or "DISTINCT"
Public Sub Union(oSelect As SQLSelect, Optional sType = "")
    Dim UnionArray() As Variant
    UnionArray = Array(SQLSelect, sType)
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

Public Sub OrderBy(sField As String, Optional sDirection As String = "ASC")
    If sDirection = "DESC" Then
        sDirection = "DESC"
    Else
        sDirection = "ASC"
    End If
    Dim ArrLen As Integer
    If Not IsArrayAllocated(aOrder) Then
        ReDim aOrder(0)
        ArrLen = -1
    Else
        ArrLen = UBound(aOrder)
    End If
    ReDim Preserve aOrder(0 To ArrLen + 1)
    aOrder(ArrLen + 1) = Array(sField, sDirection)
End Sub

Property Let Table(sValue As String)
    addTable sValue, ""
End Property

Public Function iSQLQuery_ToString() As String
    Dim return_string As String

    If UBound(vFields) < 0 Then
        return_string = ""
    Else
        return_string = "SELECT " & DistinctString & FieldList & " FROM " & _
            JoinString & WhereString & GroupByString & HavingString & OrderByString
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

Private Function JoinString()
    Dim R As Long
    Dim Lines() As String
    Dim Line As String
    ReDim Lines(0 To UBound(aJoin))
    For R = 0 To UBound(aJoin)
        Linearray = aJoin(R)
        Line = ""
        If Linearray(0) <> "" Then
            Line = Linearray(0) & " JOIN "
        End If
        Line = Line & Linearray(1)
        If Linearray(2) <> "" Then
            Line = Line & " " & Linearray(2)
        End If
        If Linearray(3) <> "" Then
            Line = Line & " ON " & Linearray(3)
        End If
        Lines(R) = Line
    Next R
    JoinString = Join(Lines, " ")
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

Private Function OrderByString() As String
    OrderByString = ""
    If IsArrayAllocated(aOrder) Then
        OrderByString = " ORDER BY " & JoinArrayofArrays(aOrder)
    End If
End Function

Private Function FieldList()
    Dim R As Long
    Dim Lines() As String
    Dim Line As String
    ReDim Lines(0 To UBound(vFields))
    For R = 0 To UBound(vFields)
        Linearray = vFields(R)
        Line = Linearray(0)
        If Linearray(1) <> "" Then
            Line = Line & " AS " & Linearray(1)
        End If
        Lines(R) = Line
    Next R
    FieldList = Join(Lines, ", ")
End Function

Public Function UnionString()
    Dim NewSelect As iSQLQuery
    aUnion = UnionArray(0)
    Set NewSelect = aUnion(0)
    UnionString = " UNION " & NewSelect.toString()
End Function

'Generate a query of the form
'  SELECT sField FROM sTableValue WHERE sProperty = vValue
Public Sub getByProperty(sTableValue As String, sField As String, sProperty As String, vValue)
    Set oWhere = New SQLWhere
    addTable sTableValue, ""
    AddField sField
    oWhere.Create sProperty, "=", vValue
End Sub
