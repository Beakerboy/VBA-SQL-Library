VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iSQLQuery

Private vFields As Variant
Private vValues As Variant
Private sTable As String
Private oWhere As SQLWhere
Private oWhereGroup As SQLWhereGroup

Private Sub Class_Initialize()
    Set oWhere = New SQLWhere
    Set oWhereGroup = New SQLWhereGroup
End Sub

Property Let Table(sValue As String)
    sTable = sValue
End Property

Property Let Fields(vValue)
    vFields = vValue
End Property

Property Let Values(vValue)
    vValues = vValue
End Property

Private Function FieldsAndValues() As String
    Dim numfields As Integer
    numfields = UBound(vFields)
    
    Dim vFieldsAndValues() As Variant
    ReDim vFieldsAndValues(numfields)
    
    Dim counter As Integer
    counter = 0
    For counter = 0 To numfields
        vFieldsAndValues(counter) = vFields(counter) & "=" & vValues(counter)
    Next counter
    FieldsAndValues = Join(vFieldsAndValues, ", ")
End Function

Public Function iSQLQuery_ToString() As String
    iSQLQuery_ToString = "UPDATE " & sTable & " SET " & FieldsAndValues & WhereString
End Function

Public Sub AddWhere(Field, op As String, Value, Optional GroupType As String = "AND")
    If oWhere.Operator <> "" Then
        Dim NewWhere As SQLWhere
        Set NewWhere = New SQLWhere
        NewWhere.Create Field, op, Value
        oWhereGroup.SetGroup oWhere, NewWhere, GroupType
        
        'Clear SQLWhere since we are using SQLWhereGroup instead
        Set oWhere = New SQLWhere
    Else
        oWhere.Create Field, op, Value
    End If

End Sub

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
