VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Implements iSQLQuery

Private vFields As Variant
Private vValues As Variant
Private sTable As String
Private oSQL As SQLQuery

Private Sub Class_Initialize()
    Set oSQL = New SQLQuery
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
    iSQLQuery_ToString = "UPDATE " & sTable & " SET " & FieldsAndValues & oSQL.WhereString
End Function

Public Sub AddWhere(Field, Value, Optional op As String = "=", Optional GroupType As String = "AND")
    oSQL.AddWhere Field, Value, op, GroupType
End Sub

Public Sub AddWhereGroup(NewWhereGroup As SQLWhereGroup, Optional GroupType As String = "AND")
    oSQL.AddWhereGroup NewWhereGroup, GroupType
End Sub
