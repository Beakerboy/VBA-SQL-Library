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
Private vFields As Variant
Private sTable As String
Private bDistinct As Boolean
Private aHaving() As String
Private aWhere() As String
Private aJoin() As String
Private aOrder() As String
Private oWhere As SQLWhere              'Can have either a SQLWhere or a SQLWhereGroup
Private oWhereGroup As SQLWhereGroup

Property Let Fields(vValue)
  vFields = vValue
End Property

Private Sub Class_Initialize()
    Set oWhere = New SQLWhere
    Set oWhereGroup = New SQLWhereGroup
    bDistinct = False
End Sub

'Add a single WHERE clause to the SQL statement
' op is the operation
' value is the value
' EXAMPLE: addwhere 'id' '=', 2
'   is equivalent to "WHERE id=2"
'

Public Sub addWhere(Field, op As String, Value)
    oWhere.Create Field, op, Value
End Sub

Property Let Table(sValue As String)
  sTable = sValue
End Property

Public Function ToString()
    If UBound(vFields) < 0 Then
        ToString = ""
     Else
        ToString = "SELECT " & FieldList & " FROM " & sTable & " WHERE " & oWhere.ToString
    End If
End Function

Private Function WhereString() As String
    If oWhere.Operator <> "" Then
        WhereString = oWhere.ToString
    Else
        WhereString = oWhereGroup.ToString
    End If
End Function

Private Function FieldList()
    FieldList = Join(vFields, ", ")
End Function

'Generate a query of the form
'  SELECT sField FROM sTableValue WHERE sProperty = vValue
Public Sub getByProperty(sTableValue As String, sField As String, sProperty As String, vValue)
    sTable = sTableValue
    vFields = Array(sField)
    oWhere.Create sProperty, "=", vValue
End Sub
