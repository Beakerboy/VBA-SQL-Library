VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private vFields As Variant
Private sTable As String
Private bDistinct As Boolean
Private aHaving() As String
Private aWhere() As String
Private aJoin() As String
Private aOrder() As String

'An array
'DWhere['id']=[0=> '=', 1=>4] is equivalent to "WHERE id=2"
'ToDo, do not use Dictionary, because we may want multiple WHEREs in the same key
Private DWhere As Dictionary

Property Let Fields(vValue)
  vFields = vValue
End Property

Private Sub Class_Initialize()
    Set DWhere = New Dictionary
    bDistinct = False
End Sub

'Add a WHERE clause to the SQL statement
' op is the operation
' value is the value
' EXAMPLE: addwhere 'id' '=', 2
'   is equivalent to "WHERE id=2"
'
'What if the field is used in two conditions? Can a dictionary handle two
'  records with the same Key?

Public Sub addWhere(Field, op, Value)
    Dim arrayProp(1) As Variant
    arrayProp(0) = op
    arrayProp(1) = Value
    DWhere.Add Key:=Field, Item:=arrayProp
End Sub

Property Let Table(sValue As String)
  sTable = sValue
End Property

Public Function toString()
    If UBound(vFields) < 0 Then
        toString = ""
     Else
        toString = "SELECT " & FieldList & " FROM " & sTable & " " & ImplodeWhere
    End If
End Function

Private Function FieldList()
    FieldList = Join(vFields, ", ")
End Function

'Create a string of where clauses
'Currently only works for one where statement.
'How to handle ANDs and ORs
Private Function ImplodeWhere()
    sImplode = ""
    If DWhere.count > 0 Then
        sImplode = "WHERE "
        For Each Key In DWhere.Keys
            arrayProp = DWhere(Key)
            sImplode = sImplode & Key & arrayProp(0) & arrayProp(1)
        Next Key
        sImplode = sImplode & " "
    End If
    ImplodeWhere = sImplode
End Function

'Generate a query of the form
'  SELECT sField FROM sTableValue WHERE sProperty = vValue
Public Sub getByProperty(sTableValue As String, sField As String, sProperty As String, vValue)
    sTable = sTableValue
    vFields = Array(sField)
    Dim arrayProp(1) As Variant
    arrayProp(0) = "="
    arrayProp(1) = vValue
    DWhere.Add Key:=sProperty, Item:=arrayProp
End Sub
