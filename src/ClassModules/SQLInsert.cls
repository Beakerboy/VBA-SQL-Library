VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'This Class provides means to insert data into an SQL database
'Either vValues or mySelect must be set. The length of vFields must match the
'  number of items in mySelect or vValues
Option Explicit

' Class: SQLInsert
' A SQL Insert query
Implements iSQLQuery

Private vFields As Variant      'An array of field names
Private sTable As String        'The table name
Private sReturning As String
Private vValues As Variant      'An array of values - OPTIONAL
Private MySelect As SQLSelect  'An SQL statement which can be used to create
                               '  the values - OPTIONAL

' Property: Table
Property Let Table(sValue As String)
    sTable = sValue
End Property

' Property: Fields
Property Let Fields(vValue)
    vFields = vValue
End Property

' Property: Values
Property Let Values(vValue)
    vValues = vValue
End Property

' Property: From
' Set the SELECT query that is used in the INSERT statement.
' Parameter:
'   vValue - A <SQLSelect> Object to use as in the INSERT query
Property Set From(vValue As SQLSelect)
    Set MySelect = vValue
End Property

' Property: Returning
Property Let Returning(sValue As String)
    sReturning = sValue
End Property

' Constructor: Class_Initialize
' Initialize class members
Private Sub Class_Initialize()
End Sub
  
' Function: iSQLQuery_ToString
' Create an SQL statement from the object data
Public Function iSQLQuery_ToString() As String
    Dim MyQuery As iSQLQuery
    Set MyQuery = MySelect
    Dim ReturnString As String
    ReturnString = "INSERT INTO " & sTable & ImplodeFields
    If MySelect Is Nothing Then
        'Should check that the length of vValues = length of vFields.
        
        ReturnString = ReturnString & " VALUES" & ImplodeValues
    Else
        'Should check that number of items returned matches vFields
        ReturnString = ReturnString & " (" & MyQuery.toString & ")"
    End If
    If sReturning <> "" Then
        ReturnString = ReturnString & " RETURNING " & sReturning
    End If
    iSQLQuery_ToString = ReturnString
End Function

    
'Join the array of values into a string
Private Function ImplodeValues() As String
    If IsArray(vValues(0)) Then
        ImplodeValues = " (" & JoinArrayofArrays(vValues, ", ", "), (") & ")"
    Else
        ImplodeValues = " (" & Join(vValues, ", ") & ")"
    End If
End Function

Private Function ImplodeFields() As String
    Dim ReturnString As String
    ReturnString = " (" & Join(vFields, ", ") & ")"
    ImplodeFields = ReturnString
End Function
