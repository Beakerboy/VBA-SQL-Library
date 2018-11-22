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
Implements iSQLQuery

Private vFields As Variant      'An array of field names
Private sTable As String        'The table name
Private sReturning As String
Private vValues As Variant      'An array of values - OPTIONAL
Private MySelect As SQLSelect  'An SQL statement which can be used to create
                               '  the values - OPTIONAL

'***************************CONSTRUCTOR****************************************
Private Sub Class_Initialize()
End Sub

'***************************LET PROPERTIES*************************************
Property Let Table(sValue As String)
    sTable = sValue
End Property

Property Let Fields(vValue)
    vFields = vValue
End Property

Property Let Values(vValue)
    vValues = vValue
End Property

Property Set setSelect(vValue As SQLSelect)
    Set MySelect = vValue
End Property

Property Let Returning(sValue As String)
    sReturning = sValue
End Property

'***************************FUNCTIONS******************************************

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

