VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This Class provides means to insert data into an SQL database
'Either vValues or mySelect must be set. The length of vFields must match the
'  number of items in mySelect or vValues
Option Explicit
Private vFields As Variant     'An array of field names
Private sTable As String       'The table name
Private vValues As Variant     'An array of values - OPTIONAL
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

Property Let setSelect(vValue)
    Set MySelect = vValue
End Property

'***************************FUNCTIONS******************************************

'Join the array of values into a string
Private Function ImplodeValues()
    ImplodeValues = " (" & Join(vValues, ", ") & ") "
End Function

Private Function ImplodeFields()
    if getDimension (vValues) = 2 Then
        ImplodeFields = "(" & Join2D(arr2, ", ", "), (") & ")"
    else
        ImplodeFields = " (" & Join(vFields, ", ") & ") "
    end if
End Function

' Create an SQL statement from the object data
Public Function ToString()
    Dim returnstring As String
    returnstring = "INSERT INTO " & sTable & ImplodeFields
    If MySelect Is Nothing Then
        'Should check that the length of vValues = length of vFields.
        
        returnstring = returnstring & "VALUES" & ImplodeValues
    Else
        'Should check that number of items returned matches vFields
        returnstring = returnstring & "(" & MySelect.ToString & ")"
    End If
    ToString = returnstring
End Function
