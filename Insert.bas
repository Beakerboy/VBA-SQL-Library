'This Class provides means to insert data into an SQL database
'Either vValues or mySelect must be set. The length of vFields must match the
'  number of items in mySelect or vValues

Private vFields As Variant     'An array of field names
Private sTable As String       'The table name
Private vValues As Variant     'An array of values - OPTIONAL
Private mySelect As SQLSelect  'An SQL statement which can be used to create
                               '  the values - OPTIONAL

'***************************CONSTRUCTOR****************************************
Private Sub Class_Initialize()
    bReturnNewId = False
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
    Set mySelect = vValue
End Property

'***************************FUNCTIONS******************************************

' Create an SQL statement from the object data
Public Function toString()
    returnString = "INSERT INTO " & sTable & " (" & implode(vFields) & ") "
    If mySelect Is Nothing Then
        'Should check that the length of vValues = length of vFields
        returnString = returnString & "VALUES (" & implode(vValues) & ")"
    Else
        'Should check that number of items returned matches vFields
        returnString = returnString & "(" & mySelect.toString & ")"
    End If
    toString = returnString
End Function

Public Function Insert(Optional sReturn = "")
    Set MyDatabase = New Database
    If sReturn = "" Then
        MyDatabase.CustomQuery toString()
    Else
        stSQL = MyDatabase.PrepareInsert(toString, sReturn)
        Insert = MyDatabase.CustomQuery(stSQL, sReturn)
    End If
End Function
