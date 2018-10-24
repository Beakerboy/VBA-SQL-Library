VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************DATABASE CLASS**********************************
'Generalizes multiple database types to allow one SQL library to
'interact with them
'
'*********************************VARIABLES************************************
Private sDSN As String      'DSN
Private sUid As String      'User Name
Private sPwd As String      'Password
Private sType As String     'Type of Database (psql, mssql, mysql)

'***************************CONSTRUCTOR****************************************

'***************************LET PROPERTIES*************************************

Public Property Let DSN(sValue As String)
    sDSN = sValue
End Property

Public Property Let DBType(sValue As String)
    sType = sValue
End Property

Public Property Let UserName(sValue As String)
    sUid = sValue
End Property

Public Property Let Password(sValue As String)
    sPwd = sValue
End Property
'**********************************METHODS*************************************
Private Sub Class_Initialize()

End Sub

Private Function MakeConnectionString() As String
    MakeConnectionString = "DSN=" & sDSN & ";Uid=" & sUid & ";Pwd=" & sPwd & ";"
End Function
'Custom Query
'---------------------------------------------------------------
'should there be multiple functions, one that resurs and one
'that just executes? Return multiple values vs one value?
'
'===============================================================
Public Function Execute(stSQL, Optional return_column = "")
    Set cnt = New ADODB.Connection
    cnt.ConnectionString = MakeConnectionString()
    Set rst = New ADODB.Recordset
    cnt.Open
    return_id = 0

    rst.Open oSQL.ToString, cnt
    If return_column <> "" Then
        Execute = rst(return_column)
    ElseIf TypeOf oSQL Is SQLSelect Then
        Execute = ToArray(rst)
    End If
    
    If CBool(rst.State And adStateOpen) = True Then rst.Close
    Set rst = Nothing
    If CBool(cnt.State And adStateOpen) = True Then cnt.Close
    Set cnt = Nothing
End Function

          Public Function ToArray(rst)
          End Function
          
Public Function PrepareInsert(stSQL, return_column)
    Select Case sType
        Case Is = "mssql"
            PrepareInsert = "SET NOCOUNT ON;" & stSQL & ";" & "SELECT SCOPE_IDENTITY() as " & return_column & ";"
        Case Is = "psql"
            PrepareInsert = stSQL & " RETURNING " & return_column
        Case Is = "mysql"
            PrepareInsert = stSQL & ";SELECT LAST_INSERT_ID() AS " & return_column
    End Select
End Function
