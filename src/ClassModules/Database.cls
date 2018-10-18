VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Database"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This class represents a generic database.
Private sDSN As String      'DSN
Private sUid As String      'User Name
Private sPwd As String      'Password
Private sType As String     'Type of Database (psql, mssql, mysql)

'***************************CONSTRUCTOR****************************************

'***************************LET PROPERTIES*************************************
Public Sub Authenticate(MyCredentials As Credentials)
    sUid = Credentials.Name
    sPwd = Credentials.Pass
End Sub
Private Sub Class_Initialize()
    sDSN = "SVRLIMS"
    sType = "mssql"
End Sub

Private Function MakeConnectionString() As String
    MakeConnectionString = "DSN=" & sDSN & ";Uid=" & sUid & ";Pwd=" & sPwd & ";"
End Function
'return_column the value to return.
'   For sql server, use the value "NEWID" to get the new primary key on an insert statement
Public Function CustomQuery(stSQL, Optional return_column = False)
  Set cnt = New ADODB.Connection
  cnt.ConnectionString = MakeConnectionString()
  Set rst = New ADODB.Recordset
  cnt.Open
  return_id = 0
  
  If return_column <> False Then
    'If return_column = "NEWID" Then
        'If we want the new key value, we need to modify the query
        'Set rst = cnt.Execute(stSQL2)
        'return_id = rst(return_column)
    'Else
      'If we are excecuting a select statement, we submit the query and get the value
      rst.Open stSQL, cnt
      return_id = rst(return_column)
    'End If
  Else
    'If we do not want anything returned, we just submit the SQL
    rst.Open stSQL, cnt
  End If
  If CBool(rst.State And adStateOpen) = True Then rst.Close
  Set rst = Nothing
  If CBool(cnt.State And adStateOpen) = True Then cnt.Close
  Set cnt = Nothing
  If return_column <> False Then CustomQuery = return_id
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
