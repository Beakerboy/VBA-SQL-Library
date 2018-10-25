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
Private bTestMode As Boolean

'***************************CONSTRUCTOR****************************************
Private Sub Class_Initialize()
    bTestMode = False
End Sub
'***************************LET PROPERTIES*************************************

Public Property Let DSN(sValue As String)
    sDSN = sValue
End Property

Public Property Let DBType(sValue As String)
    sType = sValue
End Property

Public Property Let Username(sValue As String)
    sUid = sValue
End Property

Public Property Let Password(sValue As String)
    sPwd = sValue
End Property
Public Property Let TestMode(bValue As Boolean)
    bTestMode = bValue
End Property
'**********************************METHODS*************************************

Private Function MakeConnectionString() As String
    MakeConnectionString = "DSN=" & sDSN & ";Uid=" & sUid & ";Pwd=" & sPwd & ";"
End Function

'Execute
'
'Send a query to the database and return a result if needed
'
Public Function Execute(oSQL, Optional return_column = "")
    If VarType(oSQL) = 9 Then
        sSQL = oSQL.toString
        If TypeOf oSQL Is SQLSelect Then
            SQLType = "select"
        Else
            SQLType = "insert"
        End If
    Else
        sSQL = oSQL
    End If
    Set cnt = New ADODB.Connection
    cnt.ConnectionString = MakeConnectionString()
    Set rst = New ADODB.Recordset
    cnt.Open
    return_id = 0
    
    If bTestMode Then
        MsgBox sSQL
        If SQLType = "select" Then
            rst.Open sSQL, cnt
            If return_column <> "" Then
                return_value = rst(return_column)
                MsgBox return_column & " = " & return_value
                Execute = return_value
            ElseIf SQLType = "select" Then
                Execute = ToArray(rst)
            End If
        Else
            Execute = 42
        End If
    Else
        rst.Open sSQL, cnt
        If return_column <> "" Then
            Execute = rst(return_column)
        ElseIf SQLType = "select" Then
            Execute = ToArray(rst)
        End If
    End If
    
    If CBool(rst.State And adStateOpen) = True Then rst.Close
    Set rst = Nothing
    If CBool(cnt.State And adStateOpen) = True Then cnt.Close
    Set cnt = Nothing
End Function

'InsertGetNewID
'
'Prepare an SQLInsert object to be executed. Execute the query and return the result
'
Public Function InsertGetNewId(oSQL As SQLInsert, Optional ReturnColumn = "")
    If ReturnColumn = "" Then ReturnColumn = "somethingunique"
    InsertGetNewId = Execute(PrepareInsert(oSQL, ReturnColumn), ReturnColumn)
End Function

Public Function ToArray(rst)
End Function
          
Public Function PrepareInsert(ByVal oSQL, return_column)
    stSQL = oSQL.toString
    Select Case sType
        Case Is = "mssql"
            PrepareInsert = "SET NOCOUNT ON;" & stSQL & ";" & "SELECT SCOPE_IDENTITY() as " & return_column & ";"
        Case Is = "psql"
            PrepareInsert = stSQL & " RETURNING " & return_column
        Case Is = "mysql"
            PrepareInsert = stSQL & ";SELECT LAST_INSERT_ID() AS " & return_column
    End Select
End Function
