VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************DATABASE CLASS**********************************
'Generalizes multiple database types to allow one SQL library to
'interact with them
'
Option Explicit
'*********************************VARIABLES************************************
Private sDSN As String          'DSN
Private sUid As String          'User Name
Private sPwd As String          'Password
Private cnt                     'Connection
Private rst As iSQLRecordset    'Recordset
Private sType As String         'Type of Database (psql, mssql, mysql)

'***************************CONSTRUCTOR****************************************
Private Sub Class_Initialize()
    Set rst = New SQLRecordset
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

Public Property Set Recordset(MyRecordset As iSQLRecordset)
    Set rst = Nothing
    Set rst = MyRecordset
End Property

'**********************************METHODS*************************************

Private Function MakeConnectionString() As String
    MakeConnectionString = "DSN=" & sDSN & ";Uid=" & sUid & ";Pwd=" & sPwd & ";"
End Function

'Execute
'
'Send a query to the database and return a result if needed
'
Public Function Execute(oSQL As iSQLQuery, Optional return_column = "")
    Dim sSQL As String
    Dim SQLType As String
    
    sSQL = oSQL.ToString

    Dim cnt As ADODB.Connection
    Set cnt = New ADODB.Connection
    cnt.ConnectionString = MakeConnectionString()
    cnt.Open
    
    rst.OpenRecordset sSQL, cnt              'Execute the query
    If return_column <> "" Then
        Execute = rst.GetValue(return_column)
    Else
        'If it is a select Query, return the recordset as an array
    End If
    
    'Bitwise And the State property with The adStateOpen constant
    If CBool(rst.State And adStateOpen) = True Then rst.CloseRecordset
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
          
Public Function PrepareInsert(ByVal oSQL As SQLInsert, return_column)
    Dim stSQL As String
    Dim MyQuery As iSQLQuery
    Set MyQuery = oSQL
    stSQL = MyQuery.ToString
    Dim MySQLGroup As New SQLQueryGroup
    Select Case sType
        Case Is = "mssql"
            MySQLGroup.AddQuery "SET NOCOUNT ON"
            MySQLGroup.AddQuery stSQL
            MySQLGroup.AddQuery "SELECT SCOPE_IDENTITY() as " & return_column
            Set PrepareInsert = MySQLGroup
        Case Is = "psql"
            oSQL.Returning = return_column
            Set PrepareInsert = oSQL
        Case Is = "mysql"
            MySQLGroup.AddQuery stSQL
            MySQLGroup.AddQuery "SELECT LAST_INSERT_ID() AS " & return_column
            Set PrepareInsert = MySQLGroup
    End Select
End Function
