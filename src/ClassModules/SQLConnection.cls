VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iSQLConnection

'A wrapper for the ADODB.Connection Object.
'This allows unit testing without an active database implementation
'The native Connection uses several function names which are reserved keywords

Private ocnt As ADODB.Connection

Private Sub Class_Initialize()
    Set ocnt = New ADODB.Connection
End Sub

Public Property Get iSQLConnection_Connection()
    Set iSQLConnection_Connection = ocnt
End Property

Public Sub iSQLConnection_OpenConnection()
    ocnt.Open
End Sub

Public Sub iSQLConnection_CloseConnection()
    ocnt.Close
End Sub

Public Function iSQLConnection_State()
    iSQLConnection_State = ocnt.State
End Function

Public Property Let iSQLConnection_ConnectionString(sConnection As String)
    ocnt.ConnectionString = sConnection
End Property
