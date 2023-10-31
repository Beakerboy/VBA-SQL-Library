VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iSQLRecordset

'A wrapper for the ADODB.Recordset Object.
'This allows unit testing without an active database implementation
'The native Recordset uses several function names which are reserved keywords

Private orst As ADODB.Recordset

Private Sub Class_Initialize()
    Set orst = New ADODB.Recordset
End Sub

Public Function iSQLRecordset_State()
    iSQLRecordset_State = orst.State
End Function

Public Sub iSQLRecordset_OpenRecordset(MyQuery As String, MyConnection)
    orst.Open MyQuery, MyConnection.Connection
End Sub

Public Sub iSQLRecordset_CloseRecordset()
    orst.Close
End Sub

Public Function iSQLRecordset_GetValue(MyFieldname)
    iSQLRecordset_GetValue = orst(MyFieldname)
End Function

Public Function iSQLRecordset_GetRows(num As Integer)
    iSQLRecordset_GetRows = orst.GetRows(num)
End Function

Property Get iSQLRecordset_EOF()
    iSQLRecordset_EOF = orst.EOF
End Property
