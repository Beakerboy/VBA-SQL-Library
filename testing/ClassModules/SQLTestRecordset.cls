VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLTestRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iSQLRecordset

Private sQuery As String

Public Function iSQLRecordset_State()

End Function

Public Sub iSQLRecordset_OpenRecordset(MyQuery As String, Connection)
    sQuery = MyQuery
End Sub

Public Sub iSQLRecordset_CloseRecordset()

End Sub

Public Function iSQLRecordset_GetValue(MyFieldname)
    iSQLRecordset_GetValue = sQuery
End Function

Public Function iSQLRecordset_GetRows(num As Integer)
    iSQLRecordset_GetRows = num
End Function

Property Get iSQLRecordset_EOF()
    iSQLRecordset_EOF = False
End Property

