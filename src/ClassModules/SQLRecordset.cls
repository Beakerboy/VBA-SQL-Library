VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private orst As ADODB.Recordset

Private Sub Class_Initialize()
    Set orst = New ADODB.Recordset
End Sub

Public Property Get State()
    State = orst.State
End Property

Public Sub OpenRecordset(MyQuery As String, Connection)
    orst.Open MyQuery, Connection
End Sub

Public Sub CloseRecordset()
    orst.Close
End Sub

Public Function GetValue(MyFieldname As String)
    GetValue = orst(MyFieldname)
End Function
