VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLTestRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private orst As ADODB.Recordset
Private sQueryType As String

Private Sub Class_Initialize()
    Set orst = New ADODB.Recordset
End Sub

Public Property Get State()
    State = orst.State
End Property

Public Sub OpenRecordset(MyQuery As String, Connection)
    GetQueryType MyQuery
    MsgBox MyQuery
    If sQueryType = "Select" Then
        orst.Open MyQuery, Connection
    End If
End Sub

Public Sub CloseRecordset()
    orst.Close
End Sub
Public Function GetValue(MyFieldname)
    Dim iResult As Variant
    If sQueryType = "Select" Then
        iResult = orst(MyFieldname)
    Else
        iResult = Int(101 * Rnd + 1)       'Random Number from 1 to 100
    End If
    MsgBox MyFieldname & "=" & iResult
    GetValue = iResult
End Function

Private Sub GetQueryType(MyQuery)
    Dim FirstChars As String
    FirstChars = Left(MyQuery, 3)
    If FirstChars = "SEL" Then
        sQueryType = "Select"
    ElseIf FirstChars = "INS" Or FirstChar = "SET" Then
        sQueryType = "Insert"
    End If
End Sub
