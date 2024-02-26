VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iSQLRecordset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Interface: iSQLRecordset
' Defines a wrapper for custom recordsets or ADODB.Recordset

' Sub: OpenRecordset
Public Sub OpenRecordset(MyQuery As String, Connection)

End Sub

' Sub: CloseRecordset
Public Sub CloseRecordset()

End Sub

' Function: GetValue
Public Function GetValue(MyFieldname)

End Function

' Function: GetValue
Public Function GetRows(num As Integer)

End Function

' Function: State
Public Function State()

End Function

' Property: EOF
Property Get EOF()

End Property
