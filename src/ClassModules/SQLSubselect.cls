VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLSubselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private MySelect As SQLSelect
Private sAs As String

Private Sub Class_Initialize()
    Set MySelect = New SQLSelect
End Sub

Property Set SelectSQL(oValue As SQLSelect)
    Set MySelect = oValue
End Property

Property Let SelectAs(sValue As String)
    sAs = sValue
End Property

Property Get SelectSQL() As SQLSelect
    Set SelectSQL = MySelect
End Property

Property Get SelectAs() As String
    SelectAs = aAs
End Property

Public Function toString() As String
    Dim MySQLQuery As iSQLQuery
    Set MySQLQuery = MySelect
    Dim return_string As String
    return_string = "(" & MySQLQuery.toString & ")"
    If sAs <> "" Then
        return_string = return_string & " AS " & sAs
    End If
    toString = return_string
End Function
