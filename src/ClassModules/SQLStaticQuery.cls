VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLStaticQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Implements iSQLQuery

Private sQuery As String
Private oSQL As SQLQuery

Private Sub Class_Initialize()
    Set oSQL = New SQLQuery
End Sub

Public Property Let Query(sValue)
    sQuery = sValue
End Property

Public Function iSQLQuery_ToString() As String
    Dim ReturnString As String
    ReturnString = sQuery
    Dim sKey As Variant
    For Each sKey In oSQL.Arguments.Keys
        sValue = oSQL.Arguments(sKey)
        ReturnString = Replace(sQuery, sKey, sValue)
    Next sKey
    iSQLQuery_ToString = ReturnString
End Function

Public Sub AddArgument(sName As String, vValue)
    oSQL.AddArgument sName, vValue
End Sub



