VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLQueryGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The is class is an array of differnet queries, which will be executed as one
Implements iSQLQuery

Private aQueries() As String
Private Sub Class_Initialize()
    ReDim aQueries(0) As String
End Sub

Public Sub AddQuery(sQuery As String)
    'Lengthen the array by one and add the new value to the end
    iLength = UBound(aQueries)
    If iLength = 0 And aQueries(0) = "" Then
        aQueries(0) = sQuery
    Else
        ReDim Preserve aQueries(0 To UBound(aQueries) + 1) As String
        aQueries(UBound(aQueries)) = sQuery
    End If
End Sub
Public Function iSQLQuery_ToString() As String
    iSQLQuery_ToString = Join(aQueries, ";")
End Function
