VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLlibUpdateTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iTestCase
Sub iTestCase_Setup()

End Sub
Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array("SimpleUpdateTest")
End Function

Function iTestCase_GetObject()
    Set iTestCase_GetObject = New SQLlibUpdateTests
End Function

Function SimpleUpdateTest()
    Dim MyUpdate As SQLUpdate
    Dim Interfaced As iSQLQuery
    Set MyUpdate = Create_SQLUpdate
    With MyUpdate
        .Table = "users"
        .Fields = Array("username")
        .Values = Array(str("admin' WHERE id=1;DROP TABLE users;"))
        .AddWhere "id", 1
    End With
    Set Interfaced = MyUpdate
    SimpleUpdateTest = AssertObjectStringEquals(Interfaced, "UPDATE users SET username='admin'' WHERE id=1;DROP TABLE users;' WHERE id=1")
End Function
