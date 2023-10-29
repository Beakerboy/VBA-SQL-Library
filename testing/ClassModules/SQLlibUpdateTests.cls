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

Dim TestCase As iTestCase

' Constructor: Class_Initialize
' Initialize class members
Private Sub Class_Initialize()
    Set TestCase = CreateTestCase()
End Sub

Sub iTestCase_Setup()

End Sub

Sub iTestCase_TearDown()

End Sub

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

' Sub: iTestCase_RunTest
' Run a specific test.
Public Sub iTestCase_RunTest(Test As String, Optional clsObj = Nothing)
    TestCase.RunTest Test, Me
End Sub

' Sub: iTestCase_RunAllTests
' Run all tests.
Public Sub iTestCase_RunAllTests(Optional ByVal clsObj = Nothing)
    TestCase.RunAllTests Me
End Sub
