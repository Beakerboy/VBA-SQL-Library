VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLlibDeleteTests"
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

Function SimpleDeleteTest()
    Dim MyDelete As SQLDelete
    Dim Interfaced As iSQLQuery
    Set MyDelete = Create_SQLDelete()
    MyDelete.Table = "users"
    
    Set Interfaced = MyDelete
    Result = Result And AssertObjectStringEquals(Interfaced, "DELETE FROM users")
    
    MyDelete.AddWhere "age", ":age", "<"
    MyDelete.AddArgument ":age", 13

    SimpleDeleteTest = AssertObjectStringEquals(Interfaced, "DELETE FROM users WHERE age<13")
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
