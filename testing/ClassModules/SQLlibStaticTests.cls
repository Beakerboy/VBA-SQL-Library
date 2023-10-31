VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLlibStaticTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iTestCase

Dim Interfaced As iSQLQuery
Dim MyStatic As SQLStaticQuery
Dim TestCase As iTestCase

' Constructor: Class_Initialize
' Initialize class members
Private Sub Class_Initialize()
    Set TestCase = CreateTestCase()
End Sub

Sub iTestCase_Setup()
    Set MyStatic = Create_SQLStaticQuery
    Set Interfaced = MyStatic
End Sub

Sub iTestCase_TearDown()

End Sub

Function NoArgumentTest()
    MyStatic.Query = "DELETE FROM users"
    NoArgumentTest = AssertObjectStringEquals(Interfaced, "DELETE FROM users")
End Function

Function AddargumentMissingColonTest()
    MyStatic.Query = "SELECT name FROM users WHERE id=:id"
    MyStatic.AddArgument "id", 4
    AddargumentMissingColonTest = AssertObjectStringEquals(Interfaced, "SELECT name FROM users WHERE id=:id")
End Function

Function CorrectUseTest()
    MyStatic.Query = "SELECT name FROM users WHERE id=:id"
    MyStatic.AddArgument ":id", 4
    CorrectUseTest = AssertObjectStringEquals(Interfaced, "SELECT name FROM users WHERE id=4")
End Function

Function ChangeArgumentTest()
    MyStatic.Query = "SELECT name FROM users WHERE id=:id"
    MyStatic.AddArgument ":id", 4
    MyStatic.AddArgument ":id", 40
    ChangeArgumentTest = AssertObjectStringEquals(Interfaced, "SELECT name FROM users WHERE id=40")
End Function

Function EscapeTextTest()
    MyStatic.Query = "SELECT name FROM users WHERE id=:id"
    MyStatic.AddArgument ":id", "text"
    EscapeTextTest = AssertObjectStringEquals(Interfaced, "SELECT name FROM users WHERE id='text'")
End Function

Function MultipleArgumentTest()
    MyStatic.Query = "SELECT name FROM users WHERE id=:id AND type=:type"
    MyStatic.AddArgument ":type", "admin"
    MultipleArgumentTest = AssertObjectStringEquals(Interfaced, "SELECT name FROM users WHERE id=:id AND type='admin'")
End Function

Function ClearArgumentsTest()
    MyStatic.Query = "SELECT name FROM users WHERE id=:id AND type=:type"
    MyStatic.AddArgument ":type", "admin"
    MyStatic.ClearArguments
    MyStatic.AddArgument ":id", 4
    ClearArgumentsTest = AssertObjectStringEquals(Interfaced, "SELECT name FROM users WHERE id=4 AND type=:type")
    MyStatic.AddArgument ":type", "admin"
    Call AssertObjectStringEquals(Interfaced, "SELECT name FROM users WHERE id=4 AND type='admin'")
End Function

' Function: InsertArgumentInValueTest
' Ensure that argument replacement will not search within string literals and replace the substring
Function InsertArgumentInValueTest()
    MyStatic.Query = "SELECT name FROM users WHERE id=:id AND type=:type"
    MyStatic.AddArgument ":id", "4:type"
    MyStatic.AddArgument ":type", ";DELETE FROM users;:id"
    InsertArgumentInValueTest = AssertObjectStringEquals(Interfaced, "SELECT name FROM users WHERE id='4:type' AND type=';DELETE FROM users;:id'")
End Function

' Function: InsertArgumentInValueTest
' Ensure that argument replacement will not search within string literals and replace the substring
Function InsertArgumentInStringTest()
    MyStatic.Query = "SELECT name FROM users WHERE id=:id AND type=':id'"
    MyStatic.AddArgument ":id", "4"
    InsertArgumentInStringTest = AssertObjectStringEquals(Interfaced, "SELECT name FROM users WHERE id='4' AND type=':id'")
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
