VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLlibSelectTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iTestCase

Dim Interfaced As iSQLQuery
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

Function SimpleSelectTest()

    Set MySelect = Create_SQLSelect
    MySelect.Table = "users"
    MySelect.Fields = Array("id", "username")
    MySelect.AddWhere "created", "'2000-01-01'", ">"
    Set Interfaced = MySelect
    Expected = "SELECT id, username FROM users WHERE created>'2000-01-01'"
    Call AssertObjectStringEquals(Interfaced, Expected)

    MySelect.AddWhere "type", "'admin'"
    Call AssertObjectStringEquals(Interfaced, "SELECT id, username FROM users WHERE created>'2000-01-01' AND type='admin'")
    
    MySelect.AddWhere "flag", "NULL", "IS", "OR"
    Call AssertObjectStringEquals(Interfaced, "SELECT id, username FROM users WHERE (created>'2000-01-01' AND type='admin') OR flag IS NULL")
End Function

Function GetByPropertyTest()
    Dim MyOtherSelect As SQLSelect
    Set MyOtherSelect = Create_SQLSelect
    MyOtherSelect.getByProperty "users", "id", "name", ":name"
    MyOtherSelect.AddArgument ":name", "admin"
    Set Interfaced = MyOtherSelect
    GetByPropertyTest = AssertObjectStringEquals(Interfaced, "SELECT id FROM users WHERE name='admin'")
End Function

Function JoinTest()
    'Check Join
    Set MySelect = Create_SQLSelect
    With MySelect
        .addTable "users", "u"
        .InnerJoin "countries", "c", "u.country=c.country"
        .Fields = Array("u.uname", "c.capital")
    End With
    Set Interfaced = MySelect
    JoinTest = AssertObjectStringEquals(Interfaced, "SELECT u.uname, c.capital FROM users u INNER JOIN countries c ON u.country=c.country")
    
    MySelect.AddField "t.zone"
    MySelect.InnerJoin "timezones", "t", "c.capital=t.city"
    Call AssertObjectStringEquals(Interfaced, "SELECT u.uname, c.capital, t.zone FROM users u INNER JOIN countries c ON u.country=c.country INNER JOIN timezones t ON c.capital=t.city")
End Function

Function DistinctTest()
    'Distinct
    Set MySelect = Create_SQLSelect
    With MySelect
        .addTable "customers", "c"
        .Fields = Array("c.country")
        .Distinct
        .OrderBy ("c.country")
    End With
    Set Interfaced = MySelect
    DistinctTest = AssertObjectStringEquals(Interfaced, "SELECT DISTINCT c.country FROM customers c ORDER BY c.country ASC")
End Function

Function MultipleTableTest()
   'Distinct
    Set MySelect = Create_SQLSelect
    With MySelect
        .AddTable "countries", "c"
        .AddTable "users", "u"
        .Fields = Array("u.uname", "c.capital")
    End With
    Set Interfaced = MySelect
    MultipleTableTest = AssertObjectStringEquals(Interfaced, "SELECT u.uname, c.capital FROM countries c, users u")
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
