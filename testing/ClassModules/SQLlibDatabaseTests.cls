VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLlibDatabaseTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iTestCase

Dim MyDatabase As SQLDatabase
Dim MyRecordset As New SQLTestRecordset
Dim MyConnection As New SQLTestConnection
Dim SimpleInsert As SQLInsert
Dim TestCase As iTestCase

' Constructor: Class_Initialize
' Initialize class members
Private Sub Class_Initialize()
    Set TestCase = CreateTestCase()
End Sub

Sub iTestCase_Setup()
    Set MyDatabase = Create_SQLDatabase()
    
    With MyDatabase
        .DSN = "mydsn"
        .Password = "Pa$$word"
        .Username = "myusername"
        Set .Recordset = MyRecordset
        Set .Connection = MyConnection
    End With
    

    Set SimpleInsert = Create_SQLInsert
    With SimpleInsert
        .Table = "users"
        .Fields = Array("id")
        .Values = Array(1)
    End With
End Sub

Sub iTestCase_TearDown()

End Sub

Function mssqlTest()
    MyDatabase.DBType = "mssql"
        
    Actual = MyDatabase.InsertGetNewId(SimpleInsert)
    Expected = "SET NOCOUNT ON;INSERT INTO users (id) VALUES (1);SELECT SCOPE_IDENTITY() as somethingunique"
    mssqlTest = AssertEquals(Actual, Expected)
End Function

Function psqlTest()
    MyDatabase.DBType = "psql"

    Actual = MyDatabase.InsertGetNewId(SimpleInsert, "id")
    Expected = "INSERT INTO users (id) VALUES (1) RETURNING id"
    psqlTest = AssertEquals(Actual, Expected)

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
