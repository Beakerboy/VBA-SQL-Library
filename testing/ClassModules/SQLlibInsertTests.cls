VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLlibInsertTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements iTestCase

Dim Interfaced As iSQLQuery
Dim MyInsert As SQLInsert

Sub iTestCase_Setup()
    Set MyInsert = Create_SQLInsert
End Sub

Sub iTestCase_TearDown()

End Sub

Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array("SimpleInsertTest", "InsertSelectTest", "InsertMultipleTest")
End Function

Function iTestCase_GetObject()
    Set iTestCase_GetObject = New SQLlibInsertTests
End Function

Function SimpleInsertTest()
    With MyInsert
        .Table = "users"
        .Fields = Array("name", "type")
        .Values = Array("'foo'", "'admin'")
        .Returning = "id"
    End With
    Set Interfaced = MyInsert
    Expected = "INSERT INTO users (name, type) VALUES ('foo', 'admin') RETURNING id"
    SimpleInsertTest = AssertEquals(Interfaced.toString, Expected)
End Function

Function InsertSelectTest()
    Dim MySelect As SQLSelect
    Set MySelect = Create_SQLSelect
    With MySelect
        .Table = "account_types"
        .Fields = Array("'foo'", "id")
        .AddWhere "type", ":type"
        .AddArgument ":type", "admin"
    End With
    With MyInsert
        .Table = "users"
        .Fields = Array("name", "type_id")
        .Values = Array()
        .Returning = "id"
        Set .From = MySelect
    End With
    Set Interfaced = MyInsert
    Expected = "INSERT INTO users (name, type_id) (SELECT 'foo', id FROM account_types WHERE type='admin') RETURNING id"
    InsertSelectTest = AssertEquals(Interfaced.toString, Expected)
End Function

Public Function InsertMultipleTest()
    MyInsert.Table = "users"
    MyInsert.Fields = Array("name", "type")
    Dim Values(1) As Variant
    
    Values(0) = Array("'foo'", "'admin'")
    Values(1) = Array("'bar'", "'editor'")
    MyInsert.Values = Values
    Set Interfaced = MyInsert
    Expected = "INSERT INTO users (name, type) VALUES ('foo', 'admin'), ('bar', 'editor')"
    InsertMultipleTest = AssertObjectStringEquals(Interfaced, Expected)
End Function
