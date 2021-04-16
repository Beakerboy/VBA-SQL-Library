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
Sub iTestCase_Setup()

End Sub
Function iTestCase_GetAllTests()
    iTestCase_GetAllTests = Array("SimpleDeleteTest")
End Function

Function iTestCase_GetObject()
    Set iTestCase_GetObject = New SQLlibDeleteTests
End Function

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

