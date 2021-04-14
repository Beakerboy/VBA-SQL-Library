Function SQLlib_SQLDelete_RunAllTests()
    Result = True
    Dim MyDelete As SQLDelete
    Dim Interfaced As iSQLQuery
    Set MyDelete = Create_SQLDelete()
    MyDelete.Table = "users"
    
    Set Interfaced = MyDelete
    Result = Result And AssertObjectStringEquals(Interfaced, "DELETE FROM users")
    
    MyDelete.AddWhere "age", ":age", "<"
    MyDelete.AddArgument ":age", 13
    Result = Result And AssertObjectStringEquals(Interfaced, "DELETE FROM users WHERE age<13")

    SQLlib_SQLDelete_RunAllTests = Result
End Function
