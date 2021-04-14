Function SQLlib_SQLUpdate_RunAllTests()
    Result = True
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
    Result = Result And AssertObjectStringEquals(Interfaced, "UPDATE users SET username='admin'' WHERE id=1;DROP TABLE users;' WHERE id=1")
    SQLlib_SQLUpdate_RunAllTests = Result
End Function
