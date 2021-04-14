Function SQLlib_SQLDatabase_RunAllTests()
    Result = True
    Dim MyDatabase As SQLDatabase
    Set MyDatabase = Create_SQLDatabase()
    Dim MyRecordset As New SQLTestRecordset
    Dim MyConnection As New SQLTestConnection
    With MyDatabase
        .DSN = "mydsn"
        .DBType = "mssql"
        .Password = "Pa$$word"
        .Username = "myusername"
        Set .Recordset = MyRecordset
        Set .Connection = MyConnection
    End With
    
    Dim SimpleInsert As SQLInsert
    Set SimpleInsert = Create_SQLInsert
    With SimpleInsert
        .Table = "users"
        .Fields = Array("id")
        .Values = Array(1)
    End With
    Actual = MyDatabase.InsertGetNewId(SimpleInsert)
    Expected = "SET NOCOUNT ON;INSERT INTO users (id) VALUES (1);SELECT SCOPE_IDENTITY() as somethingunique"
    Result = Result And AssertEquals(Actual, Expected)
    
    MyDatabase.DBType = "psql"
    Result = Result And AssertEquals(MyDatabase.InsertGetNewId(SimpleInsert, "id"), "INSERT INTO users (id) VALUES (1) RETURNING id")
    SQLlib_SQLDatabase_RunAllTests = Result
End Function

