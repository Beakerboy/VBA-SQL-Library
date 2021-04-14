Function SQLlib_SQLInsert_RunAllTests()
    Dim Interfaced As iSQLQuery
    Dim MyInsert As SQLInsert
    Set MyInsert = Create_SQLInsert
    MyInsert.Table = "users"
    MyInsert.Fields = Array("name", "type")
    MyInsert.Values = Array("'foo'", "'admin'")
    MyInsert.Returning = "id"
    Set Interfaced = MyInsert
    Expected = "INSERT INTO users (name, type) VALUES ('foo', 'admin') RETURNING id"
    AssertEquals Interfaced.toString, Expected
    
    Dim MySelect As SQLSelect
    Set MySelect = Create_SQLSelect
    With MySelect
        .Table = "account_types"
        .Fields = Array("'foo'", "id")
        .AddWhere "type", ":type"
        .AddArgument ":type", "admin"
    End With
    With MyInsert
        .Fields = Array("name", "type_id")
        .Values = Array()
        Set .From = MySelect
    End With
    AssertEquals Interfaced.toString, "INSERT INTO users (name, type_id) (SELECT 'foo', id FROM account_types WHERE type='admin') RETURNING id"
    
    'Insert Multiple Values
    Set MyInsert = Create_SQLInsert
    MyInsert.Table = "users"
    MyInsert.Fields = Array("name", "type")
    Dim Values(1) As Variant
    
    Values(0) = Array("'foo'", "'admin'")
    Values(1) = Array("'bar'", "'editor'")
    MyInsert.Values = Values
    Set Interfaced = MyInsert
    AssertObjectStringEquals Interfaced, "INSERT INTO users (name, type) VALUES ('foo', 'admin'), ('bar', 'editor')"
End Function
