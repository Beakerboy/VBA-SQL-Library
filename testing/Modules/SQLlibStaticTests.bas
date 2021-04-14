function SQLlib_SQLStatic_RunAllTests()
    Dim Interfaced As iSQLQuery
    Dim MyStatic As SQLStaticQuery
    Set MyStatic = Create_SQLStaticQuery
    MyStatic.Query = "DELETE FROM users"
    Set Interfaced = MyStatic
    AssertObjectStringEquals Interfaced, "DELETE FROM users"
    
    'Name missing ":"
    MyStatic.Query = "SELECT name FROM users WHERE id=:id"
    MyStatic.AddArgument "id", 4
    AssertObjectStringEquals Interfaced, "SELECT name FROM users WHERE id=:id"
    
    'Proper function
    MyStatic.AddArgument ":id", 4
    AssertObjectStringEquals Interfaced, "SELECT name FROM users WHERE id=4"
    
    'Can Change value
    MyStatic.AddArgument ":id", 40
    AssertObjectStringEquals Interfaced, "SELECT name FROM users WHERE id=40"
    
    'Text is escaped
    MyStatic.AddArgument ":id", "text"
    AssertObjectStringEquals Interfaced, "SELECT name FROM users WHERE id='text'"
    
    'Multiple arguments
    MyStatic.Query = "SELECT name FROM users WHERE id=:id AND type=:type"
    MyStatic.ClearArguments
    MyStatic.AddArgument ":type", "admin"
    AssertObjectStringEquals Interfaced, "SELECT name FROM users WHERE id=:id AND type='admin'"

    MyStatic.ClearArguments
    MyStatic.AddArgument ":id", 4
    AssertObjectStringEquals Interfaced, "SELECT name FROM users WHERE id=4 AND type=:type"
    
    MyStatic.AddArgument ":type", "admin"
    AssertObjectStringEquals Interfaced, "SELECT name FROM users WHERE id=4 AND type='admin'"
    
    'Can not place an argument in a value
    MyStatic.AddArgument ":id", "4:type"
    MyStatic.AddArgument ":type", ";DELETE FROM users;:id"
    AssertObjectStringEquals Interfaced, "SELECT name FROM users WHERE id='4:type' AND type=';DELETE FROM users;:id'"
End Function
