Function SQLlib_SQLSelect_RunAllTests()
    SQLlib_SQLSelect_RunAllTests = True
    Dim Interfaced As iSQLQuery
    Set MySelect = Create_SQLSelect
    MySelect.Table = "users"
    MySelect.Fields = Array("id", "username")
    MySelect.AddWhere "created", "'2000-01-01'", ">"
    Set Interfaced = MySelect
    Expected = "SELECT id, username FROM users WHERE created>'2000-01-01'"
    SQLlib_SQLSelect_RunAllTests = SQLlib_SQLSelect_RunAllTests And AssertObjectStringEquals(Interfaced, Expected)
    
    MySelect.AddWhere "type", "'admin'"
    SQLlib_SQLSelect_RunAllTests = SQLlib_SQLSelect_RunAllTests And AssertObjectStringEquals(Interfaced, "SELECT id, username FROM users WHERE created>'2000-01-01' AND type='admin'")
    
    MySelect.AddWhere "flag", "NULL", "IS", "OR"
    SQLlib_SQLSelect_RunAllTests = SQLlib_SQLSelect_RunAllTests And AssertObjectStringEquals(Interfaced, "SELECT id, username FROM users WHERE (created>'2000-01-01' AND type='admin') OR flag IS NULL")

    Dim MyOtherSelect As SQLSelect
    Set MyOtherSelect = Create_SQLSelect
    MyOtherSelect.getByProperty "users", "id", "name", ":name"
    MyOtherSelect.AddArgument ":name", "admin"
    Set Interfaced = MyOtherSelect
    SQLlib_SQLSelect_RunAllTests = SQLlib_SQLSelect_RunAllTests And AssertObjectStringEquals(Interfaced, "SELECT id FROM users WHERE name='admin'")
    
    'Check Join
    Set MySelect = Create_SQLSelect
    With MySelect
        .addTable "users", "u"
        .InnerJoin "countries", "c", "u.country=c.country"
        .Fields = Array("u.uname", "c.capital")
    End With
    Set Interfaced = MySelect
    SQLlib_SQLSelect_RunAllTests = SQLlib_SQLSelect_RunAllTests And AssertObjectStringEquals(Interfaced, "SELECT u.uname, c.capital FROM users u INNER JOIN countries c ON u.country=c.country")
    
    MySelect.AddField "t.zone"
    MySelect.InnerJoin "timezones", "t", "c.capital=t.city"
    SQLlib_SQLSelect_RunAllTests = SQLlib_SQLSelect_RunAllTests And AssertObjectStringEquals(Interfaced, "SELECT u.uname, c.capital, t.zone FROM users u INNER JOIN countries c ON u.country=c.country INNER JOIN timezones t ON c.capital=t.city")
    
    'Distinct
    Set MySelect = Create_SQLSelect
    With MySelect
        .addTable "customers", "c"
        .Fields = Array("c.country")
        .Distinct
        .OrderBy ("c.country")
    End With
    Set Interfaced = MySelect
    SQLlib_SQLSelect_RunAllTests = SQLlib_SQLSelect_RunAllTests And AssertObjectStringEquals(Interfaced, "SELECT DISTINCT c.country FROM customers c ORDER BY c.country ASC")
End Function
