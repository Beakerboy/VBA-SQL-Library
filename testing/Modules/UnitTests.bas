Attribute VB_Name = "UnitTests"
Public Function RunUnitTests()
    '*****************Check Where*****************
    Dim MyWhere As New SQLWhere
    MyWhere.Create "id", "=", 2
    CheckValue MyWhere, "id=2"
    
    '****************Check Where Group************
    Dim MyWhereGroup As New SQLWhereGroup
    Dim MyOtherWhere As New SQLWhere
    MyOtherWhere.Create "type", "=", "'toys'"
    MyWhereGroup.SetGroup MyWhere, MyOtherWhere, "AND"
    CheckValue MyWhereGroup, "id=2 AND type='toys'"
    
    Dim MyThirdWhere As New SQLWhere
    MyThirdWhere.Create "color", "=", "'pink'"
    
    MyWhereGroup.AddWhere MyThirdWhere, "OR"
    CheckValue MyWhereGroup, "(id=2 AND type='toys') OR color='pink'"
    
    Dim MyOtherWhereGroup As New SQLWhereGroup
    MyOtherWhereGroup.SetGroup MyWhere, MyThirdWhere, "OR"
    MyWhereGroup.AddWhere MyOtherWhereGroup, "AND"
    CheckValue MyWhereGroup, "((id=2 AND type='toys') OR color='pink') AND (id=2 OR color='pink')"
    
    '*******************Check Select********************
    Dim MySelect As New SQLSelect
    MySelect.Table = "users"
    MySelect.Fields = Array("id", "username")
    MySelect.AddWhere "created", ">", "'2000-01-01'"
    Dim Interfaced As iSQLQuery
    Set Interfaced = MySelect
    CheckValue Interfaced, "SELECT id, username FROM users WHERE created>'2000-01-01'"
    
    MySelect.AddWhere "type", "=", "'admin'"
    CheckValue Interfaced, "SELECT id, username FROM users WHERE created>'2000-01-01' AND type='admin'"
    
    MySelect.AddWhere "flag", "IS", "NULL", "OR"
    CheckValue Interfaced, "SELECT id, username FROM users WHERE (created>'2000-01-01' AND type='admin') OR flag IS NULL"

    Dim MyOtherSelect As New SQLSelect
    MyOtherSelect.getByProperty "users", "id", "name", "'admin'"
    Set Interfaced = MyOtherSelect
    CheckValue Interfaced, "SELECT id FROM users WHERE name='admin'"
    
    '******************************Check Join*********************************
    Set MySelect = New SQLSelect
    With MySelect
        .addTable "users", "u"
        .InnerJoin "countries", "c", "u.country=c.country"
        .Fields = Array("u.uname", "c.capital")
    End With
    Set Interfaced = MySelect
    CheckValue Interfaced, "SELECT u.uname, c.capital FROM users u INNER JOIN countries c ON u.country=c.country"
    '*******************Check SubSelect********************
    Dim MySubselect As New SQLSubselect
    Set MySubselect.SelectSQL = MyOtherSelect
    MySubselect.SelectAs = "user_id"
    CheckValue MySubselect, "(SELECT id FROM users WHERE name='admin') AS user_id"
    
    '*******************Check Distinct*****************************************
    Set MySelect = New SQLSelect
    With MySelect
        .addTable "customers", "c"
        .Fields = Array("c.country")
        .Distinct
    End With
    Set Interfaced = MySelect
    CheckValue Interfaced, "SELECT DISTINCT c.country FROM customers c"
    
    '*********************Check Insert***********************
    Dim MyInsert As New SQLInsert
    MyInsert.Table = "users"
    MyInsert.Fields = Array("name", "type")
    MyInsert.Values = Array("'foo'", "'admin'")
    MyInsert.Returning = "id"
    Set Interfaced = MyInsert
    CheckValue Interfaced, "INSERT INTO users (name, type) VALUES ('foo', 'admin') RETURNING id"
    
    MyInsert.Fields = Array("name", "type_id")
    MyInsert.Values = Array()
    
    Set MySelect = New SQLSelect
    MySelect.Table = "account_types"
    MySelect.Fields = Array("'foo'", "id")
    MySelect.AddWhere "type", "=", "'admin'"
    Set MyInsert.setSelect = MySelect
    CheckValue Interfaced, "INSERT INTO users (name, type_id) (SELECT 'foo', id FROM account_types WHERE type='admin') RETURNING id"
    
    '*********************Check Insert Multiple Values***********************
    Set MyInsert = New SQLInsert
    MyInsert.Table = "users"
    MyInsert.Fields = Array("name", "type")
    Dim Values2D(1, 1) As Variant
    Values2D(0, 0) = "'foo'"
    Values2D(0, 1) = "'admin'"
    Values2D(1, 0) = "'bar'"
    Values2D(1, 1) = "'editor'"
    MyInsert.Values = Values2D
    Set Interfaced = MyInsert
    CheckValue Interfaced, "INSERT INTO users (name, type) VALUES ('foo', 'admin'), ('bar', 'editor')"
    
    '******************************Check Update********************************
    Dim MyUpdate As New SQLUpdate
    With MyUpdate
        .Table = "users"
        .Fields = Array("username")
        .Values = Array(str("admin' WHERE id=1;DROP TABLE users;"))
        .AddWhere "id", "=", 1
    End With
    Set Interfaced = MyUpdate
    CheckValue Interfaced, "UPDATE users SET username='admin'' WHERE id=1;DROP TABLE users;' WHERE id=1"
End Function

Function CheckValue(MyObject, ExpectedValue)
    If MyObject.toString <> ExpectedValue Then
        MsgBox "Expected: " & ExpectedValue & vbNewLine & "Provided: " & MyObject.toString
    End If
End Function


