Attribute VB_Name = "SQLUnitTests"
Public Function SQLlib_RunAllTests()
    
    RunAllModuleTests ("SQLlib_SQLDatabase")
    
    RunAllModuleTests ("SQLlib_SQLInsert")

    RunAllModuleTests ("SQLlib_SQLSelect")
    
    RunAllModuleTests ("SQLlib_SQLStatic")

    Dim Interfaced As iSQLQuery
    '******************************Check Delete********************************
    Dim MyDelete As SQLDelete
    Set MyDelete = Create_SQLDelete()
    MyDelete.Table = "users"
    
    Set Interfaced = MyDelete
    AssertObjectStringEquals Interfaced, "DELETE FROM users"
    
    MyDelete.AddWhere "age", ":age", "<"
    MyDelete.AddArgument ":age", 13
    AssertObjectStringEquals Interfaced, "DELETE FROM users WHERE age<13"

    '******************************Check Update********************************
    Dim MyUpdate As SQLUpdate
    Set MyUpdate = Create_SQLUpdate
    With MyUpdate
        .Table = "users"
        .Fields = Array("username")
        .Values = Array(str("admin' WHERE id=1;DROP TABLE users;"))
        .AddWhere "id", 1
    End With
    Set Interfaced = MyUpdate
    AssertObjectStringEquals Interfaced, "UPDATE users SET username='admin'' WHERE id=1;DROP TABLE users;' WHERE id=1"

    '*******************Check SubSelect****************************************
    'Dim MySubselect As New SQLSubselect
    'Set MySubselect.SelectSQL = MyOtherSelect
    'MySubselect.SelectAs = "user_id"
    'CheckSQLValue MySubselect, "(SELECT id FROM users WHERE name='admin') AS user_id"
            
    '*****************Check Create*****************
    'Dim MyCreate As SQLCreate
    'Set MyCreate = Create_SQLCreate
    'With MyCreate
    '    .Table = "users"
    '    .Fields = Array(Array("id", "int"), Array("username", "varchar", 50))
    'End With
    'Dim Interfaced As iSQLQuery
    'Set Interfaced = MyCreate
    'CheckSQLValue Interfaced, "CREATE TABLE users (id int, username varchar(50))"
    
    '****************Check Where Group*****************************************
    'Dim MyWhereGroup As New SQLWhereGroup
    'Dim MyOtherWhere As New SQLCondition
    'MyOtherWhere.Create "type", "'toys'"
    'MyWhereGroup.SetGroup MyWhere, MyOtherWhere, "AND"
    'CheckSQLValue MyWhereGroup, "id=2 AND type='toys'"
    
    'Dim MyThirdWhere As New SQLCondition
    'MyThirdWhere.Create "color", "'pink'"
    
    'MyWhereGroup.AddWhere MyThirdWhere, "OR"
    'CheckSQLValue MyWhereGroup, "(id=2 AND type='toys') OR color='pink'"
    
    'Dim MyOtherWhereGroup As New SQLWhereGroup
    'MyOtherWhereGroup.SetGroup MyWhere, MyThirdWhere, "OR"
    'MyWhereGroup.AddWhere MyOtherWhereGroup, "AND"
    'CheckSQLValue MyWhereGroup, "((id=2 AND type='toys') OR color='pink') AND (id=2 OR color='pink')"
    SQLlib_RunAllTests = True
End Function
