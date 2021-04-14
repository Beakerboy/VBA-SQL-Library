Attribute VB_Name = "SQLUnitTests"
Public Function SQLlib_RunAllTests()
    Result = True
    
    Result = Result And RunAllModuleTests("SQLlib_SQLDatabase")
    
    Result = Result And RunAllModuleTests("SQLlib_SQLInsert")
    
    Result = Result And RunAllModuleTests("SQLlib_SQLSelect")
    
    Result = Result And RunAllModuleTests("SQLlib_SQLStatic")

    Result = Result And RunAllModuleTests("SQLlib_SQLDelete")

    Result = Result And RunAllModuleTests("SQLlib_SQLUpdate")

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

    '*******************Check SubSelect****************************************
    'Dim MySubselect As New SQLSubselect
    'Set MySubselect.SelectSQL = MyOtherSelect
    'MySubselect.SelectAs = "user_id"
    'CheckSQLValue MySubselect, "(SELECT id FROM users WHERE name='admin') AS user_id"

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

    SQLlib_RunAllTests = Result
End Function
