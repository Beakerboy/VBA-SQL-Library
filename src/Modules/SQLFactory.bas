Attribute VB_Name = "SQLFactory"
Option Explicit
Public Function Create_SQLDatabase() As SQLDatabase
    Set Create_SQLDatabase = New SQLDatabase
End Function

Public Function Create_SQLInsert() As SQLInsert
    Set Create_SQLInsert = New SQLInsert
End Function

Public Function Create_SQLSelect() As SQLSelect
    Set Create_SQLSelect = New SQLSelect
End Function

Public Function Create_SQLDelete() As SQLDelete
    Set Create_SQLDelete = New SQLDelete
End Function

Public Function Create_SQLCreate() As SQLCreate
    Set Create_SQLCreate = New SQLCreate
End Function

Public Function Create_SQLUpdate() As SQLUpdate
    Set Create_SQLUpdate = New SQLUpdate
End Function

Public Function Create_SQLStaticQuery() As SQLStaticQuery
    Set Create_SQLStaticQuery = New SQLStaticQuery
End Function
