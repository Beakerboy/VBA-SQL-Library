Attribute VB_Name = "SQLFactory"
Option Explicit

' Function: Create_SQLDatabase
' Factory method to create a new SQLDatabase object
Public Function Create_SQLDatabase() As SQLDatabase
    Set Create_SQLDatabase = New SQLDatabase
End Function

' Function: Create_SQLInsert
' Factory method to create a new SQLInsert object
Public Function Create_SQLInsert() As SQLInsert
    Set Create_SQLInsert = New SQLInsert
End Function

' Function: Create_SQLSelect
' Factory method to create a new SQLSelect object
Public Function Create_SQLSelect() As SQLSelect
    Set Create_SQLSelect = New SQLSelect
End Function

' Function: Create_SQLDelete
' Factory method to create a new SQLDelete object
Public Function Create_SQLDelete() As SQLDelete
    Set Create_SQLDelete = New SQLDelete
End Function

' Function: Create_SQLCreate
' Factory method to create a new SQLCreate object
Public Function Create_SQLCreate() As SQLCreate
    Set Create_SQLCreate = New SQLCreate
End Function

' Function: Create_SQLUpdat
' Factory method to create a new SQLUpdate object
Public Function Create_SQLUpdate() As SQLUpdate
    Set Create_SQLUpdate = New SQLUpdate
End Function

' Function: Create_SQLStaticQuery
' Factory method to create a new SQLStaticQuery object
Public Function Create_SQLStaticQuery() As SQLStaticQuery
    Set Create_SQLStaticQuery = New SQLStaticQuery
End Function
