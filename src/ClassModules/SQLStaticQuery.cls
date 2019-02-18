VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLStaticQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' Class: SQLStatic
' A class to reuse an SQL query with multiple values.
Option Explicit
Implements iSQLQuery

Private sQuery As String
Private oSQL As SQLQuery

' Constructor: Class_Initialize
' Initializes the object.
Private Sub Class_Initialize()
    Set oSQL = New SQLQuery
End Sub

' Property: Query
' Sets the SQL Query.
Public Property Let Query(sValue)
    sQuery = sValue
End Property

' Function: ToString
' Implements the interfaced ToString function
Public Function iSQLQuery_ToString() As String
    iSQLQuery_ToString = oSQL.ReplaceArguments(sQuery)
End Function

' Function: AddArgument
' Adds a value to the SQL query
Public Sub AddArgument(sName As String, vValue)
    oSQL.AddArgument sName, vValue
End Sub

' Sub: ClearArguments
' Clears the values of any query arguments.
Public Sub ClearArguments()
    oSQL.ClearArguments
End Sub
