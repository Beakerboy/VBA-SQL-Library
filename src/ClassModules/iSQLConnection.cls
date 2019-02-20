VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "iSQLConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

' Interface: iSQLConnection
' Defines a wrapper for a cusom connection or ADODB.Connection

' Property: Connection 
Public Property Get Connection()

End Property

' Property: ConnectionString 
Public Property Let ConnectionString(sConnection As String)

End Property

' Sub: OpenConnection
Public Sub OpenConnection()

End Sub

' Sub: CloseConnection
Public Sub CloseConnection()

End Sub

' Function: State 
Public Function State()

End Function

