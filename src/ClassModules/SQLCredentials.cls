VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Credentials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*****************************CREDENTIALS CLASS********************************
'Passes the login information to the database to ensure credentials are not
'hard-coded
'
'*********************************VARIABLES************************************
Private sName As String
Private sPass As String
'***************************CONSTRUCTOR****************************************

'***************************LET PROPERTIES*************************************
Public Property Let Pass(sValue As String)
    sPass = sValue
End Property

Public Property Let Name(sValue As String)
    sName = sValue
End Property

'***************************GET PROPERTIES*************************************
Public Property Get Name() As String
    Name = sName
End Property

Public Property Get Pass() As String
    Pass = sPass
End Property
