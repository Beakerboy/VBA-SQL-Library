VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private dArguments As Dictionary

Public Sub Class_Initialize()
    Set dArguments = New Dictionary
End Sub

Public Property Get Arguments()
    Set Arguments = dArguments
End Property

Public Sub AddArgument(sName As String, sValue)
    If VarType(sValue) = vbString Then
        sValue = str(sValue)
    End If
    dArguments(sName) = sValue
End Sub
