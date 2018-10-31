VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLSubselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private MySelect As SQLSelect
Private sAs As String

Private Sub Class_Initialize()
    Set MySelect = New SQLSelect
End Sub

Property Set SelectSQL(oValue As SQLSelect)
    Set MySelect = oValue
End Property

Property Let SelectAs(sValue As String)
    sAs = sValue
End Property

Property Get SelectSQL() As SQLSelect
    Set SQLSelect = MySelect
End Property

Property Get SelectAs() As String
    SelectAs = aAs
End Property

Public Function ToString() As String
    ToString = "(" & MySelect.ToString & ") AS " & sAs
End Function
