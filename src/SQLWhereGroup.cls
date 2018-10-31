VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLWhereGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Where1 As SQLWhere
Private Where2 As SQLWhere
Private WhereGroup1 As SQLWhereGroup
Private WhereGroup2 As SQLWhereGroup
Private BoolOperator As String      'AND or OR

Private Sub Class_Initialize()
    Set Where1 = New SQLWhere
    Set Where2 = New SQLWhere
    
End Sub

Public Sub SetGroup(oWhere1, oWhere2)
    If TypeOf oWhere1 Is SQLWhere Then
        Set Where1 = oWhere1
    Else
        Set WhereGroup1 = oWhere1
    End If
    
    If TypeOf oWhere2 Is SQLWhere Then
        Set Where2 = oWhere2
    Else
        Set WhereGroup2 = oWhere2
    End If
End Sub

Public Function ToString()
    Dim sWhere1 As String
    If Where1.Operator <> "" Then
        sWhere1 = Where1.ToString
    Else
        sWhere1 = WhereGroup1.ToString
    End If
    
    Dim sWhere2 As String
    If Where2.Operator <> "" Then
        sWhere2 = Where1.ToString
    Else
        sWhere2 = WhereGroup1.ToString
    End If
    
    ToString = "(" & sWhere1 & " " & BoolOperator & " " & sWhere2 & ")"
End Function
