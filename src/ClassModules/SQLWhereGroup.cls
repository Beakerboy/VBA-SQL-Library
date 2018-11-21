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

Public Property Get Operator()
    Operator = BoolOperator
End Property

Public Sub SetGroup(ByVal oWhere1, ByVal oWhere2, op As String)
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
    If op = "AND" Or op = "OR" Then
        BoolOperator = op
    Else
        MsgBox op & " is not a supprorted operator"
    End If
End Sub

Public Sub AddWhere(oWhere, op As String)
    
    If Where1.Operator <> "" Then
        Set WhereGroup1 = New SQLWhereGroup
        If Where2.Operator <> "" Then
            WhereGroup1.SetGroup Where1, Where2, BoolOperator
            Set Where2 = New SQLWhere
        Else
            WhereGroup1.SetGroup Where1, WhereGroup2, BoolOperator
            Set WhereGroup2 = New SQLWhereGroup
        End If
        Set Where1 = New SQLWhere
    Else
        If Where2.Operator <> "" Then
            WhereGroup1.AddWhere Where2, BoolOperator
            Set Where2 = New SQLWhere
        Else
            WhereGroup1.AddWhere WhereGroup2, BoolOperator
            Set WhereGroup2 = New SQLWhereGroup
        End If
    End If
    If TypeOf oWhere Is SQLWhere Then
        Set Where2 = oWhere
    Else
        Set WhereGroup2 = oWhere
    End If
    If op = "AND" Or op = "OR" Then
        BoolOperator = op
    Else
        MsgBox op & " is not a supprorted operator"
    End If
End Sub

Public Function toString()
    Dim sWhere1 As String
    If Where1.Operator <> "" Then
        sWhere1 = Where1.toString
    Else
        sWhere1 = "(" & WhereGroup1.toString & ")"
    End If
    
    Dim sWhere2 As String
    If Where2.Operator <> "" Then
        sWhere2 = Where2.toString
    Else
        sWhere2 = "(" & WhereGroup2.toString & ")"
    End If
    
    toString = sWhere1 & " " & BoolOperator & " " & sWhere2
End Function
