VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sExpression1
Private sExpression2
Private sOperator As String

Property Get Operator() As String
    Operator = sOperator
End Property
Public Function Create(Expression1, Expression2, Optional Operator As String = "=")
    sExpression1 = Expression1
    sExpression2 = Expression2
    sOperator = Operator
End Function

Public Function toString() As String
    Dim OperatorString As String
    If sOperator = "IS" Then
        OperatorString = " " & sOperator & " "
    Else
        OperatorString = sOperator
    End If
    toString = sExpression1 & OperatorString & sExpression2
End Function
