VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLWhere"
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
Public Function Create(Expression1, Operator As String, Expression2)
    sExpression1 = Expression1
    sExpression2 = Expression2
    sOperator = Operator
End Function

Public Function ToString() As String
    ToString = sExpression1 & " " & sOperator & " " & sExpression2
End Function
