VERSION 1.0 CLAS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SQLQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private dArguments As Dictionary
Private oWhere As SQLCondition              'Can have either a SQLWhere or a SQLWhereGroup
Private oWhereGroup As SQLWhereGroup

Public Sub Class_Initialize()
    Set dArguments = New Dictionary
End Sub

Public Property Get Arguments()
    Set Arguments = dArguments
End Property

'Add a single WHERE clause to the SQL statement
' op is the operation
' value is the value
' EXAMPLE: addwhere 'id' '=', 2
'   is equivalent to "WHERE id=2"
'
Public Sub AddWhere(Field, Value, Optional op As String = "=", Optional GroupType As String = "AND")
    Dim NewWhere As New SQLCondition
    If Not (oWhere Is Nothing) Then
        NewWhere.Create Field, Value, op
        Set oWhereGroup = New SQLWhereGroup
        oWhereGroup.SetGroup oWhere, NewWhere, GroupType
        
        'Clear SQLWhere since we are using SQLWhereGroup instead
        Set oWhere = Nothing
    ElseIf oWhereGroup Is Nothing Then
        Set oWhere = New SQLCondition
        oWhere.Create Field, Value, op
    Else
        NewWhere.Create Field, Value, op
        oWhereGroup.AddWhere NewWhere, GroupType
    End If
End Sub

Public Sub AddWhereGroup(NewWhereGroup As SQLWhereGroup, Optional GroupType As String = "AND")
    If oWhere Is Nothing Then
        If oWhereGroup Is Nothing Then
            Set oWhereGroup = NewWhereGroup
        Else
            oWhereGroup.AddWhere NewWhereGroup, GroupType
        End If
    Else
        Set oWhereGroup = New SQLWhereGroup
        oWhereGroup.SetGroup oWhere, NewWhereGroup, GroupType
        Set oWhere = Nothing
    End If
End Sub

Public Sub ClearArguments()
    Set dArguments = New Dictionary
End Sub

Public Sub AddArgument(sName As String, sValue)
    If Left(sName, 1) = ":" Then
        If VarType(sValue) = vbString Then
            sValue = str(sValue)
        End If
        dArguments(sName) = sValue
    End If
End Sub

Public Function ReplaceArguments(sQuery As String)
    Dim dPosition As New Dictionary
    Dim aToSort() As Integer
    Dim iDictSize As Integer
    iDictSize = dArguments.count
    Dim ReturnString As String
    ReturnString = sQuery
    If iDictSize <> 0 Then
        ReDim aToSort(iDictSize - 1)
        Dim sValue As String
        Dim sKey As Variant
        Dim Num As Integer
        Dim Position As Integer
        Num = 0
        For Each sKey In dArguments.Keys
            Position = InStr(sQuery, sKey)
            aToSort(Num) = Position
            dPosition(Position) = sKey
            Num = Num + 1
        Next sKey
        QuickSort aToSort, 0, UBound(aToSort)
        Dim ExtraChars As Integer
        ExtraChars = 0
        Num = 0
        Dim vPosition As Variant
        For Each vPosition In aToSort
            sKey = dPosition(vPosition)
            sValue = dArguments(sKey)
            ReturnString = Left(ReturnString, vPosition - 1 + ExtraChars) & Replace(ReturnString, sKey, sValue, ExtraChars + vPosition, 1)
            ExtraChars = ExtraChars + Len(sValue) - Len(sKey)
        Next vPosition
    End If
    ReplaceArguments = ReturnString
End Function

Public Function WhereString() As String
    Dim sWhere As String
    If Not (oWhere Is Nothing And oWhereGroup Is Nothing) Then
        If Not (oWhere Is Nothing) Then
            sWhere = oWhere.toString
        Else
            sWhere = oWhereGroup.toString
        End If
        WhereString = " WHERE " & sWhere
    End If
End Function
