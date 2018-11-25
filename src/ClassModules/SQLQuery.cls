VERSION 1.0 CLASS
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

Public Sub Class_Initialize()
    Set dArguments = New Dictionary
End Sub

Public Property Get Arguments()
    Set Arguments = dArguments
End Property

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
            ReturnString = Left(ReturnString, vPosition - 1 + ExtraChars) & Replace(ReturnString, sKey, sValue, ExtraChars + vPosition)
            ExtraChars = ExtraChars + Len(sValue) - Len(sKey)
        Next vPosition
    End If
    ReplaceArguments = ReturnString
End Function
