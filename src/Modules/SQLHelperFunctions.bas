Attribute VB_Name = "SQLHelperFunctions"
Public Function toUnix(dt) As Long
    toUnix = DateDiff("s", "1/1/1970", dt)
End Function

Public Function toISO(dt) As String
    toISO = Format(dt, "YYYY-MM-DD") & "T" & Format(dt, "HH:MM:SS")
End Function

Public Function str(vValue) As String
    str = "'" & Replace(vValue, "'", "''") & "'"
End Function

Function Join2D(ByVal vArray As Variant, _
                Optional ByVal WordDelim As String = " ", _
                Optional ByVal LineDelim As String = vbNewLine) As String
  Dim R As Long, Lines() As String
  ReDim Lines(0 To UBound(vArray))
  For R = 0 To UBound(vArray)
    Lines(R) = Join(Application.Index(vArray, R + 1, 0), WordDelim)
  Next
  Join2D = Join(Lines, LineDelim)
End Function

Function JoinArrayofArrays(ByVal vArray As Variant, _
                Optional ByVal WordDelim As String = " ", _
                Optional ByVal LineDelim As String = vbNewLine) As String
  Dim R As Long, Lines() As String
  ReDim Lines(0 To UBound(vArray))
  For R = 0 To UBound(vArray)
    Dim InnerArray() As Variant
    InnerArray = vArray(R)
    Lines(R) = Join(InnerArray, WordDelim)
  Next
  JoinArrayofArrays = Join(Lines, LineDelim)
End Function

Function getDimension(Var As Variant) As Long
    On Error GoTo Err
    Dim i As Long
    Dim tmp As Long
    i = 0
    Do While True
        i = i + 1
        tmp = UBound(Var, i)
    Loop
Err:
    getDimension = i - 1
End Function

Public Function IsArrayAllocated(Arr As Variant) As Boolean
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' IsArrayAllocated
' Returns TRUE if the array is allocated (either a static array or a dynamic array that has been
' sized with Redim) or FALSE if the array is not allocated (a dynamic that has not yet
' been sized with Redim, or a dynamic array that has been Erased). Static arrays are always
' allocated.
'
' The VBA IsArray function indicates whether a variable is an array, but it does not
' distinguish between allocated and unallocated arrays. It will return TRUE for both
' allocated and unallocated arrays. This function tests whether the array has actually
' been allocated.
'
' This function is just the reverse of IsArrayEmpty.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim N As Long
On Error Resume Next

' if Arr is not an array, return FALSE and get out.
If IsArray(Arr) = False Then
    IsArrayAllocated = False
    Exit Function
End If

' Attempt to get the UBound of the array. If the array has not been allocated,
' an error will occur. Test Err.Number to see if an error occurred.
N = UBound(Arr, 1)
If (Err.Number = 0) Then
    ''''''''''''''''''''''''''''''''''''''
    ' Under some circumstances, if an array
    ' is not allocated, Err.Number will be
    ' 0. To acccomodate this case, we test
    ' whether LBound <= Ubound. If this
    ' is True, the array is allocated. Otherwise,
    ' the array is not allocated.
    '''''''''''''''''''''''''''''''''''''''
    If LBound(Arr) <= UBound(Arr) Then
        ' no error. array has been allocated.
        IsArrayAllocated = True
    Else
        IsArrayAllocated = False
    End If
Else
    ' error. unallocated array
    IsArrayAllocated = False
End If

End Function
