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

