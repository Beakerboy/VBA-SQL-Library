Attribute VB_Name = "HelperFunctions"
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
  ReDim Lines(1 To UBound(vArray))
  For R = 1 To UBound(vArray)
    Lines(R) = Join(Application.Index(vArray, R, 0), WordDelim)
  Next
  Join2D = Join(Lines, LineDelim)
End Function
