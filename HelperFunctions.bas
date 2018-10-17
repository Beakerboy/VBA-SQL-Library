Public Function toUnix(dt) As Long
    toUnix = DateDiff("s", "1/1/1970", dt)
End Function

Public Function toISO(dt) As String
    toISO = Format(dt, "YYYY-MM-DD") & "T" & Format(dt, "HH:MM:SS")
End Function

Public Function str(vValue) As String
    str = "'" & vValue & "'"
End Function

' Given an array, join the elements together with a specified string between them.
Public Function implode(ArrayOfValues, Optional glue = ", ") As String
    initial = True
    returnString = ""
    For Each element In ArrayOfValues
        If initial Then
            initial = False
        Else
            returnString = returnString & glue
        End If
        returnString = returnString & element
    Next element
    implode = returnString
End Function
