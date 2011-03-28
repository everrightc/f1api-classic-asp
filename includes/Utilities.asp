<%

' C# Equivilant to the String.Format
Function StringFormat(sVal, aArgs)
Dim i
    For i=0 To UBound(aArgs)
        sVal = Replace(sVal,"{" & CStr(i) & "}",aArgs(i))
    Next
    StringFormat = sVal
End Function

%>