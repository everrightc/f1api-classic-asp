<!-- #include file="F1APIOAuth.asp" -->
<%
Function GetPerson(personID)
    
    Dim requestXML     : Set requestXML = Server.CreateObject("Microsoft.XMLHTTP")
    Dim responseSplitCounter
    Dim splitResponse
    Dim splitRequestKeyValuePair
    
    requestXML.Open "GET", APICall(peopleUrl & "People/" & personID, False, "GET"), False
    requestXML.setRequestHeader "Content-Type", "application/xml"
    requestXML.Send

    ' Only set the access token if the response is 200 which is "OK"
    If requestXML.status = 200 Then
        GetPerson = requestXML.responseText
    End If
End Function
 %>