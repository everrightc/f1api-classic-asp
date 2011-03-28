<!-- #include file="sha1_js.asp" -->
<!-- #include file="Config.asp" -->
<% 

Dim OAuthBaseChars
Dim timestamp
Dim returnedNonce
Dim rndNumber
Dim i
Dim baseRequest
Dim queryStringValues
Dim rawURLRequest
Dim apiRequestURL
Dim boolIsRequestToken  : boolIsRequestToken = True
dim requestMethod

OAuthBaseChars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXTZabcdefghiklmnopqrstuvwxyz"

' =========================================
' Function to return the current time stamp
' =========================================
Function CurrentTimeStamp()
    'CurrentTimeStamp = Int(DateDiff("ms", "01/01/1970 00:00:00", Now()))
    CurrentTimeStamp = js_timestamp()
End Function

' =========================================
' Function to return a random nonce
' =========================================
Function CurrentNonce(length) 
    
    returnedNonce = ""
    
    For i = 1 to length
        ' Get a random number that we will use to find a char in the OAuth Base Chars
        returnedNonce = returnedNonce + Mid(OAuthBaseChars, RandomNumber, 1)
    Next
    
    CurrentNonce = returnedNonce
        
End Function

' =======================================
' Create the signature for the request
' =======================================
Function OAuthSignature() 
    OAuthSignature = b64_hmac_sha1(SigningKey, BaseRawRequest(RequestURL()))
End Function

' ======================================
' Create the Signed request
' ======================================
Sub SignRequest()
    ' Add the oauth signature
    queryStringValues = queryStringValues & "&oauth_signature=" & ReplaceOAuthCharacters(OAuthSignature())
End Sub


' =======================================
' Get the request URL
' =======================================
Function RequestURL() 
    RequestURL = apiRequestURL
End Function

' =======================================
' Generate a random number
' =======================================
Function RandomNumber()
    Randomize
    RandomNumber = Int((len(OAuthBaseChars) - 1)*Rnd() + 1)
End Function

' ========================================
' Generates the raw request of the OAuth call
' ========================================
Function BaseRawRequest(url)

    If InStr(url, "?") > 0 Then
        url = ReplaceOAuthCharacters(Mid(url, 1, InStr(url, "?") - 1))
    End If
    
    BaseRawRequest = requestMethod & "&" & ReplaceOAuthCharacters(url) & "&" & ReplaceOAuthCharacters(Mid(SortOAuthQueryString(), 2, len(SortOAuthQueryString())))
End Function

' ============================================
' Replaces characters for the raw request
' ============================================
Function ReplaceOAuthCharacters(str) 

    str = Replace(str, "!", "%21")
    str = Replace(str, "*", "%2A")
    str = Replace(str, "'", "%27")
    str = Replace(str, "(", "%28")
    str = Replace(str, "(", "%29")
    str = Replace(str, ";", "%3B")
    str = Replace(str, ":", "%3A")
    str = Replace(str, "@", "%40")
    str = Replace(str, "&", "%26")
    str = Replace(str, "=", "%3D")
    str = Replace(str, "+", "%2B")
    str = Replace(str, "$", "%24")
    str = Replace(str, ",", "%2C")
    str = Replace(str, "/", "%2F")
    str = Replace(str, "?", "%3F")
    str = Replace(str, "$", "%24")
    
    ReplaceOAuthCharacters = str
End Function

' =========================================
' Create the signing key for signing the URL
' =========================================
Function SigningKey() 
    
    Dim key : key = consumerSecret & "&"
    
    If boolIsRequestToken = False Then
        ' If there is an authorization token in session, use it
        If Request.Cookies("RequestTokenSecret") <> "" Then
            key = key & Request.Cookies("RequestTokenSecret")
        End If
    End If
    
    SigningKey = key
end Function

' ======================================
' Create the OAuth Query String based on the qs values passed in the url and the oauth qs values
' ======================================
Sub CreateOAuthQueryStringValues() 

        ' If there are query string values attached to the URL, keep them
        IF InStr(RequestURL(), "?") > 0 Then
            queryStringValues = Mid(RequestURL(), InStr(RequestURL(), "?") + 1, len(RequestURL()))
            queryStringValues = queryStringValues & "&"
        
        Else
            queryStringValues = ""
        End If

        ' Get the consumer key
        queryStringValues = queryStringValues & "oauth_consumer_key=" & consumerKey
        
        ' Get the oauth version
        queryStringValues = queryStringValues & "&oauth_version=1.0"
      
        ' Get the nonce
        queryStringValues = queryStringValues & "&oauth_nonce=" & CurrentNonce(6)

        ' Get the timestamp
        queryStringValues = queryStringValues & "&oauth_timestamp=" & CurrentTimeStamp()
      
        ' Get the signature method
        queryStringValues = queryStringValues & "&oauth_signature_method=HMAC-SHA1"
      
        If boolIsRequestToken = False Then
            ' Get oauth token
            queryStringValues = queryStringValues & "&oauth_token=" & Request.Cookies("RequestToken")
        End If
End Sub

Function SortOAuthQueryString() 
    Dim splitQueryString    : splitQueryString = Split(queryStringValues, "&")
    Dim sortedQueryStrings
    Dim rVal
    Dim counter

   sortedQueryStrings = Split(SortVBArray(splitQueryString), Chr(8))

   ' Now that the query string is sorted, concatenate it by a &
   For counter = 0 to UBound(sortedQueryStrings)
      IF counter = 0 Then
        rVal = "?"
      Else
        rVal = rVal & "&"
      End IF
      
      rVal = rVal & sortedQueryStrings(counter)
   Next

    SortOAuthQueryString = rVal
End Function

Function CommaSeperateValues
    CommaSeperateValues = Replace(SortOAuthQueryString(), "&", ",")
End Function

' ====================================================
' Constructs the URL for the API call
' ====================================================
Function APICall(url, isRequestToken, formMethod)

    apiRequestURL = url
    boolIsRequestToken = isRequestToken
    requestMethod = formMethod
    
    ' Create the OAuth query string values
    CreateOAuthQueryStringValues()
    
    ' Sign the request
    SignRequest()
    
    If InStr(url, "?") > 0 Then
        ' We have to cut everything after the ? if it exists
        url = Mid(url, 1, InStr(url, "?") - 1)
    End If
    
    APICall = url & SortOAuthQueryString()
    
End Function

' ======================================
' Sets the request token and request secret from the response
' ======================================
Sub SetRequestToken(strRequestToken)

    Dim responseSplitCounter
    Dim splitResponse
    Dim splitRequestKeyValuePair

    splitResponse = Split(strRequestToken, "&")
    
    For responseSplitCounter = 0 To UBound(splitResponse)
        splitRequestKeyValuePair = Split(splitResponse(responseSplitCounter), "=")
        
        Select Case splitRequestKeyValuePair(0)
            Case "oauth_token"
                Response.Cookies("RequestToken") = splitRequestKeyValuePair(1)
                Response.Cookies("RequestToken").Expires = Date + 30
            Case "oauth_token_secret"
                Response.Cookies("RequestTokenSecret") = splitRequestKeyValuePair(1)
                Response.Cookies("RequestTokenSecret").Expires = Date + 30
        End Select
    Next
End Sub
%>

<script language="javascript" runat="server">
    function js_timestamp() {
        var d = new Date();
        return Math.floor(d.getTime() / 1000);
    }
    
    function SortVBArray(arrVBArray) {
        return arrVBArray.toArray().sort().join('\b');
    }
</script>