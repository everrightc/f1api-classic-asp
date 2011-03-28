<!-- #include file="OAuth.asp" -->
<!--#include file="Utilities.asp" -->
<% 

' This is the base API url, it is the bare minimum that will not change for each request
Dim baseF1APIUrl
baseF1APIUrl = StringFormat("https://{0}.staging.fellowshiponeapi.com/",  Array(churchCode))

' The people URL is the url that is used for any people api request
Dim peopleUrl
peopleUrl = baseF1APIUrl & "v1/"

' The giving URL is the url that is used for any giving api request
' The church and the vendor must be set up to accept giving requests in order for request to return data
Dim givingUrl
givingUrl = baseF1APIUrl & "giving/v1"



' ===============================================
' Set the request token to the cookies This is a semi overload to the SetRequestToken in OAuth.asp
' ===============================================
Sub SetOAuthRequestToken()

    Dim requestTokenXML     : Set requestTokenXML = Server.CreateObject("Microsoft.XMLHTTP")
    Dim apiCallUrl : apiCallUrl = APICall(peopleUrl & "Tokens/RequestToken", True, "GET")
    
    requestTokenXML.Open "GET", apiCallUrl, False
    requestTokenXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    requestTokenXML.Send

    ' Sets the request token and secret
    SetRequestToken(requestTokenXML.responseText)
End Sub

' =================================
' Authenticate the user
' isPortalUser = Defines whether or not the user that is being authenticate is a portal user or not
' The url for fellowship one login change depends on this boolean
' =================================
Sub AuthenticateUser(isPortalUser) 
    
    ' Before authenticating we need to get a request token
    SetOAuthRequestToken
        
    ' Construct the login url
    Dim LoginURL
    
    LoginURL = peopleUrl

    If isPortalUser Then
        LoginURL = LoginUrl & "PortalUser"
    Else
        LoginURL = LoginUrl & "WeblinkUser"
    End If
    
    LoginURL = LoginURL & "/Login?"
    LoginURL = LoginURL & "oauth_token=" & Request.Cookies("RequestToken")
    LoginURL = LoginURL & "&oauth_callback=" & callBackUrl
    LoginURL = LoginURL & "&oauth_consumer_key=" & consumerKey
    
    ' Redirect to the fellowship one login page
    Response.Redirect(LoginURL)
End Sub

' ===============================
' Request an Access Token
' ===============================
Sub RequestAccessToken()

    Dim requestXML     : Set requestXML = Server.CreateObject("Microsoft.XMLHTTP")
    Dim splitContentLocation
    
    requestXML.Open "GET", APICall(peopleUrl & "Tokens/AccessToken", False, "GET"), False
    requestXML.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    requestXML.Send

    ' Only set the access token if the response is 200 which is "OK"
    If requestXML.status = 200 Then
        ' Sets the access token and secret
        SetRequestToken(requestXML.responseText)

        ' Add the person URL to cookies for future use
        splitContentLocation = Split(requestXML.getResponseHeader("Content-Location"), "/")
        Response.Cookies("personID") = splitContentLocation(UBound(splitContentLocation))
        Response.Cookies("personID").Expires = Date + 30
    End If

End Sub


%>