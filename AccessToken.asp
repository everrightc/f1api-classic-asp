<% @LANGUAGE = "VBScript" %>
<% Option Explicit %>
<% Response.Buffer = True %>
<!-- #include file="includes/PersonModel.asp" -->

<%

' This page should be called after the user has been to fellowship one to log in. After the log in process
' We still need to go and get an access token that we will carry around via cookies for future API calls
' The access token request will exchange the request token we currently have in our cookies
' for an authorized access token which will be needed for all future API calls
Call RequestAccessToken

 %>

  <!-- Assuming that the access token was successful, display the xml that is returned from retrieving a person -->
  Person ID:  <%Response.Write(Request.Cookies("personID")) %><br /><br />

  <b>Person XML</b><br />
  <% 
    ' Write out the person XML
    Dim personXML : personXML = GetPerson(Request.Cookies("personID"))
    personXML = Replace(personXML, ">", "&#62;")
    personXML = Replace(personXML, "<", "&#60")
    personXML = Replace(personXML, VbCrLf, "<br />")
    Response.Write(personXML)
  %>
