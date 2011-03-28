<% @LANGUAGE = "VBScript" %>
<% Option Explicit %>
<% Response.Buffer = True %>
<!-- #include file="includes/F1APIOAuth.asp" -->

<%

' This assumes the vendor is third party. First thing in a third party application is to get logged in from fellowship one
' The user can be logged in as a portal user or as a web link user
' NOTE: Portal User that logs in has to be linked to an individual to be successful
Call AuthenticateUser(true)

 %>

