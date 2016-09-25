<%
' Copyright 1999 Jason Abbott (jason@webott.com)
' Last updated 04/24/1999

' check to see if the user has already logged in
' if not, send them to the login page

if Session("jasonAlbumUser") = "" then
	response.redirect "authenticate.asp?url=" _
		& Request.ServerVariables("URL") & "?" _
		& Request.ServerVariables("QUERY_STRING")
end if
%>