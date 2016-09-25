<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' Last updated 3/22/2000

' check to see if the user has already logged in
' if not, send them to the login page

dim strPath

strPath = Request.ServerVariables("PATH_TRANSLATED")
strPath = Right(strPath, Len(strPath) - InStrRev(strPath,"\"))

if Session(strDataName & "User") = "" OR Session(strDataName & "User") = 0 then
	response.redirect "webAlbum_login.asp?url=" _
		& strPath & "?" _
		& Server.URLEncode(Request.ServerVariables("QUERY_STRING"))
end if
%>