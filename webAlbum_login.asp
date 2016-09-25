<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="data/webAlbum_data.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' Last updated 3/22/2000

dim strError
dim strStatus
dim strQuery
dim oRS

strStatus = "This action is available only to registered users"
strError = "The information you entered could not be validated. " _
	& "Please try again."

if Request.Form("login") <> "" then
	strQuery = "SELECT * FROM album_users WHERE " _
		& "login = '" & Request.Form("login") & "'"

	Set oRS = Server.CreateObject("ADODB.RecordSet")
	oRS.Open strQuery, strDSN, adOpenForwardOnly, adLockReadOnly, adCmdText
	
	if oRS.EOF = -1 then
		strStatus = strError
	else
		if oRS("password") = Request.Form("password") then
			Session(strDataName & "User") = oRS("user_id")
			response.redirect Request.Form("url")
		else
			strStatus = strError
		end if
	end if
	oRS.Close
	Set oRS = nothing
end if
%>

<html>
<!--#include file="webAlbum_themes.inc"-->
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=arColor(5)%>" width="60%" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=arColor(11)%>" border=0 cellpadding=3 cellspacing=0 width="100%">
<form action="webAlbum_login.asp" method="post">
<tr bgcolor="#<%=arColor(3)%>" valign="bottom">
	<td colspan=4><font face="Tahoma, Arial, Helvetica" size=4>
	<b>Login</b></font></td>
<tr>
	<td colspan=4 align="center"><font face="Arial, Helvetica" size=2>
	<%=strStatus%><br></font></td>
<tr>
	<td>&nbsp;</td>
	<td bgcolor="#<%=arColor(12)%>" align="right"><font face="Arial, Helvetica">Username:&nbsp;</td>
	<td bgcolor="#<%=arColor(12)%>"><input type="text" name="login" size=10></td>
	<td>&nbsp;</td>
<tr>
	<td>&nbsp;</td>
	<td bgcolor="#<%=arColor(12)%>" align="right"><font face="Arial, Helvetica">Password:&nbsp;</td>
	<td bgcolor="#<%=arColor(12)%>"><input type="password" name="password" size=10></td>
	<td>&nbsp;</td>
<tr>
	<td colspan=4 align="center"><br>
	<input type="submit" value="Continue"></td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<%
response.write "<input type='hidden' name='url' value='"
if Request.Form("url") <> "" then
	response.write Request.Form("url")
else
	response.write Request.QueryString("url")
end if
response.write "'>"
%>
</form>
</center>
</body>
</html>