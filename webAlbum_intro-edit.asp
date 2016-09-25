<% Option Explicit %>
<% Response.Buffer = True %>
<html>
<head>
<!--#include file="webAlbum_themes.inc"-->
<!--#include file="data/webAlbum_data.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' Last updated 3/28/2000

dim strQuery		' query passed to db
dim oRS				' recordset object
dim strLevel		' category or subcategory
dim intID			' level id
dim strIntro		' intro text
dim strTitle		' title text at page top

Set oRS = Server.CreateObject("ADODB.Recordset")
strLevel = Request.Form("level")
intID = Request.Form("id")

if Request.Form("save") = "Save" then
	strIntro = Replace(Request.Form("intro"), "'", "''")
	strQuery = "UPDATE " & strLevel & "egory SET " _
		& strLevel & "_intro = '" & strIntro & "'" _
		& "WHERE (" & strLevel & "_id)=" & intID
	oRS.Open strQuery, strDSN, adOpenForwardOnly, adLockReadOnly, adCmdText
	response.redirect "webAlbum_intro.asp?" & strLevel & "=" & intID
elseif Request.Form("cancel") = "Cancel" then
	response.redirect "webAlbum_intro.asp?" & strLevel & "=" & intID
end if

strQuery = "SELECT * FROM " & strLevel & "egory " _
	& "WHERE (" & strLevel & "_id)=" & intID
oRS.Open strQuery, strDSN, adOpenForwardOnly, adLockReadOnly, adCmdText
strTitle = oRS(strLevel & "_name")
strIntro = oRS(strLevel & "_intro")
oRS.Close
Set oRS = nothing
%>

</head>
<body bgcolor="#<%=arColor(1)%>" text="#000000" link="#<%=arColor(5)%>" vlink="#<%=arColor(5)%>" alink="#<%=arColor(7)%>">

<font face="Tahoma, Arial, Helvetica" size="5" color="#ffffff">
<b><%=strTitle%></b></font>
<hr size="1" color="#000000">

<form action="webAlbum_intro-edit.asp" method="post">
<textarea name="intro" cols=50 rows=20 wrap="virtual"><%=strIntro%></textarea>
<br>
<input type="submit" name="save" value="Save">
<input type="submit" name="cancel" value="Cancel">
<input type="hidden" name="level" value="<%=strLevel%>">
<input type="hidden" name="id" value="<%=intID%>">
</form>
</body>
</html>