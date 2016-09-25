<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="data/webAlbum_data.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' Last updated 3/22/2000

if Request.Form("logout") = "Logout" then
	Session(strDataName & "User") = ""
	response.redirect "webAlbum_admin.asp"
end if

dim strQuery		' strQuery passed to db
dim strCats			' selection box of categories
dim oConn			' connection object
dim oRS				' recordset object

select case Request.Form("add")
	case "images"
		response.redirect "webAlbum_add.asp"
	case "cat"
		strQuery = "INSERT INTO category (" _
			& "cat_name) VALUES ('" _
			& Request.Form("cat_name") & "')"
	case "subcat"
		strQuery = "INSERT INTO subcategory (" _
			& "subcat_name, " _
			& "subcat_cat" _
			& ") VALUES ('" _
			& Request.Form("subcat_name") & "', '" _
			& Request.Form("subcat_cat") & "')"
	case else
		strQuery = ""
end Select

if strQuery <> "" then
	Set oConn = Server.CreateObject("ADODB.Connection")
	oConn.Open strDSN
	oConn.Execute strQuery,,adCmdText
	oConn.Close
	Set oConn = nothing
end if

strQuery = "SELECT cat_name FROM category ORDER BY cat_name"
Set oRS = Server.CreateObject("ADODB.RecordSet")
oRS.Open strQuery, strDSN, adOpenForwardOnly, adLockReadOnly, adCmdText
do while not oRS.EOF
	strCats = strCats & "<option>" & oRS("cat_name") & vbCrLf
	oRS.Movenext
loop
oRS.Close
Set oRS = nothing
%>

<!--#include file="webAlbum_verify.inc"-->

<html>
<head>
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript">
<!--
function radioSelect(i) {
	document.adminform.add[i].checked = 1;
}
-->
</SCRIPT>
<!--#include file="webAlbum_themes.inc"-->
</head>
<body bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=arColor(5)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=arColor(11)%>" border=0 cellpadding=3 cellspacing=0>
<form name="adminform" action="webAlbum_admin.asp" method="post">
<tr bgcolor="#<%=arColor(3)%>" valign="bottom">
	<td colspan=2><font face="Tahoma, Arial, Helvetica" size=4>
	<b><nobr>webAlbum Administration</nobr></b></font></td>
<tr>
	<td colspan=2 bgcolor="#<%=arColor(12)%>"><input type="submit" value="Add"></td>
<tr>
	<td align="right" valign="top" bgcolor="#<%=arColor(12)%>"><input type="radio" name="add" value="images" checked></td>
	<td><font face="Verdana, Arial, Helvetica" size=2>New images</font></td>
<tr>
	<td align="right" valign="top" bgcolor="#<%=arColor(12)%>"><input type="radio" name="add" value="cat"></td>
	<td><font face="Verdana, Arial, Helvetica" color="#<%=arColor(14)%>" size=1>
	new category</font><br><input type="text" name="cat_name" size=15 onFocus="radioSelect(1);"></td>
<tr>
	<td align="right" valign="top" bgcolor="#<%=arColor(12)%>"><input type="radio" name="add" value="subcat"></td>
	<td><font face="Verdana, Arial, Helvetica" color="#<%=arColor(14)%>" size=1>
	new sub-category</font><br>
	<input type="text" name="subcat_name" size=15 onFocus="radioSelect(2);"><br>
	<font face="Verdana, Arial, Helvetica" color="#<%=arColor(14)%>" size=1>
	under this category:</font><br>
	<select name="subcat_cat"><%=strCats%></select>
	</td>
<tr>
	<td colspan=2 align="right" bgcolor="#<%=arColor(12)%>"><input type="submit" name="logout" value="Logout"></td>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

</form>
</body>
</html>