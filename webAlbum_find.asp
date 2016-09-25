<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="webAlbum_data.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' Last updated 3/22/2000

dim strQuery
dim strCat
dim oRS
dim strSelect

strQuery = "SELECT S.subcat_id, S.subcat_name, S.subcat_cat, " _
	& "C.cat_id, C.cat_name " _
	& "FROM category C INNER JOIN subcategory S " _
	& "ON (C.cat_id = S.subcat_cat) " _
	& "ORDER BY C.cat_name, S.subcat_name"

Set oRS = Server.CreateObject("ADODB.RecordSet")
oRS.Open strQuery, strDSN, adOpenForwardOnly, adLockReadOnly, adCmdText

do while not oRS.EOF
	if strCat <> oRS("cat_id") then
		strCat = oRS("cat_id")
		strSelect = strSelect & "<option value=""0|" & strCat _
			& """>[" & oRS("cat_name") & "]" & VbCrLf
	end if

	strSelect = strSelect & "<option value='1|" & oRS("subcat_id") _
		& "'>&nbsp;&nbsp;&nbsp;" & oRS("subcat_name") & VbCrLf

	oRS.Movenext
loop
oRS.Close
Set oRS = nothing
%>

<html>
<head>
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" SRC="calendar_popup.js"></SCRIPT>
<SCRIPT LANGUAGE="javascript">
<!--

function Validate() {
	var selCat = document.findform.pic_cat.options[document.findform.pic_cat.selectedIndex].value + '';
	if (document.findform.pic_keywords.value.length <= 0
    && document.findform.pic_description.value.length <= 0
 	 && selCat == "all"
	 && document.findform.date_start.value.length <= 0
	 && document.findform.date_end.value.length <= 0) {
		alert("You must enter criteria in at least one field");
		document.findform.pic_keywords.select();
		document.findform.pic_keywords.focus();
		return false;
	}
}
//-->
</SCRIPT>
<!--#include file="webAlbum_themes.inc"-->
</head>
<body onload="init();" bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">
<center>
<font face="Verdana, Arial, Helvetica">
<% if Request.QueryString("retry") then %>
<font size=2>
No pictures matched your query.<br>Please try different parameters:<br>
</font>
<% end if %>

<!-- framing table -->
<table bgcolor="#<%=arColor(5)%>" cellspacing=0 cellpadding=2 border=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=arColor(11)%>" cellspacing=0 cellpadding=2 border=0>
<form name="findform" action="webAlbum_found-frame.asp" target="body" method="post" onSubmit="return Validate();">
<tr bgcolor="#<%=arColor(3)%>">
	<td colspan=2><font face="Tahoma, Arial, Helvetica" size=4>
	<b>Find Pictures</b></font></td>
<tr>
	<td align="right"><font face="Tahoma, Arial, Helvetica" size=2>Key-words: </font></td>
	<td><input type="text" name="pic_keywords" size=20></td>
<tr>
	<td align="right"><font face="Tahoma, Arial, Helvetica" size=2>Description: </font></td>
	<td><input type="text" name="pic_description" size=20></td>
<tr>
	<td align="right" valign="top"><font face="Tahoma, Arial, Helvetica" size=2>Category: </font></td>
	<td><select name="pic_cat">
	<option value="all">[All]
	<%=strSelect%>
	</select>
	</td>
<tr>
	<td align="right"><font face="Tahoma, Arial, Helvetica" size=2>Between: </font></td>
	<td><input type="text" name="date_start" size=20><input type="button" value="&gt;" onClick="calpopup(3);"></td>
<tr>
	<td align="right"><font face="Tahoma, Arial, Helvetica" size=2>and: </font></td>
	<td><input type="text" name="date_end" size=20><input type="button" value="&gt;" onClick="calpopup(5);"></td>
<tr>
	<td align="right"><font face="Tahoma, Arial, Helvetica" size=2>Show: </font></td>
	<td><font face="Tahoma, Arial, Helvetica" size=2>
	<select name="per_page">
	<option>All
	<option selected>10
	<option>20
	<option>30
	<option>40
	<option>50
	</select> pictures at a time</font></td>
<tr>
	<td align="right" colspan=2 bgcolor="#<%=arColor(12)%>">
	<input type="submit" value="Find"></td>
</form>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<table cellspacing=4 cellpadding=2 border=0>
<tr>
	<td colspan=2 align="center">
	<font face="Tahoma, Arial, Helvetica" color="#<%=arColor(5)%>"><b>Examples</b></font></td>
<tr>
	<td align="center" bgcolor="#<%=arColor(2)%>">
	<font face="Verdana, Arial, Helvetica" size=2>to match</font></td>
	<td align="center" bgcolor="#<%=arColor(2)%>">
	<font face="Verdana, Arial, Helvetica" size=2>use</font></td>
<tr>
	<td align="right">
	<font face="Verdana, Arial, Helvetica" size=2>"dog" <u>or</u> "cat"</font></td>
	<td bgcolor="#<%=arColor(11)%>">
	<font face="Verdana, Arial, Helvetica" size=2>dog cat</td>
<tr>
	<td align="right">
	<font face="Verdana, Arial, Helvetica" size=2>both "dog" <u>and</u> "cat"</font></td>
	<td bgcolor="#<%=arColor(11)%>">
	<font face="Verdana, Arial, Helvetica" size=2>dog+cat</td>
<tr>
	<td align="right">
	<font face="Verdana, Arial, Helvetica" size=2>the <u>phrase</u> "dog cat"</font></td>
	<td bgcolor="#<%=arColor(11)%>">
	<font face="Verdana, Arial, Helvetica" size=2>"dog cat"</td>
<tr>
	<td align="right">
	<font face="Verdana, Arial, Helvetica" size=2><u>without</u> "dog"</font></td>
	<td bgcolor="#<%=arColor(11)%>">
	<font face="Verdana, Arial, Helvetica" size=2>-dog</td>
<tr>
	<td align="right" valign="top">
	<font face="Verdana, Arial, Helvetica" size=2>taken in 1997</font></td>
	<td bgcolor="#<%=arColor(11)%>" valign="top">
	<font face="Verdana, Arial, Helvetica" size=2>1/1/97<br>1/1/98</td>
<tr>
	<td align="right" valign="top">
	<font face="Verdana, Arial, Helvetica" size=2>taken between<br>October 1998 and now</font></td>
	<td bgcolor="#<%=arColor(11)%>" valign="top">
	<font face="Verdana, Arial, Helvetica" size=2>10/1/98<br>[leave blank]*</td>
</table>
</center>
<p>
<font size=1>
*if you enter a value for one date and leave the other blank, the program will assume the current date for the blank field
</font></font>

</body>
</html>