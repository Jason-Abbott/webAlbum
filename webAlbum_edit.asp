<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="webAlbum_data.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' Last updated 3/22/2000

dim intIndex
dim arPics
dim strAction
dim strName
dim strDesc
dim strSource
dim strCount
dim strLocation
dim strKeyWords
dim strDate
dim strPath
dim strSubCat
dim strCat
dim strQuery
dim strSelect
dim intPicID
dim oConn
dim oRS

strAction = Request.QueryString("action")
intIndex = CInt(Request.QueryString("index"))
arPics = Session("picList")
strName = arPics(intIndex)
Set oConn = Server.CreateObject("ADODB.Connection")
Set oRS = Server.CreateObject("ADODB.RecordSet")
oConn.Open strDSN

if strAction = "update" then
	' retrieve picture details
	strQuery = "SELECT * FROM pictures WHERE pic_name='" & strName & "'"
	oRS.Open strQuery, oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	intPicID = CInt(oRS("pic_id"))
	strDesc = oRS("pic_description")
	strDate = oRS("pic_taken")
	strLocation = oRS("pic_location")
	strKeyWords = oRS("pic_keywords")
	strSource = oRS("pic_source")
	oRS.Close
	
	' retrieve category information
	strQuery = "SELECT * FROM picmix WHERE pic_id=" & intPicID
	oRS.Open strQuery, oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
	do while not oRS.EOF
		strSubCat = strSubCat & "," & oRS("subcat_id") & ","
		oRS.MoveNext
	loop
	oRS.Close
	strCount = ""
elseif strAction = "new" then
	strDesc = ""
	strDate = Date
	strLocation = ""
	strKeyWords = ""
	strSubCat = ""
	strSource = ""
	
	strCount = " (" & intIndex + 1 & " of " & UBound(arPics) + 1 & " new)"
end if

strPath = Server.Mappath(".\pictures\hi-res") & Chr(92) & strName & ".jpg"

' retrieve categories and sub-categories
strQuery = "SELECT S.subcat_id, S.subcat_name, S.subcat_cat, " _
	& "C.cat_id, C.cat_name " _
	& "FROM category C INNER JOIN subcategory S " _
	& "ON (C.cat_id = S.subcat_cat) " _
	& "ORDER BY C.cat_name, S.subcat_name"
oRS.Open strQuery, oConn, adOpenForwardOnly, adLockReadOnly, adCmdText
do while not oRS.EOF
	if strCat <> oRS("cat_id") then
		strCat = oRS("cat_id")
		strSelect = strSelect & "<option value=0>[" & oRS("cat_name") & "]" & VbCrLf
	end if
	strSelect = strSelect & "<option value=""" & oRS("subcat_id") & """"
	
	if Instr(strSubCat, "," & oRS("subcat_id") & ",") <> 0 then
		strSelect = strSelect & " selected"
	end if
	
	strSelect = strSelect & ">&nbsp;&nbsp;" & oRS("subcat_name") & VbCrLf
	oRS.Movenext
loop

oRS.Close
Set oRS = nothing
oConn.Close
Set oConn = nothing
	
' ~!@#$%^&*(){}\[\]"'`:;/\\|<>?+=
' ~ ! @ # $ % ^ & * (  { } [ ] ' ` : | < > ? + =
%>
<html>
<head>
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript" SRC="calendar_popup.js"></SCRIPT>
<SCRIPT TYPE="text/javascript" LANGUAGE="javascript">
function Validate() {
	var strName = document.editform.pic_name.value;
	if (strName.match(/[\s~!@#$%^&*(){}\[\]"'`:;/\\|<>?+=]/)) {
		alert("The name cannot contain any of these characters: ~ ! @ # $ % ^ & * ( ) { } [ ] ' ` : | < > ? + = or space");
		document.editform.pic_name.select();
		document.editform.pic_name.focus();
		return false;
	}
}
</SCRIPT>
<!--#include file="webAlbum_themes.inc"-->
</head>
<body onload="init();" bgcolor="#<%=arColor(1)%>" link="#<%=arColor(7)%>" vlink="#<%=arColor(7)%>" alink="#<%=arColor(6)%>">
<center>

<!-- framing table -->
<table bgcolor="#<%=arColor(5)%>" border=0 cellpadding=2 cellspacing=0><tr><td>
<!-- end framing table -->

<table bgcolor="#<%=arColor(11)%>" border=0 cellpadding=3 cellspacing=0>
<form name="editform" action="webAlbum_updated.asp" method="post">
<tr>
	<td colspan=2 bgcolor="#<%=arColor(3)%>">
	<font face="Tahoma, Arial, Helvetica" size=4>
	<b>Picture Details</b><%=strCount%></font></td>
<tr>
	<td rowspan=2 bgcolor="#<%=arColor(3)%>" align="center" valign="center">
	<img src="./pictures/hi-res/<%=strName%>.jpg" border=1></td>
	<td>
	<font face="Verdana, Arial, Helvetica" color="#<%=arColor(14)%>" size=1>
	Name:</font><br>
	<input type="text" name="pic_name" value="<%=strName%>"><br>

	<font face="Verdana, Arial, Helvetica" color="#<%=arColor(14)%>" size=1>
	Categories:</font><br>
	<select name="pic_cat" size="5" multiple><%=strSelect%></select><br>
	
	<font face="Verdana, Arial, Helvetica" color="#<%=arColor(14)%>" size=1>
	Date taken:</font><br>
	<input type="text" name="pic_taken" value="<%=strDate%>" size=18><input type="button" value="&gt;" onClick="calpopup(2);"><br>

	<font face="Verdana, Arial, Helvetica" color="#<%=arColor(14)%>" size=1>
	Place taken:</font><br>
	<input type="text" name="pic_location" value="<%=strLocation%>" size=20><br>
	
	<font face="Verdana, Arial, Helvetica" color="#<%=arColor(14)%>" size=1>
	Source:</font><br>
	<select name="pic_source">
	<option value="scan" <%if strSource = "scan" then%>selected<%end if%>>Scanned print
	<option value="digital" <%if strSource = "digital" then%>selected<%end if%>>Digital camera
	<option value="video" <%if strSource = "video" then%>selected<%end if%>>Captured video
	<option value="web" <%if strSource = "web" then%>selected<%end if%>>Web page
	</select>
	</td>
<tr>
	<td align="center" bgcolor="#<%=arColor(12)%>">
	<input type="submit" name="save" value="Save" onClick="return Validate();">
<% if strAction = "update" then %>
	<input type="submit" name="cancel" value="Cancel">
<% elseif strAction = "new" then %>
	<input type="submit" name="skip" value="Skip">
<% end if %>
	</td>
<tr>
	<td colspan=2>
	<font face="Verdana, Arial, Helvetica" color="#<%=arColor(14)%>" size=1>
	Key words:</font><br>
	<input type="text" name="pic_keywords" size=60 value="<%=strKeyWords%>"><br>
	<font face="Verdana, Arial, Helvetica" color="#<%=arColor(14)%>" size=1>
	Description:</font><br>
	<textarea name="pic_description" cols=55 rows=10 wrap="virtual"><%=strDesc%></textarea>
	</td>
</center>
</table>

<!-- framing table -->
</td></table>
<!-- end framing table -->

<input type="hidden" name="index" value=<%=intIndex%>>
<input type="hidden" name="action" value="<%=strAction%>">
</form>
</body>
</html>