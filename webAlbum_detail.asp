<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="webAlbum_data.inc"-->
<!--#include file="show_status.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' Last updated 3/22/2000

dim intIndex
dim arPics
dim strPath
dim strDesc
dim strFirst
dim strRest
dim strDate
dim strName
dim strSource
dim strLocation
dim strQuery
dim oRS
dim pic

intIndex = CInt(Request.QueryString("index"))
arPics = Session("picList")

strQuery = "SELECT * FROM pictures WHERE " _
	& "pic_name='" & arPics(intIndex) & "'"

Set oRS = Server.CreateObject("ADODB.RecordSet")
oRS.Open strQuery, strDSN, adOpenForwardOnly, adLockReadOnly, adCmdText

strDesc = oRS("pic_description")

strDate = "<a href='../webCal/webCal3_month.asp?date=" _
	& Month(oRS("pic_taken")) & "/1/" _
	& Year(oRS("pic_taken")) & "' " _
	& showStatus("View calendar") & " target='body'>" _
	& FormatDateTime(oRS("pic_taken"),1) & "</a>"
		
Select case oRS("pic_source")
	case "video"
		strSource = "Video"
	case "scan"
		strSource = "Scanned print"
	case "digital"
		strSource = "Digital camera"
	case "web"
		strSource = "Other web page"
	case else
		strSource = ""
end Select

strLocation = oRS("pic_location")
strName = oRS("pic_name")

oRS.Close
Set oRS = nothing

if strDesc <> "" then
	strFirst = Left(strDesc, 1)
	strRest = Right(strDesc, Len(strDesc) - 1)
else
	strFirst = ""
	strRest = ""
end if
%>
<html>
<head>
<link rel='stylesheet' href='../style/ths1991.css' type='text/css'>
</head>
<body>

<img align="left" src="./pictures/hi-res/<%=strName%>.jpg" alt="" border=1>
<div align="right" class='PicNav'>
<table border=0 cellspacing=0 cellpadding=2>
<tr>
<% if intIndex > 0 then %>
	<form action="webAlbum_detail.asp?index=<%=intIndex - 1%>" method="post">
	<td><input type="submit" value="&lt;" class='button'></td>
	</form>
<% end if %>
<% if Session(strDataName & "User") <> "" then %>
	<form action="webAlbum_edit.asp?action=update&index=<%=intIndex%>" method="post">
	<td align="center"><input type="submit" value="Edit" class='button'></td>
	</form>
<% end if %>
<% if intIndex < UBound(arPics) then %>
	<form action="webAlbum_detail.asp?index=<%=intIndex + 1%>" method="post">
	<td align="right"><input type="submit" value="&gt;" class='button'></td>
	</form>
<% end if %>
</table>
</div>
<font class='DropCap'><%=strFirst%></font><%=strRest%>
<p>
<center>

<table cellspacing=0 cellpadding=2 border=0>
<% if strLocation <> "" then %>
<tr>
	<td valign="top" class='PicLabel'>Location:</td>
	<td class='PicDetail'><%=strLocation%></td>
<% end if %>
<tr>
	<td valign="top" class='PicLabel'>Date:</td>
	<td class='PicDetail'><%=strDate%></td>
<% if strSource <> "" then %>
<tr>
	<td valign="top" class='PicLabel'>Source:</td>
	<td class='PicDetail'><%=strSource%></td>
<% end if %>
</table>
</center>
</body>
</html>