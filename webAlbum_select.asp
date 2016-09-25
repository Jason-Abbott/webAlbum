<html>
<head>
<!--#include file="data/webAlbum_data.inc"-->
<!--#include file="webAlbum_themes.inc"-->
</head>
<body bgcolor="#000000" text="#000000" link="#000000" vlink="#000000" alink="#<%=arColor(6)%>">
<center>
<font face="Arial, Helvetica" size=1 color="#777777">select</font>
<table cellpadding=2 cellspacing=2 border=0 width="100%">

<!--#include file="show_status.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' Last updated 3/22/2000

dim strQuery
dim oRS
dim strSubDesc
dim bFirst

strQuery = "SELECT subcat_id, subcat_name, subcat_description, " _
	& "subcat_cat FROM subcategory WHERE subcat_cat=" _
	& Request.QueryString("cat") & " ORDER BY " _
	& "subcat_description, subcat_name"

Set oRS = Server.CreateObject("ADODB.RecordSet")
oRS.Open strQuery, strDSN, adOpenForwardOnly, adLockReadOnly, adCmdText

bFirst = 1
do while not oRS.EOF
	if strSubDesc <> oRS("subcat_description") then
		strSubDesc = oRS("subcat_description")
		
		response.write "<tr><td align=""center"" bgcolor=""#" _
			& arColor(3) & """ background=""../media/corner.gif"">" _
			& "<font face=""Arial, Helvetica"" size=2><b>" _
			& strSubDesc & "</b></font></td>" & VbCrLf
		bFirst = 0
	end if

	response.write "<tr><td align=""center"" "
	
	if bFirst then
		response.write "background=""../media/corner.gif"" "
		bFirst = 0
	end if
	
	response.write "onmouseover=""" _
		& "this.style.backgroundColor='#" & arColor(0) & "';"" " _
		& "onmouseout=""this.style.backgroundColor='#" _
		& arColor(1) & "';"" bgcolor=""#" _
		& arColor(1) & """><font face=""Arial, Helvetica"" size=2>" _
		& "<a href=""webAlbum_view-frame.asp?subcat=" _
		& oRS("subcat_id") & """ target=""view"" " _
		& ShowStatus("View " & oRS("subcat_name") & " pictures") & ">" _
		& oRS("subcat_name") & "</a></font></td>" & vbCrLf
	oRS.Movenext
loop
oRS.Close
Set oRS = nothing
%>
<tr>
	<td><font size=1>&nbsp;</font></td>
<tr>
	<td onmouseover="this.style.backgroundColor='#<%=arColor(0)%>';"
	 onmouseout="this.style.backgroundColor='#<%=arColor(1)%>';"
	 align="center" bgcolor="#<%=arColor(1)%>">
	<font face="Arial, Helvetica" size=2><b>
	<a href="webAlbum_find.asp" target="view"
	<%=ShowStatus("Find specific pictures")%>>Search</a></b>
	</font>
	</td>

</table>
</center>
</body>
</html>
