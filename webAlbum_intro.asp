<% Response.Buffer = True %>
<!--#include file="webAlbum_data.inc"-->
<!--#include file="webAlbum_themes.inc"-->
<body bgcolor="#<%=arColor(1)%>" text="#000000" link="#<%=arColor(5)%>" vlink="#<%=arColor(5)%>" alink="#<%=arColor(7)%>">
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' Last updated 3/22/2000

dim strQuery
dim oRS
dim strTitle
dim strIntro
dim strLevel
dim arName

if Request.QueryString("cat") <> "" then
	strLevel = "cat"
else
	strLevel = "subcat"
end if

intLevel = Request.QueryString(strLevel)

strQuery = "SELECT * FROM " & strLevel & "egory WHERE (" & strLevel & "_id) = " & intLevel
Set oRS = Server.CreateObject("ADODB.RecordSet")
oRS.Open strQuery, strDSN, adOpenForwardOnly, adLockReadOnly, adCmdText

if strLevel = "subcat" then
	if CInt(oRS("subcat_cat")) = 1 then
		arName = Split(oRS("subcat_name"),", ")
		response.redirect "../people/people_detail.asp?first=" & arName(1) & "&last=" & arName(0)
	end if
end if
strIntro = oRS(strLevel & "_intro")

strTitle = oRS(strLevel & "_name")

oRS.Close
Set oRS = nothing
%>

<font face="Tahoma, Arial, Helvetica" size="5" color="#ffffff">
<b><%=strTitle%></b></font>
<hr size="1" color="#000000">

<%=strIntro%>

<% if Session(strDataName & "User") <> "" then %>
<div align="right">
<form action="webAlbum_intro-edit.asp" method="post">
<input type="submit" value="Edit">
<input type="hidden" name="level" value="<%=strLevel%>">
<input type="hidden" name="id" value="<%=intLevel%>">
</form>
</div>
<% end if %>
</body>
</html>