<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="webAlbum_data.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' Last updated 3/22/2000

dim strQuery
dim oRS
dim strPics
dim strSubCat

strSubCat = CStr(Request.QueryString("subcat"))

strQuery = "SELECT P.pic_id, P.pic_name, P.pic_taken, M.subcat_id " _
		& "FROM pictures P INNER JOIN picmix M " _
		& "ON P.pic_id = M.pic_id WHERE " _
		& "M.subcat_id=" & strSubCat & " " _
		& "ORDER BY P.pic_taken"
			
Set oRS = Server.CreateObject("ADODB.RecordSet")
oRS.Open strQuery, strDSN, adOpenForwardOnly, adLockReadOnly, adCmdText
do while not oRS.EOF
	strPics = strPics & oRS("pic_name") & ","
	oRS.movenext
loop
oRS.Close
Set oRS = nothing

strPics = Left(strPics, Len(strPics) - 1)
Session("picList") = Split(strPics, ",")

%>

<frameset cols="*,100" border=0 framespacing=0 frameborder=0>
   <frame src="webAlbum_intro.asp?subcat=<%=strSubCat%>" NAME="pics">
   <frame src="webAlbum_index.asp?index=1&page=10" marginwidth=0 marginheight=0>
</frameset>