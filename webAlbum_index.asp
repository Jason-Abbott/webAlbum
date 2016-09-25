<% Option Explicit %>
<% Response.Buffer = False %>
<!--#include file="show_status.inc"-->
<!--#include file="webAlbum_data.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' Last updated 3/28/2000

dim arPics			' array of picture data
dim intTotal		' total pics on this subcategory
dim intStart		' first pic shown on page
dim intEnd			' last pic shown on page
dim intPerPage		' number of thumbnails shown per page
dim strRange		' index range for this page
dim strChoices		' HTML to jump to other pages (number ranges)
dim strHtml
dim x				' loop counter

intStart = CInt(Request.QueryString("index"))
intPerPage = CInt(Request.QueryString("page"))
arPics = Session("picList")
intTotal = UBound(arPics) + 1

if intTotal > intPerPage then
	' there are more pictures than allowed per page--show page links
	strChoices = "<table cellspacing='0' cellpadding='2' border='0' class='PicIndex'>" _
		& "<tr><td class='PicHead'>Go to</td>" _
		& "<tr><td class='PicPages'>"
	x = 1
	' generate index range page links
	do until x > intTotal
		if x + intPerPage > intTotal then
			' this is the last page of pics
			intEnd = intTotal
			if x = intEnd then
				' only one link in this range
				strRange = intEnd
			else
				strRange = x & " &ndash; " & intEnd
			end if
		else
			intEnd = x + intPerPage - 1
			strRange = x & " &ndash; " & intEnd
		end if
		
		strChoices = strChoices & "<a href='webAlbum_index.asp?index=" _
			& x & "&page=" & intPerPage & "'>" & strRange & "</a><br>"
		x = x + intPerPage
	loop
	strChoices = strChoices & "</td></table>"

	if intStart + intPerPage > intTotal then
		intEnd = intTotal
		if intStart = intEnd then
			' only one picture on this page
			strRange = intEnd
		else
			strRange = intStart & " &ndash; " & intEnd
		end if
	else
		intEnd = intStart + intPerPage - 1
		strRange = intStart & " &ndash; " & intEnd
	end if
	
	strHtml = strHtml & strRange & "<br>" _
		& "of " & intTotal & " pictures<div class='AllPics'>" _
		& "<a href='webAlbum_index.asp?index=1&page=" _
		& intTotal & "'>Show All</a></div>"
else
	intEnd = intTotal
	strHtml = strHTml & intTotal & " pictures<br>"
	strChoices = ""
end if

for x = intStart - 1 to intEnd - 1
	strHtml = strHtml & "<a href='webAlbum_detail.asp?index=" _
		& x & "' target='pics'" & ShowStatus("Click to enlarge and see description") _
		& "><img src='./pictures/lo-res/" & arPics(x) & "-lo.jpg' border='1' width='70' class='thumb'></a><br>"
next
%>
<html>
<head>
<link rel='stylesheet' href='../style/ths1991.css' type='text/css'>
</head>
<body class='PicIndex'>
<center>
<%=strHtml%>
<%=strChoices%>
</font>
</center>
</body>
</html>