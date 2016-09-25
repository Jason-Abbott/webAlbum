<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="webAlbum_data.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' Last updated 3/22/2000

dim oRS
dim arCats
dim strQuery
dim events
dim dateList
dim strCombine
dim strPics
dim s
dim strField
dim arWords
dim strWord
dim intPerPage

strCombine = ""
strQuery = ""

' this subroutine converts the strField information into
' proper SQL

sub parse(strField, s)
	strQuery = strQuery & strCombine & "("
	
	' if it begins and ends with a quote
	' use Mid to cut out what is between the quotes
	if Left(s,1) = """" and Right(s,1) = """" then
		strQuery = strQuery & strField & " LIKE '%" _
			& Mid(s, 2, Len(s) - 2) & "%'"

	' or if it has a plus (+) in it
	elseif InStr(s,"+") > 0 then
		arWords = Split(s, "+")
		for strWord = 0 to UBound(arWords)
			strQuery = strQuery & strField & " LIKE '%" & arWords(strWord) & "%'"
			if strWord < UBound(arWords) then
				strQuery = strQuery & " AND "
			end if
		next
		
	' or if it starts with a minus (-)
	elseif Left(s,1) = "-" then
		strQuery = strQuery & strField & " NOT LIKE '%" _
			& Right(s, Len(s)-1) & "%'"
	
	' otherwise split on spaces
	else
		arWords = Split(s)
		for strWord = 0 to UBound(arWords)
			strQuery = strQuery & strField & " LIKE '%" & arWords(strWord) & "%'"
			if strWord < UBound(arWords) then
				strQuery = strQuery & " OR "
			end if
		next
	end if
	
	strQuery = strQuery & ")"
end sub	

' now go through all the form elements to figure out which 
' ones have values to generate concise SQL

if Request.Form("pic_keywords") <> "" then
	strQuery = strQuery & strCombine
	Call parse("pic_keywords", Request.Form("pic_keywords"))
	strCombine = " AND "
end if

if Request.Form("pic_description") <> "" then
	strQuery = strQuery & strCombine
	Call parse("pic_description", Request.Form("pic_description"))
	strCombine = " AND "
end if

if Request.Form("pic_cat") <> "all" then
	strQuery = strQuery & strCombine
	arCats = split(Request.Form("pic_cat"), "|")
	if arCats(0) then
		strQuery = strQuery & "(subcat_id=" _
			& arCats(1) & ")"
	else
		strQuery = strQuery & "(cat_id=" _
			& arCats(1) & ")"
	end if
	strCombine = " AND "	
end if

' if a start or end date was entered, but not both, then
' use the current date as the missing date

if Request.Form("date_start") <> "" OR Request.Form("date_end") <> "" then
	strQuery = strQuery & strCombine & "(pic_taken BETWEEN #"
	if Request.Form("date_start") <> "" then
		strQuery = strQuery & Request.Form("date_start")
	else
		strQuery = strQuery & Date
	end if
	strQuery = strQuery & "# AND #"
	if Request.Form("date_end") <> "" then
		strQuery = strQuery & Request.Form("date_end")
	else
		strQuery = strQuery & Date
	end if
	strQuery = strQuery & "#)"
	strCombine = " AND "
end if

' now build the full query and run it

strQuery = "SELECT pic_name, pic_taken " _
	& "FROM pictures P LEFT JOIN picmix M ON P.pic_id = M.pic_id " _
	& "WHERE " & strQuery & " ORDER BY pic_taken"

Set oRS = Server.CreateObject("ADODB.RecordSet")
oRS.Open strQuery, strDSN, adOpenForwardOnly, adLockReadOnly, adCmdText

if oRS.EOF OR oRS.BOF then response.redirect "webAlbum_find.asp?retry=1"

do while not oRS.EOF
	if InStr(strPics,oRS("pic_name") & ",") = 0 then
		' only add unique pictures to the list
		strPics = strPics & oRS("pic_name") & ","
	end if
	oRS.movenext
loop
oRS.Close
Set oRS = nothing

strPics = Left(strPics, Len(strPics) - 1)
Session("picList") = Split(strPics, ",")
intPerPage = CInt(Request.Form("per_page"))

if UBound(Session("picList")) + 1 < intPerPage then
	intPerPage = UBound(Session("picList")) + 1
end if
%>

<frameset cols="*,100" border=0 framespacing=0 frameborder=0>
   <frame src="webAlbum_find.asp" NAME="pics">
   <frame src="webAlbum_index.asp?index=1&page=<%=intPerPage%>" marginwidth=0 marginheight=0>
</frameset>

