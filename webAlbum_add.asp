<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="webAlbum_data.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' updated 3/22/2000

dim strQuery
dim oRS
dim oFS
dim allTotal
dim hiName
dim strNames
dim strPath
dim oFolder
dim oFiles
dim x

' get the path to the pictures folder
strPath = Server.Mappath(".\pictures\hi-res")
strNames = " "

' connect to the pictures folder
Set oFS = CreateObject("Scripting.FileSystemObject")
Set oFolder = oFS.GetFolder(strPath)
Set oFiles = oFolder.Files

' cycle through each picture, adding all *.jpg files to
' a list of names
for each x in oFiles
	if Right(x.name, 4) = ".jpg" then
		strNames = strNames & Left(x.name, Len(x.name) - 4) & ", "
	end if
next

Set oFS = nothing
Set oFolder = nothing
Set oFiles = nothing

' if we ended up with some names then weed out all the
' ones already in the database
	
if Trim(strNames) <> "" then
	strQuery = "SELECT pic_name FROM pictures"
	Set oRS = Server.CreateObject("ADODB.RecordSet")
	oRS.Open strQuery, strDSN, adOpenForwardOnly, adLockReadOnly, adCmdText

	' if a match is found, remove the name from the list
	do while not oRS.EOF
		if Instr(strNames, " " & oRS("pic_name") & ", ") <> 0 then
			strNames = Replace(strNames, oRS("pic_name") & ", ", "")
		end if
		oRS.movenext
	loop
	oRS.Close
	Set oRS = nothing

	' chop off the trailing ", "
	strNames = Left(strNames, Len(strNames) - 2)

	' any names that survived the comparison are for new files
	if strNames <> "" then
		' save the list into a Session array
		Session("picList") = Split(strNames, ", ")
		response.redirect "webAlbum_edit.asp?action=new&index=0"
	end if
end if
%>

<body>
No new pictures were found.
</body>
