<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="webAlbum_data.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' updated 3/22/2000

if Request.Form("cancel") = "Cancel" then
	' if cancel was hit then send back to picture detail page
	response.redirect "webAlbum_detail.asp?index=" & Request.Form("index")
elseif Request.Form("skip") = "Skip" then
	' or if skip was hit (auto-add mode) then goto next new picture
	response.redirect "webAlbum_edit.asp?action=new&index=" & Request.Form("index") + 1
end if

'-------------------------------------------
' DATA UPDATE
'-------------------------------------------

dim x			' loop counter
dim oConn
dim oRS
dim strQuery
dim strCat
dim strSubCat
dim arPics
dim strName
dim intPicID
dim intIndex
		
intIndex = CInt(Request.Form("index"))
arPics = Session("picList")
strName = arPics(intIndex)

Set oConn = Server.CreateObject("ADODB.Connection")
transOpen(oConn)
Set oRS = Server.CreateObject("ADODB.Recordset")

if Request.Form("action") = "update" then
	if Request.Form("pic_name") <> strName then
		Call renameFile(Request.Form("pic_name"), strName)
		arPics(intIndex) = Request.Form("pic_name")
		Session("picList") = arPics
	end if

	strQuery = "SELECT * FROM pictures WHERE pic_name='" & strName & "'"
	oRS.Open strQuery, oConn, adOpenStatic, adLockOptimistic, adCmdText
	intPicID = picSave(oRS)
	
	Call catSave(oConn, oRS, intPicID, Request.Form("pic_cat"), true)
	Set oRS = nothing
	transClose(oConn)

	' send user back to detail view of updated picture detail
	response.redirect "webAlbum_detail.asp?index=" & Request.Form("index")

elseif Request.Form("action") = "new" then
	oRS.Open "pictures", oConn, adOpenStatic, adLockOptimistic, adCmdTable
	oRS.AddNew
	intPicID = picSave(oRS)
	Call catSave(oConn, oRS, intPicID, Request.Form("pic_cat"), false)
	Set oRS = nothing
	transClose(oConn)

	' send back to edit form for additional new pictures
	response.redirect "webAlbum_edit.asp?action=new&index=" & Request.Form("index") + 1
end if

' save picture information------------------------------------------------
function picSave(ByRef oRS)
	' not atomic--needs multiple form values
	oRS.Fields("pic_name") = Request.Form("pic_name")
	oRS.Fields("pic_description") = Request.Form("pic_description")
	oRS.Fields("pic_taken") = Request.Form("pic_taken")
	oRS.Fields("pic_location") = Replace(Request.Form("pic_location"), "'", "''")
	oRS.Fields("pic_source") = Request.Form("pic_source")
	oRS.Fields("pic_keywords") = Replace(Request.Form("pic_keywords"), "'", "''")
	oRS.Update
	picSave = oRS.Fields("pic_id")
	oRS.Close
end function

' save category information-----------------------------------------------
sub catSave(ByRef oConn, ByRef oRS, intPicID, strSubCat, bUpdate)
	if bUpdate then
		' clear old category assignments
		strQuery = "DELETE FROM picmix WHERE pic_id=" & intPicID
		oConn.Execute strQuery,,adCmdText
	end if
	
	oRS.CursorLocation = adUseClient
	oRS.Open "picmix", oConn, adOpenStatic, adLockBatchOptimistic, adCmdTable
	for each x in Split(strSubCat, ", ")
		oRS.AddNew
		oRS.Fields("pic_id") = intPicID
		oRS.Fields("subcat_id") = x
	next
	oRS.UpdateBatch
	oRS.Close
end sub

' rename physical picture file--------------------------------------------
sub renameFile(strNewName, strOldName)
	dim strBasePath
	dim strHiPath
	dim strLoPath
	dim oFS
	dim oLoFile
	dim oHiFile
	
	strBasePath = Server.Mappath(".\pictures")
	strHiPath = strBasePath & "\hi-res\"
	strLoPath = strBasePath & "\lo-res\"
	
	Set oFS = CreateObject("Scripting.FileSystemObject")
	Set oLoFile = oFS.GetFile(strLoPath & strOldName & "-lo.jpg")
	oLoFile.Name = strNewName & "-lo.jpg"
	Set oHiFile = oFS.GetFile(strHiPath & strOldName & ".jpg")
	oHiFile.Name = strNewName & ".jpg"
	
	Set oLoFile = nothing
	Set oHiFile = nothing
	Set oFS = nothing
end sub
%>