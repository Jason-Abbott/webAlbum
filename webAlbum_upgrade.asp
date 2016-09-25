<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="data/webAlbum_data.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' updated 3/22/2000

dim strQuery
dim oRS
dim oRS2
dim oConn
dim strCats
dim arCats
dim strSubCats
dim arSubCats
dim x

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open strDSN
oConn.BeginTrans

Set oRS = Server.CreateObject("ADODB.RecordSet")
strQuery = "SELECT pic_id, pic_cat, pic_subcat FROM pictures"
oRS.Open strQuery, oConn, adOpenForwardOnly, adLockReadOnly, adCmdText

Set oRS2 = Server.CreateObject("ADODB.RecordSet")
oRS2.CursorLocation = adUseClient
oRS2.Open "picmix", oConn, adOpenStatic, adLockBatchOptimistic, adCmdTable

do while not oRS.EOF
	response.write "updating " & oRS("pic_id") & "<br>"
	strCats = Replace(oRS("pic_cat"),",,"," ")
	strCats = Replace(strCats,",","")
	strSubCats = Replace(oRS("pic_subcat"),",,"," ")
	strSubCats = Replace(strSubCats,",","")
	
	arCats = Split(strCats)
	arSubCats = Split(strSubCats)
	
	for x = 0 to UBound(arCats)
		oRS2.AddNew
		oRS2.Fields("pic_id") = oRS("pic_id")
		oRS2.Fields("subcat_id") = arSubCats(x)
		oRS2.Fields("cat_id") = arCats(x)
	next
	
	oRS.MoveNext
loop
oRS2.UpdateBatch

if oConn.Errors.Count = 0 AND Err.Number = 0 then
	oConn.CommitTrans
else
	oConn.RollbackTrans
end if

oRS.Close
Set oRS = nothing
oConn.Close
Set oConn = nothing

%>