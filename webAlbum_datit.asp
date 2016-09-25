<% Option Explicit %>
<% Response.Buffer = True %>
<!--#include file="data/webAlbum_data.inc"-->
<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' updated 3/22/2000

dim strQuery
dim oRS
dim oConn
dim strName
dim intPos
dim intGrade

Set oConn = Server.CreateObject("ADODB.Connection")
oConn.Open strDSN
oConn.BeginTrans

Set oRS = Server.CreateObject("ADODB.RecordSet")
strQuery = "SELECT * FROM pictures P INNER JOIN picmix M " _
	& "ON P.pic_id = M.pic_id"
oRS.CursorLocation = adUseClient
oRS.Open strQuery, oConn, adOpenForwardOnly, adLockBatchOptimistic, adCmdText

do while not oRS.EOF
	strName = oRS("pic_name")
	response.write strName
	if CInt(oRS("cat_id")) = 1 then
		intPos = InStr(strName,"_")
		intGrade = Right(strName,Len(strName) - intPos)
		if IsNumeric(intGrade) then
			if intGrade > 0 and intGrade <= 12 then
				select case intGrade
					case 1
						oRS.Fields("pic_taken") = "9/5/79"
					case 2
						oRS.Fields("pic_taken") = "9/5/80"
					case 3
						oRS.Fields("pic_taken") = "9/5/81"
					case 4
						oRS.Fields("pic_taken") = "9/5/82"
					case 5
						oRS.Fields("pic_taken") = "9/5/83"
					case 6
						oRS.Fields("pic_taken") = "9/5/84"
					case 7
						oRS.Fields("pic_taken") = "9/5/85"
					case 8
						oRS.Fields("pic_taken") = "9/5/86"
					case 9
						oRS.Fields("pic_taken") = "9/5/87"
					case 10
						oRS.Fields("pic_taken") = "9/5/88"
					case 11
						oRS.Fields("pic_taken") = "9/5/89"
					case 12
						oRS.Fields("pic_taken") = "9/5/90"
				end select
				response.write " (" & oRS("pic_taken") & ")"
			end if		
		elseif intGrade = "snr" then
			oRS.Fields("pic_taken") = "3/10/91"
' 		end if
	end if
	response.write "<br>"
	
	oRS.MoveNext
loop
oRS.UpdateBatch

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