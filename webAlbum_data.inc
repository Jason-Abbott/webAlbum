<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' Last updated 3/22/2000

dim strDataName		' specify a unique name for each copy of webAlbum
dim strDSN			' connection string

strDataName = "thsAlbum"

' OLEDB connection string
strDSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" _
	& Server.Mappath("../../../database/" & strDataName & ".mdb")

' starts transaction-----------------------------------------------------
sub transOpen(ByRef oConn)
	oConn.Open strDSN
	oConn.BeginTrans
end sub

' ends transaction--------------------------------------------------------
sub transClose(ByRef oConn)
	if oConn.Errors.Count = 0 AND Err.Number = 0 then
		oConn.CommitTrans
	else
		if oConn.Errors.Count > 0 then
			response.write "Errors encountered: "
			for x = 0 to oConn.Errors.Count - 1
				response.write oConn.Errors(x).Description & "<br>"
			next
		end if
		if Err.Number > 0 then
			' this will only return the most recent error
			response.write Err.Source & " error " & Err.Number & ": " _
				& Err.Description
		end if
		oConn.RollbackTrans
	end if
	oConn.Close
	Set oConn = nothing
end sub
	
' ADO Constants-----------------------------------------------------------
	
' cursors
Const adOpenForwardOnly = 0
Const adOpenKeyset = 1
Const adOpenDynamic = 2
Const adOpenStatic = 3

' cursor location
Const adUseServer = 2
Const adUseClient = 3

' locks
Const adLockReadOnly = 1
Const adLockPessimistic = 2
Const adLockOptimistic = 3
Const adLockBatchOptimistic = 4

' commands
Const adCmdUnknown = &H0008
Const adCmdText = &H0001
Const adCmdTable = &H0002
Const adCmdStoredProc = &H0004
Const adCmdFile = &H0100
Const adCmdTableDirect = &H0200
%>