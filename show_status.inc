<%
function showStatus(msg)
	msg = Replace(msg, "'", "\'")
	showStatus = "onMouseOver=""status='" & msg & "'; " _
		& "return true;"" " & VbCrLf _
		& "onMouseOut=""status=''; return true;"""
end function
%>