<%
' Copyright 2000 Jason Abbott (jason@webott.com)
' Last updated 3/22/2000

dim intCat		' category id

' select case Request.QueryString("cat")
' 	case 1
' 		Session("color") = Array("c8fb62","a8db42","88bb22","689b02","487b02","285b02","083b02","660066","ffcc33","ffffff","c6c6d9","e0e0e0","c0c0c0","a0a0a0","808080")
' 	case 2
' 		Session("color") = Array("ffd973","ffb953","ff9933","df7913","bf5903","9f3903","7f1903","660066","ffcc33","ffffff","c6c6d9","e0e0e0","c0c0c0","a0a0a0","808080")
' 	case 3
' 		Session("color") = Array("ffff86","ffff66","dfdf46","bfbf26","9f9f06","7f7f03","5f5f00","660099","ffcc33","ffffff","c6c6d9","e0e0e0","c0c0c0","a0a0a0","808080")
' 	case else
' 		Session("color") = Array("c8fb62","a8db42","88bb22","689b02","487b02","285b02","083b02","660066","ffcc33","ffffff","c6c6d9","e0e0e0","c0c0c0","a0a0a0","808080")
' end select

Session("color") = Array("ffb953","ff9933","df7913","bf5903","9f3902","7f1901","5f0900","fff993","ffd973","ff6600","000000","e0e0e0","c0c0c0","a0a0a0","808080")

' here's a key to the colors as presently used:
' 0 =             [lightest shade]
' 1 = background  .
' 2 =             . 
' 3 =             .
' 4 = title text, heading background
' 5 =             .
' 6 = active link [darkest shade]
' 7 = link        [begin irregular colors]

intCat = Request.QueryString("cat")

%>

<frameset cols="100,*" frameborder="0" border=0 framespacing=0 border=0>
	<frameset rows="*,40" frameborder="0" border=0 framespacing=0 border=0>
		<frame name="select" src="webAlbum_select.asp?cat=<%=intCat%>" marginwidth=0 marginheight=0 scrolling="no" frameborder="0" noresize>
		<frame src="webAlbum_buttons.asp" marginwidth=0 marginheight=0 scrolling="no" frameborder=0 noresize>
	</frameset>
	<frame name="view" src="webAlbum_intro.asp?cat=<%=intCat%>">
</frameset>
