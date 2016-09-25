<%
' this bit of script allows you to set a color
' theme for webAlbum

' here's a key to the colors as presently used:
' 0 =             [lightest shade]
' 1 = background  .
' 2 =             . 
' 3 =             .
' 4 = title text, heading background
' 5 =             .
' 6 = active link [darkest shade]
' 7 = link        [begin irregular colors]

dim arColor
arColor = Array("ffb953","ff9933","df7913","bf5903","9f3902","7f1901","5f0900","fff993","ffd973","ff6600","000000","e0e0e0","c0c0c0","a0a0a0","808080")
%>

<META NAME="description" content="Troy High School class of 1991, Troy Idaho">
<META NAME="keywords" content="Troy High School, Class of 1991, Troy Idaho, Jason Abbott, James Aiken, Connie Anderson, Brett Bauer, Kelly Bingman, Randy Byers, James Christiansen, Julie Colemen, Joe Cox, Michael Davies, DeShane Frazier, Tod Geidl, Brian Glenn, Bill Hart, Tanya Hayden, Tiffany Heath, Mikko Kajava, Joel Kerley, Cynthia Kim, Jodie Knight, Susan Leslie, Mark Linderman, Sara Maki, Erik Marone, Ryan McCugh, Debbie McKenzie, Dena Neese, Tambi Oakley, Brian Pope, Tara Porter, Amanda Prewitt, Annette Rasmussen, Leslie Sattler, Amy Smith, Kara Something, Yolanda Something, Kim Traver, Amy Wallen, Tony Wallen, Darin Werlinger, Aaron Wood, Matt Young, Jadie Zimmerman">

<style type='text/css'>
	A:hover { color:#<%=arColor(7)%>; }
	A:link { text-decoration: none; }
	A:active { text-decoration: none; }
	A:visited { text-decoration: none; }
</style>