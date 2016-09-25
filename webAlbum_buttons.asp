<body bgcolor="#000000">
<center>
<form action="webAlbum_admin.asp" method="post" target="view">
<% if Session(strDataName & "User") = "" then %>
<input type="submit" value="Login">
<% else %>
<input type="submit" value="Admin">
<% end if %>
</form>
</center>
</body>