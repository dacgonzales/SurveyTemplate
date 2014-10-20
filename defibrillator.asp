<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<META HTTP-EQUIV="Pragma" CONTENT="no-cache">
		<meta http-equiv="refresh" content="<%=(Session.Timeout-5)%>;">
	</HEAD>
	<body>
			<% Response.Write Now() %>
			<% Response.CacheControl = "no-cache" %>
			<% Response.AddHeader "Pragma", "no-cache" %>
			<% Response.Expires = -1 %>
	</body>
</HTML>
