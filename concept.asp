<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include virtual="/db/database.asp"-->		
<!--#include virtual="/inc/encryption.asp"-->	
<!--#include virtual="/inc/common200408.asp"-->		
<!--#include virtual="/inc/questions200408.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
	<title>Concept</title>
	<LINK REL=stylesheet HREF="/css/survey.css" TYPE="text/css"> <!-- Calls style sheet.  -->
</head>

<body>
<div class="concept">
<% GETCONCEPT (Request.QueryString("cid"))%>
<form><input type="button" value="close" onClick="self.close()" class="conceptClose"></form> 
</div>
</body>
</html>
