<%
'Detect Client Browser Environment
Dim strUserAgent 
DIM Browser_MSIE, Browser_NS, Browser_OPERA, Browser_FIREFOX
Dim user_browser
' Browser Boolean Constants
user_browser = "IE"
strUserAgent = UCASE(cstr(request.ServerVariables("HTTP_USER_AGENT")))
'Explorer
If InStr(strUserAgent, "MSIE") > 0 Then
    user_browser = "IE"
End If
 
'Opera
If InStr(strUserAgent, "OPERA") > 0 Then
    user_browser = "OPERA"
End If
 
'Netscape
If InStr(strUserAgent, "NETSCAPE") > 0 Or InStr(strUserAgent, "GECKO") > 0 Then 
    user_browser = "NS"
End If

'Firefox

If InStr(strUserAgent, "FIREFOX") > 0 Then
    str_browser = "Firefox"
End If
'Response.Write(str_browser)
Function GetCompatibleCSS()
Dim str_css
	str_css = ""
	Select Case str_browser
		Case "Firefox"
			str_css = "survey_firefox.css"
		Case Else
			str_css = "survey.css"			
	End Select
	GetCompatibleCSS = str_css
End Function
%>