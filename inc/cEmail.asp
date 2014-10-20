<!-- 
METADATA 
TYPE="typelib" 
UUID="CD000000-8B95-11D1-82DB-00C04FB1625D" 
NAME="CDO for Windows 2000 Library" 
-->
<%
'#######################################
'	Filename: userSurvey.asp
'	Author: Rex B. Isles, MCP
'	Owner: Datassential Research
'	Created: 13 August 2006
'
'	Purpose: 
'#######################################


Sub SendEmail(email_to, email_subject, email_Msg )

	' configure the CDO object
	dim bDebugMode
	bDebugMode = false
	
	Set cdoConfig = CreateObject("CDO.Configuration")
    Set Flds = cdoConfig.Fields

	'With cdoConfig.Fields 
	'  .Item(cdoSendUsingMethod)  = cdoSendUsingPort 
	'  .Item(cdoSMTPServer)       = "smpt.gmail.com"
	'  .Item(cdoSMTPAuthenticate) = cdoBasic
	'  .Item(cdoSendUserName)     = "panel@datassential.com"
	'  .Item(cdoSendPassword)     = "4jacumba"
	'  .Update 
	'End With

	schema = "http://schemas.microsoft.com/cdo/configuration/"
	Flds.Item(schema & "sendusing") = 2
	Flds.Item(schema & "smtpserver") = "smtp.gmail.com" 
	Flds.Item(schema & "smtpserverport") = 465 
	Flds.Item(schema & "smtpauthenticate") = 1
	Flds.Item(schema & "sendusername") = "panel@datassential.com"
	Flds.Item(schema & "sendpassword") =  "4jacumba"
	Flds.Item(schema & "smtpusessl") = 1
	Flds.Update

	Dim myMail

	Set myMail = CreateObject("CDO.Message") 
'	email_bcc = "rex@datassential.com;lorenzo@datassential.com;oliver@datassential.com;marylouise@datassential.com;emily@datassential.com;hooman@datassential.com"
	'email_bcc = "programmer@datassential.com;emily@datassential.com;hooman@datassential.com"
	'email_bcc = "programmer@datassential.com;David.gonzales@datassential.com;emily@datassential.com;deron@datassential.com;Wasilewski-Katie@aramark.com;ron.appin@gmail.com"
    'email_bcc = "marc.razal@datassential.com;David.gonzales@datassential;marylouise@datassential.com"
    email_bcc = "marc.razal@datassential.com;David.gonzales@datassential.com;hooman@datassential.com"
    email_cc = ""
	if email_to = "" Then
		'email_to = email_bcc
        email_to = "marc.razal@datassential.com"
		'email_bcc = ""
	end if
		
'	email_bcc = "lorenzo@datassential.com"
	myMail.Configuration = cdoConfig
	myMail.From = "panel@datassential.com"

'	if bDebugMode Then
'		email_subject = "TEST EMAIL!  " & email_subject
'		email_Msg = replace(email_Msg,"</body>","<br><br><div>[recipients]</div></body>")
'		email_Msg = replace(email_Msg,"[recipients]",email_to)
'		email_to = "lorenzo@datassential.com"
'	end if

	if bDebugMode then
        myMail.To = "marc.razal@datassential.com" 'email_to
    else
        myMail.To = email_to
    end if
	myMail.bcc =  email_bcc
	
	myMail.Subject = email_subject
    if bDebugMode then
	    myMail.cc = "" & email_to
    else
        'myMail.cc = "hooman@datassential.com"
        'myMail.cc = email_to
        myMail.cc = email_cc
    end if
	myMail.HTMLBody = email_Msg

	call myMail.Send()

	set myMail = nothing
	set cdoConfig = nothing
End Sub
%>
