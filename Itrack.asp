<%@ Language=VBScript %>
<!--#include virtual="/surveys/inc/adovbs.inc"-->
<!--#include virtual="/db/database.asp"-->	
<!--#include file="inc/common.asp"-->
<%
Dim pssWrd

callbackUrl = "http://surveys.datassential.com/surveys/sur_20140907_ehfiPeCrre14/Itrack.asp"
indexUrl = "http://surveys.datassential.com/surveys/sur_20140907_ehfiPeCrre14/index.asp"


if Request.QueryString("assetgroupId") <> "" then
    'FOR FUTURE
    'expect surveyId

    assetGroupId = request.QueryString("assetgroupId")
    respId = request.QueryString("respid")    
    
    uCode = safeADODBexecuteReturn("SELECT uCode from Operator_segment where email = '" & respId & "'")
    callbackUrl = callbackUrl & "?uCode="&uCode    

    Response.Redirect("http://imarkit.itracks.com/API/V5/ViewDocument.aspx?assetGroupId=5257&ExternalUserId=" & uCode & "&nextPage=" & callbackUrl)
else
    'STEPS
    '1. Get lastPage of respodents based on uCode
    '2. Encrypt last page.
    '3. Append encrypted page to request.
    '4. Send http request to index page

    uCode = request.QueryString("uCode")
    surveycur = CInt(safeADODBexecuteReturn("select B.LastPage from Operator_segment A LEFT OUTER JOIN dbo.sur20140907sara_lee_chef_pierre B ON A.Email = B.resp_id where uCode = '" &uCode& "'")) + 6
    
    surveycur = EncryptString(surveycur)
    indexUrl = indexUrl & "?surveycur=" & surveycur

    Response.Write indexUrl
    Response.End()

    Response.Redirect(indexUrl)
    
end if
    
%>
