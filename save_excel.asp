<%@ Language=VBScript %>
<!--#include virtual="/surveys/inc/adovbs.inc"-->
<!--#include virtual="/db/database.asp"-->	
<!--#include file="inc/common.asp"-->
<%
Dim pssWrd
pssWrd = request.Form("expPass")
pssWrd = trim(pssWrd)
if pssWrd <> "p5datass3ntial2645" then
    response.Redirect("exportAuth.asp?qryErr=1")
else
    Response.Clear
    'Response.Charset = "ANSI"
    Response.ContentType = "application/octet-stream"
    Response.ContentType = "application/vnd.ms-excel"
    Response.AddHeader "Content-Disposition", "attachment; filename=emailaddress_List.xls"
    'Response.ContentType = "application/download"
    set con = safeADODBcon()
    'set rs = con.execute("lsp_get_email_status '" & Session("tableName") & "',1") 'if with the "qouta" column
    set rs = con.execute("lsp_get_email_status '" & Session("tableName") & "'")
    if not rs.eof then
        response.write "<table border=1>"
        response.Write "<tr><td>id</td><td>emailAddress</td><td>status</td></tr>"
        while not rs.eof
            response.Write "<tr>"
            response.write "<td>" & rs("id") & "</td><td>" & rs("resp_id") & "</td><td>" & rs("status") & "</td>"
            response.Write "</tr>"
            rs.movenext
        wend
        response.write "</table>"
    end if
    set rs = nothing
    set con = nothing
end if
%>
