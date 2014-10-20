<!--#include virtual="/surveys/inc/adovbs.inc"-->
<!--#include virtual="/db/database.asp"-->	
<!--#include file="inc/common.asp"-->
<%

' Get the Session Variables
surveyID = getSurveyTable()
'getSurveyVariables(surveyID)
getSurveyVariables2(surveyID) 
'getSurveyVariables2(251)


'-------------------------< Settings >-----------------------------------
' 2008 - bernard
sur_table = Session("tableName")
que_table = Session("qTableName")
' since SurveyType is not saved in Session we just manually set it here
if cInt(Session("surveyType")) = 1 Then
	is_operator = true
else
	is_operator = false
end if
'------------------------------------------------------------------------
%>

<html>
<head>
<title>Datassential Research :</title>
<meta http-equiv="refresh" content="30" />
<link rel="shortcut icon" href="images/datassential.gif">
<style type="text/css">	 
	body {background-color: #416578; }
	th {background-color: #ffffff; font-family: arial; font-size: 14px;}
	a {color: #ffffff ;text-decoration:none}
	a:link { text-decoration:none}
	a:vlink { text-decoration:none}
	a:hover { text-decoration:underline}
	.tblstyle1 {border-style: solid; border-color: #ffaabb; border-width: medium}	
	.tblstyle {border-style: solid; border-color: #ffffff; border-width: medium}	
	.divstyle {border: 1px solid #004D85; background-color: #ffffff;width:635}
	.spanstyle  {font-family: arial; font-size: 16px; font-weight: bolder; color: #000000}		
	.tdstyle	{background-color:#006699 ;color:#FFFFFF; font-family: arial; font-size: 12px; padding-left: 10px; padding-top: 5px; padding-bottom: 5px}    
    .tdresult {background-color:#006699 ;color:#FFFFFF; text-align:center  ;padding:4px; font-family: arial; font-size: 12px; width: 25%}
    .tdresulttotal {background-color:#006699 ;color:#FFFFFF; text-align:center ;padding:4px; font-family: arial; font-size: 15px; width: 25%}
    .tdsubresult {background-color:#006699 ;color:#FFFFFF; text-align:center ;padding:4px; font-family: arial; font-size: 12px; width: 25%}
 </style>
</head>

<!-- HTML CODE -->
<body>
<center>
<div align="center" class="divstyle">
	<div align="center" id="header" ><img src="images/headerbarlogo.jpg" width="635" height="58"></div><br>
	
	<span class="spanstyle"><%=Session("SurveyName")%></span><br><br>
			
		<table  width="635">
			<tr>
				<td class="tdstyle">Completed</td>
				<%DisplayData RetrieveTotalCompleted%>
			</tr>

			<tr>
				<td class="tdstyle">Disqualified</td>
				<%DisplayData RetrieveTotalDisqualified%>
			</tr>
			<tr>
				<td class="tdstyle">Discontinued</td>
				<%DisplayData RetrieveTotalDiscontinued%>
			</tr>
			<!--<tr>
				<td class="tdstyle">Quota reached</td>
			</tr>-->
			<tr>
				<td class="tdstyle">Total Records</td>   
				<%DisplayData RetrieveTotalRecords%>
			</tr>
            <tr>
                <td class="tdstyle">Export completes / disqualified / discontinued</td>
                <td class="tdresult">
                    <a href="exportAuth.asp" target="_blank">Export</a>
                </td>
            </tr>
		</table>
		<%	'------------------------------------------
			'Specify the var_name for the segment question
			
			%>
			
</div>
</center>  
</body>
</html>

<!-- Functions  -->
<%
Function RetrieveTotalQuota()
	dim sql
	sql = ""
	sql = sql & " select count(*) "
	sql = sql & " from  " & sur_table
	sql = sql & " (NOLOCK) where timestart is not null and quota is not null and timestop is null"
	if is_operator Then 
		'sql = sql & " AND resp_id NOT IN ( select resp_id FROM " & sur_table
		'sql = sql & " WHERE timestart IS NOT NULL AND timestop IS NOT NULL ) "
	End If
	Set RetrieveTotalQuota = safeADODBcon.execute(sql)
End Function
Function DisplayData(rs)
    On Error Resume Next
	  Do Until rs.EOF 
	  	For each x in rs.Fields
		    If NOT IsNull(rs(1).name) Then 		    		    		
				    Response.write("<td class=""tdresult""><strong>" & rs(0).value & "</strong></td>")
		   	End If
	   	Next
	  	rs.MoveNext
	  Loop
End Function

Function RetrieveTotalCompleted()
	dim sql
	sql = ""
	sql = sql & " select count(*) "
	sql = sql & " from  " & sur_table
	sql = sql & " (NOLOCK) where timestart is not null and timestop is not null"
	'response.Write sql
	'response.End
	Set RetrieveTotalCompleted = safeADODBcon.execute(sql)
End Function

Function RetrieveTotalDisqualified()
	dim sql
	sql = ""
	sql = sql & " select count(*) "
	sql = sql & " from  " & sur_table
	sql = sql & " (NOLOCK) where timestart is not null and timescreen is not null"
	if is_operator Then 
		'sql = sql & " AND resp_id NOT IN ( select resp_id FROM " & sur_table
		'sql = sql & " WHERE timestart IS NOT NULL AND timestop IS NOT NULL ) "
	End If
	Set RetrieveTotalDisqualified = safeADODBcon.execute(sql)
End Function

Function RetrieveTotalDiscontinued()
	dim sql
	sql = " SELECT COUNT (*) FROM " & sur_table
	'sql = sql & " WHERE timestart IS NOT NULL AND (timestop IS NULL AND timescreen IS NULL AND quota IS NULL) " 'with quota column
    sql = sql & " (NOLOCK) WHERE timestart IS NOT NULL AND (timestop IS NULL AND timescreen IS NULL) "
	if is_operator Then 
		'sql = sql & " AND resp_id NOT IN ( select resp_id FROM " & sur_table
		'sql = sql & " WHERE timestart IS NOT NULL AND (timestop IS NOT NULL OR timescreen IS NOT NULL) ) "
	End If
	Set RetrieveTotalDiscontinued = safeADODBcon.execute(sql)
End Function

Function RetrieveTotalRecords()
	dim sql
	sql = ""
	sql = sql & " select count(*) "
	sql = sql & " from " &  sur_table
	sql = sql & " (NOLOCK) where timestart is not null"
	Set RetrieveTotalRecords = safeADODBcon.execute(sql)
End Function

Function getSurveyVariables2(surveyID)
	Session("surveyName") =""
	Session("status") =""
	Session("surveyType") =""
	Session("TableName")=""
	Set con = safeADODBcon()	
	Set rs = con.Execute("SELECT * FROM surMaster (NOLOCK) WHERE id = " & surveyID)
	If rs.EOF Then
		'Error
	Else
		Session("surveyName") = rs("surveyName")
		Session("status") = rs("status")
		Session("surveyType") = rs("surveyType")
		Session("TableName") = rs("TableName")
		Session("qTableName") = rs("qTableName")
	End If
	rs.Close
	Set rs = Nothing
	safeADODBclose()
	
End Function 

Sub RetrieveSegmentCounts(var_name)
	dim sql, sQText, sOption, rs, ndx, aryOptions, bIsOther
	sql = "Select top 1 q_text,options,is_other from " & que_table & " (NOLOCK) where var_name = '" & var_name & "'"
	set rs = safeADODBcon.execute(sql)
	If rs.EOF Then
		exit sub
	End If
	bIsOther = rs("is_other")
	sQText = rs("q_text")
	aryOptions = split(rs("options"),"|")
	sql = ""
	for ndx = lBound(aryOptions) to uBound(aryOptions)
		sql = sql & " WHEN " & cStr(ndx+1) & " THEN '" & aryOptions(ndx) & "' "
	next
	if bIsOther Then
		sql = sql & " WHEN " & cStr(uBound(aryOptions)+2) & " THEN 'Other' "
	End If
	sql = "SELECT (CASE [" & var_name & "] " & sql
	sql = sql & " END )"
	sql = sql & " AS [" & var_name & "],count([" & var_name & "]) AS [ctr]  FROM " & sur_table
	sql = sql & " (NOLOCK) WHERE timestart IS NOT NULL AND timestop IS NOT NULL AND timescreen IS NULL"
	sql = sql & " AND [" & var_name & "] IS NOT NULL "
	sql = sql & " GROUP BY [" & var_name & "] with rollup"
	set rs = safeADODBcon.execute(sql)
	If rs.EOF Then
		exit sub
	End If
%>
		<table width="635" ID="Table1">
			<tr>
				<td  colspan="2" class="tdstyle" align="center">
					<strong>Total Respondents per Segment <br>
					(<%=var_name%> - <%=sQText%>)</strong>
				</td>
			</tr>
		<%while not rs.EOF
			if isnull(rs(var_name)) then
				sOption = " - TOTAL COMPLETED - "
			else
				sOption = rs(var_name)
			end if%>
			<tr>
				<td class="tdstyle"><strong><%=sOption%></strong></td>
				<td class="tdresult"><strong><%=rs("ctr")%></strong></td>
			</tr>
		<%	rs.movenext
		  wend%>
		</table>
<%
End Sub
%>