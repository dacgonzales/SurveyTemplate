<!--#include file="hex_md5_js.asp"-->
<!--#include file="claim.asp"-->
<!--#include file="cEmail.asp"-->
<%
ADODB_CON = null
Function safeVar( var ) 
	If Request.QueryString(var) = ""  Then
		safeVar = Request.Form(var)
	Else
		safeVar = Request.QueryString(var)	
	End If
End Function

Function safeSql( var )
	safeSql = ("'" & Replace(var,"'","''") & "'" )
End Function

Function safeSurveyCur()
	If DeCryptString(request.QueryString("surveycur")) <> "" Then
		safeSurveyCur = DeCryptString(request.QueryString("surveycur"))
	Else
		safeSurveyCur = 0
	End If
End Function

Function isNotDone() 
	'/ not: finish, error, etc
	isNotDone = (Len(safeSurveyCur()) < 4)
End Function

Function safeUid()
	If Session("uid") = "" Then
		safeUid = 0
	Else
		'safeUid = CInt(DeCryptString(Session("uid")))	
		safeUid = Session("uid")
	End If 
End Function

Function doneSession()
	Session("uid") = ""
	Session("surveyName") = ""
	Session("surveyLocation") = ""
	Session("tableName") = ""
	Session("qTableName") = ""
	Session("totalcompletes") = ""
	Session("totalquickserve6") = ""
	Session("totalmidscale3") = ""
	Session("totalcasual2") = ""
	Session("totalfine1") = ""
	Session("totalbusiness18") = ""
	Session("totalhealthcare13") = ""
	Session("totallodging12") = ""
	Session("totalschools16") = ""
	Session("totalcolleges17") = ""
	Session("totalother23") = ""
	Session("totaltesting100") = ""
	Session("is_panel") = ""
	Session("is_gozing") = ""
End Function

Function safeADODBcon()
	If IsNull(ADODB_CON) Then
		Set ADODB_CON = Server.CreateObject("ADODB.Connection")
	End If
	'/ open
	If ADODB_CON.State = 0 Then
		'/ ADODB_CON.ConnectionTimeout = 3000
		ADODB_CON.ConnectionTimeout = 0 '/ wait indefinitely
		ADODB_CON.Open(connStr)
	End If	
	Set safeADODBcon = ADODB_CON
End Function

Function safeADODBclose()
	On Error Resume Next
	If Not IsNull(ADODB_CON) Then
		ADODB_CON.Close()
		Set ADODB_CON = Nothing
		ADODB_CON = Null
	End If
End Function


Function safeADODBexecute( sql_str  )
	Set con = safeADODBcon()
	'response.Write("Session(options2split_11): " & Session("options2split_11") & "<br>")
	'response.Write(sql_str & "<br>")
	'response.End()
	con.Execute(sql_str)
	safeADODBclose()
End Function


Function safeADODBexecuteReturn( sql_str  )
	Set con = safeADODBcon()
	Set RS = con.Execute(sql_str)
	If NOT RS.EOF Then
		safeADODBexecuteReturn = RS(0)
	Else
		safeADODBexecuteReturn = ""
	End If
	safeADODBclose()
End Function

Function getSurveyTable()
	Dim fso, sourceFolder
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set targetFolder = fso.GetFile(Server.MapPath(Request.ServerVariables("PATH_INFO")))
	tableName = Replace(targetFolder.ParentFolder.Name,"_","")
	tableName = "que20140807testsurvey2"
	sql_str = "SELECT id FROM surMaster (NOLOCK) WHERE qTableName LIKE '" & tableName & "%'"
	getSurveyTable = safeADODBexecuteReturn( sql_str  )
End Function

Function getSurveyVariables(surveyID)
	Set con = safeADODBcon()	
	Set rs = con.Execute("SELECT * FROM surMaster (NOLOCK) WHERE id = " & surveyID)
	If rs.EOF Then
		'Error
	Else
		Session("surveyName") = rs("surveyName")
		'Session("surveyLocation") = rs("surveyLocation")
		Session("tableName") = rs("tableName")
		Session("qTableName") = rs("qTableName")
		Session("totalcompletes") = rs("totalcompletes")
		Session("totalquickserve6") = rs("totalquickserve6")
		Session("totalmidscale3") = rs("totalmidscale3")
		Session("totalcasual2") = rs("totalcasual2")
		Session("totalfine1") = rs("totalfine1")
		Session("totalbusiness18") = rs("totalbusiness18")
		Session("totalhealthcare13") = rs("totalhealthcare13")
		Session("totallodging12") = rs("totallodging12")
		Session("totalschools16") = rs("totalschools16")
		Session("totalcolleges17") = rs("totalcolleges17")
		Session("totalother23") = rs("totalother23")
		Session("totalconvenience60") = rs("totalconvenience60")
		Session("totaltesting100") = rs("totaltesting100")
		Session("customTotalStr") = rs("customTotalStr")
		Session("is_panel") = rs("is_panel")
		Session("is_gozing") = rs("is_gozing")
		Session("inquisite_id") = rs("inquisite_id")
		Session("emailSource") = rs("emailSource")		
	End If
	rs.Close
	Set rs = Nothing
	safeADODBclose()	
End Function

Function checkSurveyCompletes(emailAddress)
	columnID = ""
	If Session("is_panel") Then
		columnID = "resp_id"
	Else
		If Session("is_gozing") Then
			columnID = "ui"
		Else
			If Session("customTotalStr") <> "" Then
				columnID = Session("customTotalStr")
			Else
				If NOT isNull(Session("totalcompletes")) Then
					columnID = "normal"
				Else
					columnID = "none"
				End If
			End If
		End If
	End If		
	returnCode = 0
	Set con = safeADODBcon()	
	Select Case columnID
		Case "ui" ' goZing Survey
			returnCode = 1 '"continue"
			Set rs = con.Execute("SELECT count(distinct(ui)) AS 'TOTAL' FROM " & Session("tableName") & " (NOLOCK) WHERE timeStop IS NOT NULL AND len(ui) > 10")
			If rs("TOTAL") >= Session("totalcompletes") Then
				response.redirect("/gozing/end.asp") ' gozing fail page redirect
			End If	
			rs.Close
			Set rs = Nothing
		Case "resp_id" ' panel Survey
			segmentVar = 0
			Set rs = con.Execute("SELECT Actual_segment FROM Operator_segment (NOLOCK) WHERE Email = '" & emailAddress & "'")
			If rs.EOF Then
				panelMember = False
			Else
				segmentVar = rs("Actual_segment")
				Session("segmentVar") = segmentVar
				panelMember = True
			End If
			rs.Close
			Set rs = Nothing
			
			If panelMember = True Then
				CommandText = "SELECT x.Actual_segment, count(o.resp_id) as 'TOTAL' " & _
					" FROM " & Session("tableName") & " o, Operator_segment x " & _
					" (NOLOCK) WHERE o.resp_id = x.Email  " & _
					" AND o.timeStop IS NOT NULL  " & _
					" AND x.Actual_segment = " & segmentVar &_
					" Group By x.Actual_segment " ' checks for number of completes in this segment
				Set rs = con.Execute(CommandText)
				totVar = 0
				If NOT rs.EOF Then
					totVar = rs("TOTAL")
				End If
				rs.Close
				Set rs = Nothing
				Set rs = con.Execute("SELECT TOP 1 id FROM " & Session("tableName") & " (NOLOCK) WHERE resp_id = '" & emailAddress & "' AND (timeStop IS NOT NULL OR timescreen IS NOT NULL)")
				If rs.EOF then
					alreadyTaken = False
				Else
					alreadyTaken = True
				End If
				rs.Close
				Set rs = Nothing

				'Set rs = con.Execute("SELECT COUNT(id) [ctr] FROM " & Session("tableName") & " (NOLOCK) WHERE timestop IS NOT NULL")
                'if cInt(rs(0)) >= 300 then
                '    alreadyTaken = True
                'end if
				'rs.Close
				'Set rs = Nothing

				returnCode = 1 'continue
				If alreadyTaken Then
					returnCode = -5 'already took the survey (using code for "full")
				Else
					Select Case cInt(segmentVar)
						Case 6	'利uick Service 
							If totVar >= Session("totalquickserve6")  or isNull(Session("totalquickserve6")) Then
								returnCode = -5 '"full"
							End If
						Case 3	'冶idscale Dining  
							If totVar >= Session("totalmidscale3") or isNull(Session("totalmidscale3")) Then
								returnCode = -5 '"full"
							End If
						Case 2	'低asual Dining 
							If totVar >= Session("totalcasual2") or isNull(Session("totalcasual2")) Then
								returnCode = -5 '"full"
							End If
						Case 1	'佝ine Dining  
							If totVar >= Session("totalfine1") or isNull(Session("totalfine1")) Then
								returnCode = -5 '"full"
							End If
						Case 18	'伯usiness & Industry 
							If totVar >= Session("totalbusiness18") or isNull(Session("totalbusiness18"))  Then
								returnCode = -5 '"full"
							End If
						Case 13	'佚ealthcare 
							If totVar >= Session("totalhealthcare13") or isNull(Session("totalhealthcare13"))  Then
								returnCode = -5 '"full"
							End If
						Case 12	'兵odging 
							If totVar >= Session("totallodging12") or isNull(Session("totallodging12"))  Then
								returnCode = -5 '"full"
							End If
						Case 16	'刨chools 
							If totVar >= Session("totalschools16") or isNull(Session("totalschools16"))  Then
								returnCode = -5 '"full"
							End If
						Case 17	'低olleges & Universities (30) 
							If totVar >= Session("totalcolleges17") or isNull(Session("totalcolleges17"))  Then
								returnCode = -5 '"full"
							End If
						Case 23	'別ther 
							If totVar >= Session("totalother23") or isNull(Session("totalother23")) Then
								returnCode = -5 '"full"
							End If
						Case 60	'低onvenience 
							If totVar >= Session("totalconvenience60") or isNull(Session("totalconvenience60")) Then
								returnCode = -5 '"full"
							End If
						Case 100	'劫esting
							If totVar >= Session("totaltesting100") or isNull(Session("totaltesting100")) Then
								returnCode = -5 '"full"
							End If
						Case Else ' any others... just in case
						    'ADDED THIS TO ACCEPT OTHER SEGMENT CODE other than 23 DECEMBER 13, 2010 by MARC
							If totVar >= Session("totalother23") or isNull(Session("totalother23")) Then
								returnCode = -5 '"full"
							End If
					End Select
				End If
			Else
				returnCode = -3 'not panelMember
			End If
		Case "nakano"
			Set rs = con.Execute("SELECT email FROM res20040805nakano WHERE Email = '" & emailAddress & "'")
			If rs.EOF Then
				panelMember = False
			Else
				panelMember = True
			End If
			rs.Close
			Set rs = Nothing
			If panelMember Then
				Set rs = con.Execute("SELECT TOP 1 id FROM " & Session("tableName") & " WHERE resp_id = '" & emailAddress & "' AND timeStop IS NOT NULL")
				If rs.EOF then
					alreadyTaken = False
				Else
					alreadyTaken = True
				End If
				rs.Close
				Set rs = Nothing
				If alreadyTaken Then
					returnCode = -5 'full/already taken
				End If
			Else
				returnCode = -3 'not panelMember
			End If	
			
		Case "none" ' no complete auto cutoff
			returnCode = 1
		Case Else ' everything else		
			returnCode = 1
			Set rs = con.Execute("SELECT count(distinct(id)) AS 'TOTAL' FROM " & Session("tableName") & " (NOLOCK) WHERE timeStop IS NOT NULL")
			If rs("TOTAL") >= Session("totalcompletes")  Then
				returnCode = -5 'full
			End If	
			rs.Close
			Set rs = Nothing
	End Select
	safeADODBclose()
	checkSurveyCompletes = returnCode
End Function

Function getNextSessionItem(sessionName)
	returnStr = ""
	If Session(sessionName) <> "" Then
		tmpArray = Split(Session(sessionName),"|")
		tmpStr = ""
		If ubound(tmpArray) < 0 Then
			returnStr = Session(sessionName)
		Else
			For i = 1 to ubound(tmpArray)
				if tmpStr <> "" then
					tmpStr = tmpStr & "|"
				end if
				tmpStr = tmpStr & tmpArray(i) 
			Next
			Session(sessionName) = tmpStr
			returnStr = tmpArray(0)
		End If
	End If
	getNextSessionItem = returnStr
End Function

Function startSurvey (resp_id)
	surveycur = safeSurveyCur()
	If cInt(surveycur) = 1 Then
		If resp_id <> "" Then
			colstr = ", resp_id"
			valstr = ", " & safeSql(resp_id)
		End If
		If Session("gozingui") <> "" Then
			colstr = ", ui"
			valstr = ", " & safeSql(Session("gozingui"))
		End If
        'surveyIDStat = getSurveyTable() 'LOBBY OCTOBER 3, 2012
		Set con = safeADODBcon()
		'/ 5.7.5; wiseobject <- CRITICAL! handle unexpected INSERT failure
		Dim uid, wz_i
		wz_i = 1
		'Response.Write("Session(tableName): " & surveyID)
		'Response.End()
		Do While wz_i <= 3
			con.Execute("INSERT INTO " & Session("tableName") & " ([timeStart] " & colstr & ") VALUES (GETDATE() "& valstr & ")")
			Set rs = con.Execute("SELECT @@IDENTITY as 'identity'") 
			'/Set rs = con.Execute("SELECT distinct(@@IDENTITY) as 'identity' FROM " & Session("tableName")) 
			If Not IsNull(rs) Then
				uid = rs("identity")
                'con.execute("INSERT INTO SurveysContactStatus VALUES (" & surveyIDStat & "," & uid & "," & safeSql(resp_id) & ",3)") 'LOBBY OCTOBER 3, 2012
				If Not IsNull(uid) Then
					Exit Do
				End If
				Set rs = Nothing
			End If	
			'/ have we lost connection?
			If con.State = 0 Then
				con.Open(connStr)
			End If
			'/ try to delete the record
			con.Execute("IF (SELECT COUNT(*) FROM " & Session("tableName") & " (NOLOCK) WHERE resp_id = " & safeSql(resp_id) & ") > 0" & vbCrLf & _
						    "DELETE FROM " & Session("tableName") & " WHERE resp_id = " & safeSql(resp_id))
			wz_i = wz_i + 1	
		Loop	
		'Session("uid") = EncryptString(uid)
		Session("uid") = uid
		SaveBrowser(safeSql(safeUid()))
		safeADODBclose()
	End If
End Function

Function endSurvey (endType)
    'surIDstat = getSurveyTable() 'LOBBY OCTOBER 3, 2012
	Select Case endType
		Case "success"
			'/ save end time 
			safeADODBexecute("UPDATE " & Session("tableName") & " SET timeStop = GETDATE() WHERE id = " & safeSql(safeUid()))
            
            'emailLob = safeADODBexecuteReturn("SELECT resp_id FROM " & Session("tableName") & " (NOLOCK) WHERE id = " & safeUid())
            'LOBBY OCTOBER 3, 2012
            'if not isnull(emailLob) then
            '    safeADODBexecute("UPDATE SurveysContactStatus SET c_status = 1 WHERE tbl_id = " & Session("uid") & " AND survey_id = " & surIDstat)
            'end if
			
            If Session("gozingid") <> "" Then
				Response.Redirect("/gozing/end.asp?sc=1")
			Else
                'InitAmGC()
				'Session.Abandon()

				additionalQueryStr = ""
				If Session("cti") = "true" Then
					additionalQueryStr = "&hid=cti"
				End If
				Response.Redirect("index.asp?surveycur=" & EncryptString("-6") & additionalQueryStr)
			End If
		Case "screen"
			'/ save end time 
			safeADODBexecute("UPDATE " & Session("tableName") & " SET timeScreen = GETDATE() WHERE id = " & safeSql(safeUid()))

            'emailLob = safeADODBexecuteReturn("SELECT resp_id FROM " & Session("tableName") & " (NOLOCK) WHERE id = " & safeUid())
            'LOBBY OCTOBER 3, 2012
            'if not isnull(emailLob) then
            '    safeADODBexecute("UPDATE SurveysContactStatus SET c_status = 2 WHERE tbl_id = " & Session("uid") & " AND survey_id = " & surIDstat)
            'end if

			If Session("gozingid") <> "" Then
				Response.Redirect("/gozing/end.asp")
			Else
				'Session.Abandon()
				Response.Redirect("index.asp?surveycur=" & EncryptString("-4"))
			End If
		
		case "quota"
		    safeADODBexecute("UPDATE " & Session("tableName") & " SET quota = GETDATE() WHERE id = " & safeSql(safeUid()))
            
            'emailLob = safeADODBexecuteReturn("SELECT resp_id FROM " & Session("tableName") & " (NOLOCK) WHERE id = " & safeUid())
            'LOBBY OCTOBER 3, 2012
            'if not isnull(emailLob) then
            '    safeADODBexecute("UPDATE SurveysContactStatus SET c_status = 4 WHERE tbl_id = " & Session("uid") & " AND survey_id = " & surIDstat)
            'end if

            'Session.Abandon()
		    Response.Redirect("index.asp?surveycur=" & EncryptString("-5"))
		    
		Case Else
			If Session("gozingid") <> "" Then
				Response.Redirect("/gozing/end.asp?sc=1")
			Else
				Response.Redirect("index.asp?surveycur=" & EncryptString(endType))
			End If
	End Select
End Function

Function startSurveySimple (resp_id)
	surveycur = safeSurveyCur()
	If cInt(surveycur) = 1 Then
		If resp_id <> "" Then
			colstr = ", resp_id"
			valstr = ", " & safeSql(resp_id)
		End If
		If Session("gozingui") <> "" Then
			colstr = ", ui"
			valstr = ", " & safeSql(Session("gozingui"))
		End If
		Set con = safeADODBcon()
		con.Execute("INSERT INTO " & Session("tableName") & "resp ([timeStart] " & colstr & ") VALUES (GETDATE() "& valstr & ")")
		Set rs = con.Execute("SELECT distinct(@@IDENTITY) as 'identity' FROM " & Session("tableName") & "resp") 
		uid = rs("identity")
		'Session("uid") = EncryptString(uid)
		Session("uid")  = uid
		safeADODBclose()
	End If
End Function

Function endSurveySimple (endType)
	Select Case endType
		Case "success"
			'/ save end time 
			safeADODBexecute("UPDATE " & Session("tableName") & "resp SET timeStop = GETDATE() WHERE id = " & safeSql(safeUid()))
			If Session("gozingid") <> "" Then
				Response.Redirect("/gozing/end.asp?sc=1")
			Else
				Session.Abandon()
				additionalQueryStr = ""
				If Session("cti") = "true" Then
					additionalQueryStr = "&hid=cti"
				End If
				Response.Redirect("index.asp?surveycur=" & EncryptString("-6") & additionalQueryStr)
			End If
		Case Else
			If Session("gozingid") <> "" Then
				Response.Redirect("/gozing/end.asp?sc=1")
			Else
				Response.Redirect("index.asp?surveycur=" & EncryptString(endType))
			End If
	End Select
End Function

'/ 5.19.4; wiseobject <-- copied from the original implementation of ernesto 
Function genericImplementUpdateRecord( var_key )	
	If request.Form(var_key) = "" Then
		Exit Function
	End If
	'response.Write(request.QueryString(var_key))
	'response.End()
	columnArray = split(request.Form(var_key),"|")
	for quesCount = 0 to ubound(columnArray)-1
		colstr = colstr & "[" &  Right(columnArray(quesCount),(len(columnArray(quesCount))-1))	 & "] = '" & replace(request.Form(columnArray(quesCount)),"'","''") & "', "		
	next
	colstr = Left(colstr,(len(colstr)-2))	
	'response.Write("colstr: " & colstr)
	safeADODBexecute("UPDATE " & Session("tableName") & " SET "& colstr &" WHERE id=" & safeSql(safeUid()))
End Function

Function genericImplementSaveAnswers( var_key )	
	If request.Form(var_key) = "" Then
		Exit Function
	End If
	'response.Write(var_key & "a<br>")
	'response.End()
	columnArray = split(request.Form(var_key),"|")
	sql_str = ""
	for quesCount = 0 to ubound(columnArray)-1
		sql_str = sql_str & " INSERT INTO " & Session("tableName") & " (respondentID, var_name, responseTime, answerValue) VALUES (" & safeSql(safeUid()) & ",'" & Right(columnArray(quesCount),(len(columnArray(quesCount))-1)) & "', getDate(), '" & replace(request.Form(columnArray(quesCount)),"'","''") & "') "
	next
	safeADODBexecute(sql_str)
End Function

'/ 8.3.4; ejohnson <-- returns an array of numbers in random order
Function buildRandArray(numToRandom)
	randArray = Array()
	ReDim randArray(numToRandom)
	For i = 1 to numToRandom
		check = "start"
		Do Until check = "yes"
			Randomize		
			tempRandom = Int((Rnd * numToRandom)+1)
			If randArray(tempRandom) = "" then
				randArray(tempRandom) = i
				check = "yes"
			End If
		Loop
		'getRand i
	Next
	buildRandArray = randArray
End Function
' Example of use:
'     someArray = buildRandArray(9)
'     for i = 1 to ubound(someArray)     'NOTE: starts at 1 position NOT 0
'          response.write someArray(i)
'     next
	
'/ 8.3.4; ejohnson <-- returns an array of numbers in random order
Function buildRandArrayReturn(numToRandom)
	tempArray = Array()
	ReDim tempArray(numToRandom)
	For i = 1 to numToRandom
		check = "start"
		Do Until check = "yes"
			Randomize		
			tempRandom = Int((Rnd * numToRandom)+1)
			If tempArray(tempRandom) = "" then
				tempArray(tempRandom) = i
				check = "yes"
			End If
		Loop
		'getRand i
	Next
	buildRandArrayReturn = tempArray
End Function
' Example of use:
'     someArray = buildRandArray(9)
'     for i = 1 to ubound(someArray)     'NOTE: starts at 1 position NOT 0
'          response.write someArray(i)
'     next


'/ 8.3.4; ejohnson <-- returns an array of numbers in random order
Function buildRandListReturn(listToRandom)
	tempArray = Split(listToRandom, "|")
	tempRandArray = buildRandArrayReturn(ubound(tempArray)+1)
	Dim returnArray()
	ReDim returnArray(ubound(tempArray))
	For i = 1 to ubound(tempRandArray)
		returnArray(i-1) = tempArray(tempRandArray(i)-1)
	Next
	buildRandListReturn = returnArray
End Function

Function getSingleColumnValue(colname, identName, identValue)
	Set con = safeADODBcon()	
	set rs = con.Execute("SELECT " & colname & " FROM " & Session("tableName") & " (NOLOCK) WHERE " & identName & " = '" & identValue & "'")
	resultVar = rs(0)
	rs.close
	set rs = Nothing
	safeADODBclose()
	getSingleColumnValue = resultVar
End Function

Function goNext(surveycur)
	Set con = safeADODBcon()
	Set rs = con.Execute("SELECT pageNum FROM " & Session("qTableName") & " (NOLOCK) WHERE pageNum = " & surveycur)
	If rs.EOF And rs.BOF Then
		endSurvey("success")
	Else
		response.redirect("index.asp?surveycur=" & EncryptString(surveycur))		
	End If
	rs.close
	set rs = Nothing
	safeADODBclose()
End Function	


Function buildRandSessionStr(conceptpages)
	conceptArray = buildRandListReturn(conceptpages)
	tmpStr = ""
	For i = 0 to ubound(conceptArray)
		if tmpStr <> "" then
			tmpStr = tmpStr & "|"
		end if
		tmpStr = tmpStr & conceptArray(i) 
	Next
	buildRandSessionStr = tmpStr
End Function

Function checkForDynamicText(text2check)
	if text2check <> "" then
		Dim objRegExp
		Set objRegExp = New regexp
		objRegExp.Global = True
		objRegExp.Pattern = "\^[A-Za-z0-9._%-]+\/\^"
		Dim svMatches
		Set svMatches = objRegExp.Execute(cStr(text2check))
		
		For Each objMatch in svMatches
			sessionStr = Left(objMatch.Value, Len(objMatch.Value) - 2)
			sessionStr = Right(sessionStr, Len(sessionStr) - 3)
			If isEmpty(Session(sessionStr)) Then
				'error
				text2check = "error"
			Else
				text2check = Replace(text2check, objMatch.Value, Session(sessionStr))
			End If
		Next
	End If
	checkForDynamicText = text2check
End Function

Function getSurveyQuestions(surveycur, is_alone)
	private_getSurveyQuestions surveycur, is_alone, false
End Function

Function getSurveyQuestions_2(surveycur, is_alone)
	private_getSurveyQuestions surveycur, is_alone, true
End Function

function private_getSurveyQuestions(surveycur, is_alone, is_random)
	set dbConn = safeADODBcon() 	
	If is_alone Then
		CommandText = "SELECT * FROM " & Session("qTableName") & " (NOLOCK) WHERE var_name = '" & surveycur & "'"
	Else
		Select Case surveycur
			Case 1000
				CommandText = "SELECT * FROM " & Session("qTableName") & " (NOLOCK) order by pageNum, orderNum"
			Case Else
				If is_random Then
					CommandText = "SELECT max(groupBy) FROM " & Session("qTableName") & " (NOLOCK) WHERE pageNum = " & surveycur
					Set RS = dbConn.Execute(CommandText)
					If NOT isNull(RS(0)) Then 
						maxGroupNum = cInt(RS(0)) 
					End If
					If maxGroupNum > 0 Then
						groupArray = buildRandArray(maxGroupNum)
						CommandText = "SELECT *, 0 as groupOrder FROM " & Session("qTableName") & " (NOLOCK) WHERE pageNum = " & surveycur & " AND groupBy is NULL " 
						For i = 1 to ubound(groupArray)
							If CommandText <> "" Then CommandText = CommandText & " UNION "
							CommandText = CommandText & " SELECT *, " & i & " as groupOrder FROM " & Session("qTableName") & " (NOLOCK) WHERE pageNum = " & surveycur & " AND groupBy = "  & groupArray(i)
						Next
						CommandText = CommandText & " ORDER BY groupOrder, orderNum"
					Else
						CommandText = "SELECT * " & sql_mod & " FROM " & Session("qTableName") & " (NOLOCK) WHERE pageNum = " & surveycur & " order by orderNum"
					End If
					RS.Close
					Set RS = Nothing
				Else
						CommandText = "SELECT * " & sql_mod & " FROM " & Session("qTableName") & " (NOLOCK) WHERE pageNum = " & surveycur & " order by orderNum"
				End If
		End Select
	End If
	set RS = dbConn.Execute(CommandText)
    

	Do While NOT RS.EOF
		function_name = RS("functionName")
		var_name = RS("var_name")
		pageNum = RS("pageNum")
		orderNum = RS("orderNum")
		q_text = RS("q_text")
		q_text = checkForDynamicText(q_text)
		e_text = RS("e_text")
		options = RS("options")
		options = checkForDynamicText(options)
		is_random = RS("is_random")
		is_other = RS("is_other")
		is_none = RS("is_none")
		t_size = RS("t_size")
		t_max = RS("t_max")
		beginNum = RS("beginNum")
		endNum = RS("endNum")
		bottomLabel = RS("bottomLabel")
		topLabel = RS("topLabel")
		initSelect = RS("initSelect")
		labelString = RS("labelString")
		offsetLabelString = RS("offsetLabelString")
		style = RS("style")
		v_text = RS("v_text")
		If Session(var_name) = "false" Then
			writeQuestion = false
		Else
			writeQuestion = true
		End If
		If writeQuestion Then
			If surveycur <> 1000 Then
				Response.Write "<!--"
			Else
				Response.Write "<p>"
			End If
				Response.Write "Question: " & var_name & "<br>Page Number: " & pageNum & "<br>Order Number: " & orderNum
			If surveycur <> 1000 Then
				Response.Write "-->"
			Else
				Response.Write "</p>"
			End If
			IF IsTestingPhase THEN
				q_text = " <span Style='color:yellow;'>( " & var_name & " )</span> " & q_text
			END IF

            select case var_name
                case "QB3"
                    QB2therText = dbConn.Execute("SELECT QB2_othertext, QB2_othertext2 FROM " & Session("tableName") & " where id = " & safeuid())                     
                    options = replace(options, "[OTHER]", QB2therText(0))                 
                    options = replace(options, "[OTHER2]", QB2therText(1))    
                case "QD11"                    
                    q_text = replace(q_text, "[QD10]", Session("TopQD10"))                    
                    e_text = replace(e_text, "[QD10]", Session("TopQD10"))                    
                 case "QD11intro"                    
                    q_text = replace(q_text, "[BOARD]", Session("TopQD10"))                    
                    q_text = replace(q_text, "[BOARDIMAGE]", Session("TopImage"))
                case "QG2"
                    QG1 = dbConn.Execute("SELECT QG1 FROM " & Session("tableName") & " where id = " & safeuid())
                    QG1Text = "NEW look"
                    if CInt(QG1(0)) = 2 then QG1Text = "CURRENT (original) look"
                    q_text = replace(q_text, "[QG1]", QG1Text)       
                    e_text = replace(e_text, "[QG1]", QG1Text)     
            end select
            Response.Write "<div class='questionContainer'>"
			    CALLFUNCTION function_name, var_name, q_text, e_text, options, is_random, is_other, is_none, t_size, t_max, beginNum, endNum, bottomLabel, topLabel, initSelect, labelString, offsetLabelString, style, v_text, dbConn
            Response.Write "</div>"
		End If
        
		RS.MoveNext
	Loop
	RS.Close
	Set RS = Nothing
	safeADODBclose()					
end function

Function CALLFUNCTION (function_name, var_name, q_text, e_text, options, is_random, is_other, is_none, t_size, t_max, beginNum, endNum, bottomLabel, topLabel, initSelect, labelString, offsetLabelString, style, v_text, dbConn)	
	CommandText = "SELECT q_text, e_text, options, is_random, is_other, is_none, t_size, t_max, beginNum, endNum, bottomLabel, topLabel, initSelect, labelString, offsetLabelString, style, v_text FROM questiontypes (NOLOCK) WHERE functionName = '" & function_name & "'"
	Set qRS = dbConn.Execute(CommandText)
	If NOT qRS.EOF Then
		stringToEval = function_name & " var_name "
		If NOT isNull(qRS("q_text")) Then			stringToEval = stringToEval & ", q_text"			End If
		If NOT isNull(qRS("e_text")) Then			stringToEval = stringToEval & ", e_text"			End If
		If NOT isNull(qRS("options")) Then			stringToEval = stringToEval & ", options"			End If
		If NOT isNull(qRS("is_random")) Then		stringToEval = stringToEval & ", is_random"			End If
		If NOT isNull(qRS("is_other")) Then			stringToEval = stringToEval & ", is_other"			End If
		If NOT isNull(qRS("is_none")) Then			stringToEval = stringToEval & ", is_none"			End If
		If NOT isNull(qRS("t_size")) Then			stringToEval = stringToEval & ", t_size"			End If
		If NOT isNull(qRS("t_max")) Then			stringToEval = stringToEval & ", t_max"				End If
		If NOT isNull(qRS("beginNum")) Then			stringToEval = stringToEval & ", beginNum"			End If
		If NOT isNull(qRS("endNum")) Then			stringToEval = stringToEval & ", endNum"			End If
		If NOT isNull(qRS("bottomLabel")) Then		stringToEval = stringToEval & ", bottomLabel"		End If
		If NOT isNull(qRS("topLabel")) Then			stringToEval = stringToEval & ", topLabel"			End If
		If NOT isNull(qRS("initSelect")) Then		stringToEval = stringToEval & ", initSelect"		End If
		If NOT isNull(qRS("labelString")) Then		stringToEval = stringToEval & ", labelString"		End If
		If NOT isNull(qRS("offsetLabelString")) Then	stringToEval = stringToEval & ", offsetLabelString"	End If
		If NOT isNull(qRS("style")) Then			stringToEval = stringToEval & ", style"				End If
		If NOT isNull(qRS("v_text")) Then			stringToEval = stringToEval & ", v_text"			End If
	End If
	qRS.Close
	Set qRS = Nothing

	response.write ("<!-- q" & var_name & "-->" &vbNewLine)
	var_name = "q" & var_name
	Execute(stringToEval)
End Function

function fnSort(aSort, AscDesc)
		Dim intTempStore, intPlaceStore
		Dim i, j 
		For i = 0 To UBound(aSort,2) - 1
			For j = i To UBound(aSort,2)
				'Sort Ascending
				if AscDesc = "ASC" Then
					if aSort(1,i) > aSort(1,j) Then
						intTempStore = aSort(1,i)
						intPlaceStore = aSort(0,i)
						aSort(1,i) = aSort(1,j)
						aSort(0,i) = aSort(0,j)
						aSort(1,j) = intTempStore
						aSort(0,j) = intPlaceStore
					End if 'i > j
					'Sort Descending
				Else
					if aSort(1,i) < aSort(1,j) Then
						intTempStore = aSort(1,i)
						intPlaceStore = aSort(0,i)
						aSort(1,i) = aSort(1,j)
						aSort(0,i) = aSort(0,j)
						aSort(1,j) = intTempStore
						aSort(0,j) = intPlaceStore
					End if 'i < j
				End if 'intAsc = 1				
				If (aSort(1,i) = aSort(1,j)) and (i <> j) Then
					Randomize
					tempRandom = Int(Rnd * 2)
					If tempRandom = 1 Then
						intPlaceStore = aSort(0,i)
						aSort(0,i) = aSort(0,j)
						aSort(0,j) = intPlaceStore
					End If
				End If
			Next 'j
		Next 'i
		fnSort = aSort
End function 
' function to sort a two dimensional array.
' (0,x) = stores identifier (can be string or number)
' (1,x) = stores number to be sorted

Sub AddMiscColumns()
	Dim sTemplate, sSql
	sTemplate = "if exists(select * FROM INFORMATION_SCHEMA.COLUMNS where table_name = '" & Session("tableName") & "' and COLUMN_NAME='[XX]') select 1 Else select 0"
	sSql = replace(sTemplate,"[XX]", "LastPage")
	if cInt(safeADODBexecuteReturn(sSql)) = 0 then
		sSql = "ALTER TABLE [" & Session("tableName") & "] Add LastPage int"
		safeADODBexecute sSql
	end if

	sSql = replace(sTemplate,"[XX]", "SavedSessions")
	if cInt(safeADODBexecuteReturn(sSql)) = 0 then
		sSql = "ALTER TABLE [" & Session("tableName") & "] Add SavedSessions varchar(MAX)"
		safeADODBexecute sSql
	end if
	
End Sub

Function SaveBrowser(uid)
	if (uid="") then 
		uid=safeSql(safeUid())
	end if	
	
	Set con = safeADODBcon()
	dim sql
	sql = "select COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS where table_name = '" & Session("tableName") & "' and COLUMN_NAME='Browser'"
	'response.write sql
	'response.end()
	Set rs = con.Execute(sql) 

	'If Not IsNull(rs) Then
	'response.write rs.RecordCount
	'response.end()
	Dim count
	count = 0
	
	While Not rs.EOF
		count = count + 1
		rs.MoveNext
	Wend 

		if count=0 then
		
			'/ have we lost connection?
			If con.State = 0 Then
				con.Open(connStr)
			End If
			'response.write "entered"
			'response.end()
			con.Execute("ALTER TABLE " & Session("tableName") & " ADD browser varchar(MAX)")

			'Set rs = Nothing
		end if	
	'End If	
	
	Dim browser 
	browser = safeSql(Request.ServerVariables("http_user_agent"))
	sql = "UPDATE " & Session("tableName") & " SET browser=" & browser &" WHERE id = " & safeSql(safeUid())
	'response.write sql
	'response.end()
	con.Execute(sql)
End Function

'2008.04.26 (oliver) -> Added the following scripts to this template (record.asp)
Sub UpdateLastPage(tblName,ID,last_page)
Dim UpdateSql
Dim ColStr
	ColStr = ""
	ColStr = "[lastpage]"
	ColStr = ColStr & " = '" & last_page & "' " 
	UpdateSql = "UPDATE " & tblName & " SET " & ColStr 
	UpdateSql = UpdateSql & " WHERE id = " & ID
	SafeADODBexecute(UpdateSql)
End Sub
'20080414 - bernard - added
Sub Skip(ColumnSql)
	ExecuteSql = CleanSkippedQuestions(Session("tableName"),safeSql(safeUid()),ColumnSql)
	SafeADODBexecute(ExecuteSql)										
End Sub
Function CleanSkippedQuestions(tblName,ID,QuestionIDs)
Dim UpdateSql,Qctr
Dim QuestionArr
Dim ColStr
	ColStr = ""
	If Instr(QuestionIDs,"|") > 0 Then
		QuestionArr = Split(QuestionIDs,"|")
		For Qctr = LBound(QuestionArr) To UBound(QuestionArr)
			If ColStr <> "" Then
				ColStr = ColStr  & ","
			End If
			ColStr = ColStr & "[" & QuestionArr(Qctr) & "]" & "= NULL" 
		Next
		UpdateSql = "UPDATE " & tblName & " SET " & ColStr 
		UpdateSql = UpdateSql & " WHERE id = " & ID
	Else
		UpdateSql = "UPDATE " & tblName & " SET [" & QuestionIDs & "]= NULL"
		UpdateSql = UpdateSql & " WHERE id = " & ID
	End If
	CleanSkippedQuestions = UpdateSql
End Function
'---------------------------------------------------------------------------------
%>

<%
'ENCRYPTION / DECRYPTION ADDED January 27, 2011 --- MARC


'### To encrypt/decrypt include this code in your page 
'### strMyEncryptedString = EncryptString(strString)
'### strMyDecryptedString = DeCryptString(strMyEncryptedString)
'### You are free to use this code as long as credits remain in place
'### also if you improve this code let me know.


Function EncryptString(strString)
'####################################################################
'### Crypt Function (C) 2001 by Slavic Kozyuk grindkore@yahoo.com ###
'### Arguments: strString <--- String you wish to encrypt         ###
'### Output: Encrypted HEX string                                 ###
'####################################################################
On error resume next

Dim CharHexSet, intStringLen, strTemp, strRAW, i, intKey, intOffSet
Randomize Timer

intKey = Round((RND * 1000000) + 1000000)   '##### Key Bitsize
intOffSet = Round((RND * 1000000) + 1000000)   '##### KeyOffSet Bitsize

	If IsNull(strString) = False Then
		strRAW = strString
		intStringLen = Len(strRAW)
				
				For i = 0 to intStringLen - 1
					strTemp = Left(strRAW, 1)
					strRAW = Right(strRAW, Len(strRAW) - 1)
					CharHexSet = CharHexSet & Hex(Asc(strTemp) * intKey) & Hex(intKey)
				Next
		
		EncryptString = CharHexSet & "|" & Hex(intOffSet + intKey) & "|" & Hex(intOffSet)
	Else
		EncryptString = ""
	End If
	
	if 	Not Err.Number = 0 then
		response.redirect("index.asp?surveycur=" & EncryptString("-7"))
	end if
End Function




Function DeCryptString(strCryptString)
'####################################################################
'### Crypt Function (C) 2001 by Slavic Kozyuk grindkore@yahoo.com ###
'### Arguments: Encrypted HEX stringt    						  ###
'### Output: Decrypted ASCII string                               ###
'####################################################################
'### Note this function uses HexConv() and get_hxno() functions   ###
'### so make sure they are not removed							  ###
'####################################################################

Dim strRAW, arHexCharSet, i, intKey, intOffSet, strRawKey, strHexCrypData

	On error resume next
	strRawKey = Right(strCryptString, Len(strCryptString) - InStr(strCryptString, "|"))
	intOffSet = Right(strRawKey, Len(strRawKey) - InStr(strRawKey,"|"))
	intKey = HexConv(Left(strRawKey, InStr(strRawKey, "|") - 1)) - HexConv(intOffSet)
	strHexCrypData = Left(strCryptString, Len(strCryptString) - (Len(strRawKey) + 1))

	
	arHexCharSet = Split(strHexCrypData, Hex(intKey))
		
		For i=0 to UBound(arHexCharSet)
			strRAW = strRAW & Chr(HexConv(arHexCharSet(i))/intKey)
		Next
		
	DeCryptString = strRAW
	
	if 	Not Err.Number = 0 then
		response.redirect("index.asp?surveycur=" & EncryptString("-7"))
	end if
End Function



Private Function HexConv(hexVar)
Dim hxx, hxx_var, multiply		
		IF hexVar <> "" THEN
			hexVar = UCASE(hexVar)
			hexVar = StrReverse(hexVar)
			DIM hx()
			REDIM hx(LEN(hexVar))
			hxx = 0
			hxx_var = 0
			FOR hxx = 1 TO LEN(hexVar)
				IF multiply = "" THEN multiply = 1
				hx(hxx) = mid(hexVar,hxx,1)
				hxx_var = (get_hxno(hx(hxx)) * multiply) + hxx_var
				multiply = (multiply * 16)
			NEXT
			hexVar = hxx_var
			HexConv = hexVar
		END IF
End Function
	
Private Function get_hxno(ghx)
		If ghx = "A" Then
			ghx = 10
		ElseIf ghx = "B" Then
			ghx = 11
		ElseIf ghx = "C" Then
			ghx = 12
		ElseIf ghx = "D" Then
			ghx = 13
		ElseIf ghx = "E" Then
			ghx = 14
		ElseIf ghx = "F" Then
			ghx = 15
		End If
		get_hxno = ghx
End Function
	
Function BackToPage(var_key,surcur)
	Dim i
	Dim null_ctr
	null_ctr = 0
	If request.Form(var_key) = "" Then
		Exit Function
	End If

	columnArray = split(request.Form(var_key),"|")
	for quesCount = 0 to ubound(columnArray)-1
		colstr = "[" &  Right(columnArray(quesCount),(len(columnArray(quesCount))-1))	 & "] is NULL "

	    i = cInt(safeADODBexecuteReturn("select count(*) from " & Session("tableName") & " (NOLOCK) where " & colstr & " and id = " & safeUid()))

	    if i = 1 then
	        null_ctr = null_ctr + 1
	    end if

	next

	if null_ctr <> 0 then
	    response.redirect("index.asp?Nullerror=1&surveycur=" & EncryptString(surcur))
	end if
End Function
%>