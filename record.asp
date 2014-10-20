<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/db/database.asp"-->
<!--#include file="inc/common.asp"-->

<%
If Request.QueryString("surveycur") <> "" Then
	surveycur = DeCryptString(Request.Querystring("surveycur"))
End If
	
if surveycur = 1 then
	Dim rsLastSession
	If Request.QueryString("resp_id") <> "" Then
		resp_id = Request.QueryString("resp_id")
		Session("resp_id_Save") = resp_id
	Else
		If Request.Form("qresp_id") <> "" Then
			resp_id = Request.Form("qresp_id")
			Session("resp_id_Save") = resp_id
		Else
			resp_id = ""
		End If
	End If
	if resp_id <> "" then
		'------------- FOR RELOADING INCOMPLETE SURVEYS
		Set rsLastSession = LoadLastSession(resp_id)
		'-------------
		If rsLastSession.EOF Then
			If Session("is_panel") Then
				isFull = checkSurveyCompletes(resp_id)
				If (isFull) < 0 Then
					response.redirect("index.asp?surveycur=" & EncryptString(isFull))
				End If
			End If
			
			StartSurvey resp_id
			genericImplementUpdateRecord "questionIDs"			
			nextNum = 2
		Else
		'------------- FOR RELOADING INCOMPLETE SURVEYS
			Dim SavedSessions, LastPage, arySessions, ndx_session
			Dim data_error
			data_error = false
			if (isNull(rsLastSession("SavedSessions"))) then
				sql = "Update " & Session("tableName") & " set SavedSessions = resp_id, resp_id='" & Replace(resp_id,"'","") & "' where resp_id = '" & Replace(resp_id,"'","") & "'"
				call safeADODBexecute( sql )
				'response.redirect("index.asp?surveycur=2")
				data_error = true
			end if
			if not data_error then
				SavedSessions = cStr(rsLastSession("SavedSessions"))
				LastPage = cStr(rsLastSession("LastPage"))
				' "-==-" & Session(sObj) & "-||-"
				arySessions = split(SavedSessions,"-||-")
				for ndx_session = lBound(arySessions) to uBound(arySessions)
					SavedSessions = split(arySessions(ndx_session),"-==-")
					Session(SavedSessions(0)) = SavedSessions(1)
				next
				nextNum = LastPage
			else
				StartSurvey resp_id
				genericImplementUpdateRecord "questionIDs"			
				nextNum = 2
			end if
		'-------------
		End If
	Else
		StartSurvey resp_id
		genericImplementUpdateRecord "questionIDs"			
		nextNum = 2
	End If
	
	UpdateLastPage Session("tableName"),safeUid(),nextNum
	GoNext(nextNum)
End If

If Session("uid") <> "" Then
	uid = CInt(Session("uid"))
	genericImplementUpdateRecord "questionIDs" 
    BackToPage "questionIDs",surveycur
	nextNum = surveycur + 1
        
	Select Case surveycur
        case 2            
            'ITRACK
            'Response.Redirect("itrack.asp?assetGroupId=5257&respid=" & Session("resp_id_Save"))                
	End Select
        
	UpdateLastPage Session("tableName"),safeUid(),nextNum

    

	GoNext(nextNum)
Else
	response.redirect("index.asp?surveycur=" & EncryptString("-1"))
End If

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

Function GetIntValue(var_name)
	Dim var_value
	var_value = Request.Form(var_name)
	if isNumeric(var_value) Then
		var_value = cInt(var_value)
	else
		var_value = 0
	end if
	GetIntValue = var_value
End Function

'------------- FOR RELOADING INCOMPLETE SURVEYS

Sub UpdateLastPage(tblName,ID,last_page)
	Dim cmd, sessions, sObj

	sessions = ""
	for each sObj in session.Contents
		If Session(sObj) <> "" Then
			sessions = sessions & sObj & "-==-" & Session(sObj) & "-||-"
		End If
	next
	if sessions <> "" Then
		sessions = left(sessions, len(sessions)-4)
	end if
	
	set cmd = CreateObject("ADODB.Command")
	set cmd.ActiveConnection = safeADODBcon()
	with cmd
		.commandText = "UPDATE " & tblName & " SET [lastpage] = ?, [SavedSessions] = ? WHERE [id] = ?"
		.commandType = 1
		.parameters.append .createParameter("@lastpage",3,1,,last_page)
		.parameters.append .createParameter("@SavedSessions",200,1,-1,sessions)
		.parameters.append .createParameter("@id",3,1,,ID)

		.execute()
	end with
	set cmd = nothing
End Sub

Function LoadLastSession(resp_id)
	dim sql, tableName, cmd, rs
    tableName = Session("tableName")
    sql = "Select top 1 LastPage,SavedSessions from " & tableName & " (NOLOCK) where resp_id = ? AND timestop is null AND timescreen is null order by id desc"
    'sql = "Select top 1 LastPage,SavedSessions from " & tableName & " (NOLOCK) where resp_id = ? AND timestop is null AND timescreen is null AND quota IS NULL order by id desc"
	
	
    'sql = "if not exists(select * " & _
	'		" from " & tableName &  " " & _
	'		" (NOLOCK) where resp_id = ? " & _
	'			" and (timestop is not null OR timescreen is not null)) " & _
	'	" Select top 1 LastPage,SavedSessions from " & tableName & " " & _
	'			" (NOLOCK) where resp_id = ? AND timestop is null AND timescreen is null order by id " & _
	'	 "else " & _
	'	 	"	Select top 1 LastPage,SavedSessions from " & tableName & " " & _
	'			" (NOLOCK) where id = 0"

    'WITH the "quota" column
	'sql = "if not exists(select * " & _
	'		" from " & tableName &  " " & _
	'		" (NOLOCK) where resp_id = ? " & _
	'			" and (timestop is not null OR timescreen is not null OR quota is not null)) " & _
	'	" Select top 1 LastPage,SavedSessions from " & tableName & " " & _
	'			" (NOLOCK) where resp_id = ? AND timestop is null AND timescreen is null AND quota is null order by id " & _
	'	 "else " & _
	'	 	"	Select top 1 LastPage,SavedSessions from " & tableName & " " & _
	'			" (NOLOCK) where id = 0"

    'WITH the "quota" column / and re-take enabled
	'sql = "if not exists(select * " & _
	'		" from " & tableName &  " " & _
	'		" (NOLOCK) where resp_id = ? " & _
	'			" and (timestop is not null OR timescreen is not null OR quota is not null) and id = " & safeUid() & ") " & _
	'	" Select top 1 LastPage,SavedSessions from " & tableName & " " & _
	'			" (NOLOCK) where resp_id = ? AND timestop is null AND timescreen is null AND quota is null order by id desc " & _
	'	 "else " & _
	'	 	"	Select top 1 LastPage,SavedSessions from " & tableName & " " & _
	'			" (NOLOCK) where id = 0"

	set cmd = CreateObject("ADODB.Command")
	set cmd.ActiveConnection = safeADODBcon()
	with cmd
		.commandText = sql
		.commandType = 1
		
		.parameters.append .createParameter("@resp_id1",200,1,8000,resp_id)
		'.parameters.append .createParameter("@resp_id2",200,1,8000,resp_id)
		
		set rs = .execute
	end with
	Set LoadLastSession = rs
	set cmd = nothing
	set rs = nothing	
End Function
'-------------
'MARC April 21, 2010 CONCEPT ROTATION CODE
Function IsLastConcept(current_concept,str_list)
'determine if last concept.
Dim QA1_rand_ary
Dim is_last_concept
	is_last_concept = False
	If Instr(str_list,"|") > 0 Then
		QA1_rand_ary = Split(str_list,"|")
	Else
		Redim QA1_rand_ary(0)
		QA1_rand_ary(0) = str_list
	End If
	For iJ = LBound(QA1_rand_ary) To UBound(QA1_rand_ary)
		If CInt(QA1_rand_ary(iJ)) = CInt(current_concept) Then
			'found the current position of concept from array (iJ)
			'check if iJ is the last element
			If iJ = UBound(QA1_rand_ary) Then
				'this is the last element,end the survey.
				is_last_concept = True
				Exit For
			End If
		End If
	Next
	IsLastConcept = is_last_concept	
End Function

Function NextElementIndex(current_concept,str_list)
'determine next index.
Dim QA2_ary
Dim next_element_index
	next_element_index = 0
	If Instr(str_list,"|") > 0 Then
		QA2_ary = Split(str_list,"|")
	Else
		Redim QA2_ary(0)
		QA2_ary(0) = str_list
	End If
	For iJ = LBound(QA2_ary) To UBound(QA2_ary)
		If CInt(QA2_ary(iJ)) = CInt(current_concept) Then
			'found the current position of concept from array (iJ)
			next_element_index = QA2_ary(iJ+1)
			Exit For
		End If
	Next
	NextElementIndex = next_element_index
End Function
'--------------------------------------
Function NextConcept( curPage )
	dim nxtPage, ndx, aryPages, sPagesToVisit
	sPagesToVisit = Session("ConceptOrder") 'placeholder for surveycurs
	If sPagesToVisit = "" Then 'prevent randomization if session is not yet set
		NextConcept = 5 'return to page where session will be initialized
		exit function
	End If
	aryPages = split(sPagesToVisit,"|")
	for ndx	= lBound(aryPages) to uBound(aryPages)
		If cInt(curPage) = cInt(aryPages(ndx)) Then
			if ndx <> uBound(aryPages) Then
				nxtPage = aryPages(ndx+1)
			else
				nxtPage = 100 ' page after the randomizations
			end if
			exit for
		End If
	next
	NextConcept = nxtPage
End Function
%>