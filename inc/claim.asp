<%
Sub InitAmGC()
    Dim pDebug
    Dim rpemail
    Dim rptimestart
    Dim rptimestop
    Dim respID
    Dim alreadyCnt

    Dim surveyid
    Dim surAmt, surName
    Dim uniqueID
    Dim AmReqId,AmResId,AmClaimCode,AmAmount
    
    Dim repEmail
    Dim repCode
    Dim emailType
    Dim getInfo
    Dim amMsg, autoSnd
    repEmail = ""
    repCode = ""
    getInfo = ""
    amMsg = ""
    alreadyCnt = 0

    emailType = 1 'change this to 1 if live
    pDebug = false 'true = testing, false = live

    set conRep = safeADODBcon()
    set rsRep = conRep.execute("SELECT resp_id, timestart, timestop, id FROM " & Session("tableName") & " (NOLOCK) WHERE id = " & safeUid())
    rpemail = rsRep(0)
    rptimestart = rsRep(1)
    rptimestop = rsRep(2)
    respID = rsRep(3)

    rsRep.close
    set rsRep = nothing
    set conRep = nothing
    conRep = null
    surveyid = getSurveyTable()

    surAmt = safeADODBexecuteReturn("lsp_amzonGift_get_sur_amt " & surveyid)
    surName = safeADODBexecuteReturn("lsp_amzonGift_get_sur_name " & surveyid)

    if isnull(rpemail) then
        rpemail = ""
    end if

    if not rpemail = "" then
        'for x = 1 to 3
        alreadyCnt = cInt(safeADODBexecuteReturn("lsp_amzonGift_check_if_exist " & surveyid & "," & respID))
        if alreadyCnt = 0 then
            safeADODBexecute("lsp_amzonGift_save_chq_initial '" & replace(rpemail,"'","''") & "'," & surveyid & ",'" & rptimestart & "','" & rptimestop & "'," & surAmt & "," & respID)
            uniqueID = safeADODBexecuteReturn("lsp_amzonGift_get_resp_unique_id '" & replace(rpemail,"'","''") & "'," & surveyid & "," & respID)

            if pDebug then
                AmReqId = "testREQ"
                AmResId = "testRESP"
                AmClaimCode = "testCODE"
            else
                url = "http://www.datassential.com/certphp/index.php?amount=" & surAmt & "&id=" & uniqueID & "&key=" & hex_md5("Opera" & uniqueID & surAmt) & "&k=IilD62w1y2AQu65j5927908v13p2P3vW&project=" & surName
                set xmlhttp = CreateObject("MSXML2.ServerXMLHTTP") 
                xmlhttp.open "GET", url, false 
                xmlhttp.send "" 
                smp = xmlhttp.responseText
                smp = split(smp,vblf)

                AmReqId = smp(0)
                AmResId = smp(1)
                AmClaimCode = smp(2)
                'AmAmount = smp(3) 'ignore
            end if

            SafeADODBexecute("lsp_amzonGift_update_chq " & uniqueID & ",'" & AmReqId & "','" & AmResId & "','" & AmClaimCode & "'")
            set xmlhttp = nothing

            '12/12/11
            Set emailCon = safeADODBcon()
            set getEmails = emailCon.execute("lsp_amzonGift_get_info '" & uniqueID & "'")
            repEmail = getEmails(0)
            amMsg = getEmails(2)
            autoSnd = getEmails(3)

            getEmails.close
            set getEmails = nothing
            set emailCon = nothing
            emailCon = null

            if autoSnd = 1 then
                if isnull(AmClaimCode) then AmClaimCode = ""
                if AmClaimCode = "" then
                    emailType = 3
                else
                    safeADODBexecute("lsp_amzonGift_save_sent " & uniqueID)
                end if
                getInfo = ConsEmail(AmClaimCode,surAmt,amMsg)
                select case emailType
                    case 1
                        'repEmail = repEmail & "webtestsential.not" 'remove this / comment this if live
                        SendEmail repEmail,"Datassential - Amazon Gift Code",getInfo
                    case 2
                        repEmail = repEmail & "webtestsential.not"
                        SendEmail repEmail,"TEST=Datassential - Amazon Gift Code=TEST",getInfo
                end select
            end if
        end if
        'next
    end if
End Sub

Function ConsEmail(gcode,dsign,msg)
    Dim htmlStr
    if isnull(msg) then
        msg = ""
    end if
    htmlStr = ""        
    htmlStr = htmlStr & "<html><body>"
    htmlStr = htmlStr & "<span style='font-family:Verdana;font-size:13.5px;'>" & msg & "</b></span>" & vbcrlf & vbcrlf
    if isnull(gcode) then gcode = ""
    if gcode = "" then
        htmlStr = htmlStr & "<span style='font-family:Verdana;font-size:13.5px;'>We will send your Amazon GC code shortly.</b></span>" & vbcrlf & vbcrlf
    else
        htmlStr = htmlStr & "<span style='font-family:Verdana;font-size:13.5px;'>Here is your $" & dsign & " Amazon GC code: <b>" & gcode & "</b></span>" & vbcrlf & vbcrlf
    end if
    htmlStr = htmlStr & "<span style='font-family:Verdana;font-size:12px; color:blue;'>Datassential Research Inc.</span>" & vbcrlf & vbcrlf
	htmlStr = htmlStr & "</body></html>"
	htmlStr = Replace(htmlStr, vbcrlf, "<br>")
    ConsEmail = htmlStr
End Function
%>