<%@LANGUAGE="VBSCRIPT"%>
<!--#include virtual="/db/database.asp"-->
<!--#include file="inc/common.asp"-->
<!--#include file="inc/questions.asp"-->
<!--#include file="browserdetection.asp"-->
<% 'The above Include statements call and define the path to resuable functions and data sources.	
' Sets the surveycur variable.  If there is already a value for this variable, it is retained :
Dim IsTestingPhase,IsClosed,sur_status
IsTestingPhase = True
surveyID = getSurveyTable()
IsClosed = False
If request.QueryString("surveycur") <> "" Then
	dim iMaxPage
	surveycur = DeCryptString(request.QueryString("surveycur"))
	if not isNumeric(surveycur) Then
		surveycur = -7
	Else
		surveycur = cInt(surveycur)
	End If
	If Session("qTableName") <> "" Then
		iMaxPage = cInt(safeADODBexecuteReturn( "select max(pageNum) from " & Session("qTableName") & " (NOLOCK)"))
		If surveycur > iMaxPage AND surveycur <> 1000 Then
			surveycur = -7
		End If
	End If
	If surveycur = 1000 Then
		getSurveyVariables(surveyID)
	End If
Else
	Session("uid") = ""	'clear session user id variable in case there is one.
	surveycur = 1	'Sets the "surveycur" variable at page one, unless it already has a valid value.  Querystring defines its value thereafter.
	getSurveyVariables(surveyID) 'initialize session variables based on surMaster (dbTable) ID
	AddMiscColumns()
	'2009.02.17 - Oliver
	'NOTE: Implement this for all incoming surveys.
	'Added a code that will check if the survey was closed.
	'If true, set the surveycur to "survey is now full".	
	sur_status = cInt(safeADODBexecuteReturn("select isnull(status,0) from surmaster (NOLOCK) where id =" & surveyID))
	'2008.11.21 - oliver
	'Added a code that will remove the testing phase label.
	If sur_status > 2 and sur_status < 4 Then 'hide testing phase
		IsTestingPhase = false
	End If
	If sur_status = 4 Or sur_status = 5 Then 'survey is closed/cancelled.
		IsClosed = True
		IsTestingPhase = false
	Else
	
	'MARC April 21, 2010 if CONSUMER uncomment below
	
	'	If Request.QueryString("subsid") <> "" Then
	'	If Session("tableName") = "" OR IsEmpty(Session("tableName")) Then 
	'		surveyID = getSurveyTable()
	'		getSurveyVariables(surveyID)
	'	End If
	'	resp_id = Request.QueryString("subsid")
	'	Session("subsid") = resp_id
	'	sql=""
	'	sql = " SELECT count(*) FROM " & Session("tableName")	
	'	sql = sql & " WHERE resp_id = '" & Replace(resp_id,"'","") & "'"  
	'	sql = sql & " AND timeStart IS NOT NULL "
	'	sql = sql & " AND (timeStop IS NOT NULL OR timescreen IS NOT NULL) "		
	'	intY = safeADODBexecuteReturn(sql)
	'	If intY > 0 Then 
	'		surveycur=-8	' The respondent aleady taken the survey
	'	End If
	'	Dim rsLastSession
	'	Set rsLastSession = LoadLastSession(resp_id)
	'	If NOT rsLastSession.EOF Then
	'		Dim SavedSessions, LastPage, arySessions, ndx_session
	'		if (isNull(rsLastSession("SavedSessions"))) then
	'			sql = "Update " & Session("tableName") & " set SavedSessions = resp_id, resp_id='" & Replace(resp_id,"'","") & "' where resp_id = '" & Replace(resp_id,"'","") & "'"
	'			call safeADODBexecute( sql )
	'			response.redirect("index.asp?subsid=" & resp_id)
	'		end if
	'		SavedSessions = cStr(rsLastSession("SavedSessions"))
	'		LastPage = cStr(rsLastSession("LastPage"))
	'		' "-==-" & Session(sObj) & "-||-"
	'		arySessions = split(SavedSessions,"-||-")
	'		for ndx_session = lBound(arySessions) to uBound(arySessions)
	'			SavedSessions = split(arySessions(ndx_session),"-==-")
	'			Session(SavedSessions(0)) = SavedSessions(1)
	'		next
	'		surveycur = LastPage
	'	end if
	'Else
	'	surveycur=-3
	'End If
	
	'-----------------------------------------------
		if NOT Session("is_panel") Then
			isFull = checkSurveyCompletes("")
			If (isFull) < 0 Then
		 		response.redirect("index.asp?surveycur=" & EncryptString(isFull))
			End If			 	
		end if
	End If
End If
If request.Querystring("resp_id") <> "" Then
	resp_id = request.Querystring("resp_id")	
End If

'Jan 15, 2010 --MARC
sur_status = cInt(safeADODBexecuteReturn("select isnull(status,0) from surmaster (NOLOCK) where id =" & surveyID))
If sur_status > 2 and sur_status < 4 Then 'hide testing phase
	IsTestingPhase = false
End If

If IsClosed = True Then
	surveycur = -5 'survey is full message.
End If

'MARC April 21, 2010 if CONSUMER uncomment below
'Function LoadLastSession(resp_id)
'	dim sql, tableName, cmd, rs
'	tableName = Session("tableName")
'	sql = "if not exists(select * " & _
'			" from " & tableName &  " " & _
'			" (NOLOCK) where resp_id = ? " & _
'				" and (timestop is not null OR timescreen is not null)) " & _
'		" Select top 1 LastPage,SavedSessions from " & tableName & " " & _
'				" (NOLOCK) where resp_id = ? AND timestop is null AND timescreen is null order by id " & _
'		 "else " & _
'		 	"	Select top 1 LastPage,SavedSessions from " & tableName & " " & _
'				" (NOLOCK) where id = 0"
'
'	set cmd = CreateObject("ADODB.Command")
'	set cmd.ActiveConnection = safeADODBcon()
'	with cmd
'		.commandText = sql
'		.commandType = 1
'		
'		.parameters.append .createParameter("@resp_id1",200,1,8000,resp_id)
'		.parameters.append .createParameter("@resp_id2",200,1,8000,resp_id)
'		
'		set rs = .execute
'	end with
'	Set LoadLastSession = rs
'	set cmd = nothing
'	set rs = nothing	
'End Function
'------------------------------
%>
<HTML>
<HEAD>
<TITLE>Datassential Survey <%= surveyName %></TITLE> <!-- Builds page title from "surveyName" variable. -->	
	<style type="text/css">
		/* styles specific to this survey go here: */
.concept-Statement{
    margin:10px auto;
}

.concept-Statement p{
    text-align:center;
    font-size:22px;
}

.concept-details{
    width:70%;
    margin:10px auto;
}

.concept-details .image-label td{
    text-align:center;
    padding:10px;
    font-weight:bold;
}
.concept-details .image-concept{
    width:50%;
    text-align:center;
    margin:10px auto;
}

.concept-details .image-concept span {
        display: block;
        padding-bottom:10px;
        padding-top:10px;
        font-size:16px;
    }
.concept-details .row span {
        display: block;
        padding-bottom:10px;
        padding-top:10px;
        text-align:center;
        font-size:16px;
    }

.concept-details span:first-child{
/*font-size:22px;*/
font-weight:bold;
}

.ctable {
	display: table;   /* Allow the centering to work */
	margin: 0 auto;
	
}

#horizontal-list {
	max-width: 396px; 
	<!-- list-style: none; -->
	<!-- list-style-type:square; -->
	padding-top: 20px;
}
	
#horizontal-list li {
		 display: block;
	}
	</style>
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
  
    <link rel="stylesheet" href="css/survey.css" />       
    <script language="javascript" src="/js/form.js"></script>	<!-- Sets javascript source location. -->
	<script language="javascript" src="/js/scripts.js"></script>	<!-- Sets javascript source location. -->

    <link rel="stylesheet" href="bootstrap-3.2.0-dist/css/bootstrap.css" />
    <link rel="stylesheet" href="bootstrap-3.2.0-dist/css/bootstrap-theme.css" />
    <link rel="stylesheet" href="bootstrap-3.2.0-dist/css/bootstrap.min.css" />
    <link rel="stylesheet" href="bootstrap-3.2.0-dist/css/bootstrap-theme.min.css" />    
    


    <script language="javascript" src="Scripts/jquery-1.11.1.min.js"></script>
    <script language="javascript" src="Scripts/jquery-1.11.1.js"></script>
    <script type="text/javascript" src="bootstrap-3.2.0-dist/js/bootstrap.js"></script>
    <script type="text/javascript" src="bootstrap-3.2.0-dist/js/bootstrap.min.js"></script>
    <script type='text/javascript' src="Scripts/css3-mediaqueries.js"></script>  

    <script type="text/javascript" src='tooltip.js'></script>



    <!-- bootstrap JQuery references -->
   
   

	<SCRIPT LANGUAGE="JavaScript">
	window.onerror = null;
	var bName = navigator.appName;
	var bVer = parseInt(navigator.appVersion);
	var NS4 = (bName == "Netscape" && bVer >= 4);
	var IE4 = (bName == "Microsoft Internet Explorer" 
	&& bVer >= 4);
	var NS3 = (bName == "Netscape" && bVer < 4);
	var IE3 = (bName == "Microsoft Internet Explorer" 
	&& bVer < 4);
	var blink_speed=1000;
	var i=0;
 
	if (NS4 || IE4) {
	if (navigator.appName == "Netscape") {
		layerStyleRef="layer.";
		layerRef="document.layers";
		styleSwitch="";
	} else {
		layerStyleRef="layer.style.";
		layerRef="document.all";
		styleSwitch=".style";
		}
	}

	//BLINKING
	function Blink(layerName){
		if (NS4 || IE4) { 
			if(i%2==0){
				eval(layerRef+'["'+layerName+'"]'+
				styleSwitch+'.visibility="visible"');
			} else {
				eval(layerRef+'["'+layerName+'"]'+
				styleSwitch+'.visibility="hidden"');
			}
		} else {
			var b_element = document.getElementsByTagName(layerName);
			if (i%2==0)
				b_element.style.visibility="visible";
			else
				b_element.style.visibility="hidden";
		}
		if(i<1){i++;} 
		else {i--}
		
		window.setTimeout("Blink('"+layerName+"')",blink_speed);
	}
	</script>
     <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />

    <!-- HTML5 shim and Respond.js IE8 support of HTML5 elements and media queries -->
    <!--[if lte IE 9]>
      <script src="Scripts/html5shiv.js"></script>
      <script src="Scripts/respond.min.js"></script>
     
      <link rel="stylesheet" href="css/survey ie-7.css" />
    <![endif]--> 
</head>
<body class="body">	
<form action="record.asp?surveycur=<%= EncryptString(surveycur) %>&resp_id=<%=resp_id%>" onSubmit="return validate(<%= q9 %><%= q10 %>);" method="post" name="theform" id="theform"> 	<!-- Builds form for passing data to "record.asp" page -->
           
      
     <div id="header" class="container-fluid">
                <div class="row">
                    <div class="col-md-3 col-xs-12 col-sm-12" >
   	                    <img src="images/logo.png" class="img-rounded" >
                    </div> 
                    <div class="learnMore">
                        <div class="navigation col-md-1 col-md-offset-8">                                                      
                                <a href="http://www.datassential.com" class="btn btn-info" target="_blank" role="button">Learn more</a>                            
                        </div>
                    </div>                 
                </div>
    
    </div>      
               
    <div class="container-fluid"> 
       
        <div class="row">
            <div id="main" class="col-md-6 col-md-offset-3">	
                
	            <div id="questionarea">
                     <div class="alert alert-danger hidden" role="alert">
                        <span id="error-Label"><strong>Warning!</strong></span>
                         
                    </div>
		    <script language="javascript" type="text/javascript">
		    <!-- 
			    String.prototype.trim = function() {		
			    // skip leading and trailing whitespace
			    // and return everything in between
			    var x=this;
			    x=x.replace(/^\s*(.*)/, "$1");
			    x=x.replace(/(.*?)\s*$/, "$1");
			    return x;
			    }			
		        function checkCR(evt) {
			         var evt  = (evt) ? evt : ((event) ? event : null);
			         var node = (evt.target) ? evt.target : ((evt.srcElement) ? evt.srcElement : null);
			         if ((evt.keyCode == 13) && (node.type=="text")) {
			 	        alert("Please click the Next button to continue.");
				        return false;
			         } else {
			             if (evt.keyCode == 34)
				             return false;
		            }
	              }
                document.onkeypress = checkCR;
			    var formVar = document.theform;
			    var validateArray = new Array();
			    var arraycounter = 0;
			    var tempvar;
			    function validate(){
			        errorstr = "";
			        $(".errorMsg").remove();
				    for(valCount=0; valCount < validateArray.length; valCount++){
					    tempvar = eval(validateArray[valCount]);

					    if(tempvar){
					        errorstr = errorstr + "<span class='errorMsg'>"+tempvar+"</span>";
					        
					    }
				    }
				    
				    if(errorstr != ""){
				        //alert(errorstr);
				        
				        $(".alert").removeClass("hidden").addClass("show");			
				        $("#error-Label").after(errorstr);

				        $("body, html").animate({
				            scrollTop: 0
				        }, 800, 'swing');
				        //window.scrollTo(0, 0);
				        //$(function(){
				        //    $(".alert").removeClass("hidden").addClass("show");				            
				        //});
				       
					    return false;
				    } else {
				        $(".alert").removeClass("show").addClass("hidden");		
					    formVar.submit();
				    }
			    }
			
			    //disables right click
			    /*var message="";

			    function clickIE()
			 
			    {if (document.all)
			    {(message);return false;}}
			 
			    function clickNS(e) {
			    if
			    (document.layers||(document.getElementById&&!document.all))
			    {
			    if (e.which==2||e.which==3) {(message);return false;}}}
			    if (document.layers)
			    {document.captureEvents(Event.MOUSEDOWN);document.  onmousedown=clickNS;}
			    else
			    {document.onmouseup=clickNS;document.oncontextmenu  =clickIE;}
			 
			    document.oncontextmenu=new Function("return false")*/
		     -->
		    </script>
		    <!-- Survey Questions go here: -->
		    <% 
			    Dim questionIDs
			    questionIDs =  ""
			    If IsTestingPhase Then
				    If surveycur = 1 Then%>
				    <!--<div id="toolTipBox" style="width: 250px; height: 100px; position: absolute; top: 0px; left: 50px; display: block; background-color: red">-->
				
                   <div class="jumbotron">
                       <h1>REMINDER</h1>
                       <p>Please be informed that this survey is currently in its 
						    <font style="font-weight:bold;color:red">TESTING PHASE</font>. 
						    All survey responses will be treated as <font style="font-weight:bold;color:red">TEST DATA ONLY</font>.
						    This reminder will be removed once the survey is <font style="font-weight:bold;color:red">OFFICIALLY LAUNCHED</font>.<br><br>
						    Thank you for your cooperation.</p>

                       <p style="text-align:left">Web Development Team</p>
                   </div>
				    <%
				    Else%>
                      
						    <div id="reminder" name="reminder">
						    TESTING PHASE ONLY
						    </div>
				    <br>
			    <%
				    End If
			    End If
			    If surveycur > 0 Then	
				    Dim NullError
				    NullError = cint(request.QueryString("Nullerror"))
				    if NullError = 1 then
				        response.Write "<p><b>There has been an error. Please update your browser.</b></p>"
				        response.Write "<p>Internet Explorer 7 and above users, kindly switch to COMPATIBILITY VIEW</p>"
				        response.Write "<p><u>Under the TOOLS menu, select the COMPATIBILITY MODE (if not checked) (TOOLS >> COMPATIBILITY MODE)</u></p>"
				        response.Write "<p>Kindly answer the question(s) on this page again</p>"
				    end if			
				    getSurveyQuestions_2 surveycur, false
				    if IsTestingPhase then
				        response.Write "Page: " & surveycur
				    end if
			    Else
                    'Uncomment if CONSUMER
                    'Dim passCode
                    'passCode = Session("subsid")
                    'Session.Abandon()

                    'comment if not OPERATOR
               
                    '---------------
                    Session.Abandon()

				    Select Case cInt(surveycur) %>
					    <%'MARC April 21, 2010 uncomment if consumer%>
					    <% Case -8	'completed already
    '						Session("uid") = ""	 'clears session id.%>
						    <p>Thank you for your interest but our record shows you have already taken the survey.</p>
					
					    <% Case -7	'Invalid page
						    Session("uid") = ""	 'clears session id.				%>
						    <p>Thank you for your interest but you have specified an invalid page or the page is not available.</p>
						    <p><b>Please <a href="index.html">start over</a> or <a href="mailto:webmaster@menus.com">email</a> the webmaster.</b></p>				
					    <% Case -6 'Successful end of survey:
						    Session("uid") = ""	 'clears session id.
						
						    'MARC April 21, 2010 uncomment if consumer
						    'response.redirect("http://sm1mr.com/ssred.php?S=1&ID=" & passCode)
										
						    %>

                             <div class="jumbotron">
                               <h2>
                                     <p>
                                        <%if custMsg = "" then %>
                                            Thank you for completing the survey.
                                        <%else %>
                                            <%=custMsg%>
                                        <%end if %>
                                   </p>
                                </h2>
                                   <!--remove if not operator-->
                                    <%if autoSnd = 1 then %>
                                        <%if isnull(amCode) then amCode = "" %>
                                        <%if amCode = "" then %>
                                            <p>Sorry but our system is currently overwhelmed with requests. Your gift certificate will be sent to you shortly.</p>
                                        <%else %>
                                            <p>Here is your Amazon Gift Certificate code <b><%=amCode%></b> (worth <b>$<%=amAmt%></b>). Your Amazon GC code has also been sent to your email address.</p>
                                        <%end if %>
                                    <%end if %>
                                    <!--=========-->
                                  <h4><small>Datassential Research Inc.</small></h4>       
                                              
                           </div>

                       
					    <% Case -5	'Survey is full:
						    Session("uid") = ""	 'clears session id.				
						
						    'MARC April 21, 2010 uncomment if consumer
						    'response.redirect("http://sm1mr.com/ssred.php?S=3&ID=" & passCode)
						    %>
						    <p>Thank you for your interest but the survey is now full.</p>
						
					    <% Case -4 'Unsuccessful end of survey - Disqualified:
						    Session("uid") = ""	 'clears session id.
						
						    'MARC April 21, 2010 uncomment if consumer
						    'response.redirect("http://sm1mr.com/ssred.php?S=2&ID=" & passCode)
						
						    %>         
                        <div class="bg-info">
                            <div class="jumbotron">
						        <p>Thank you for your interest and for taking the time to fill out these few questions.</p>
						        <p>Unfortunately, you do not qualify for this survey.</p>
						        <p>We look forward to your continued participation in future surveys!</p>
					        </div>   
                        </div>
					    <% Case -3	'Unsuccessful end of survey - Non Panel Member:
						    Session("uid") = ""		'clears session id.		%>
						    <p>Thank you for your interest in the survey, but this survey is only open to Panel members.</p>
						    <p>If you are a panel member then please <a href="javascript:history.back()">go back</a> and enter the email address that recieved the notice about this survey.</p>
						    <p>Thank You.</p>
						
					    <% Case -2	'Unsuccessfule end of survey - Outdated browser:		%>								
						    <b>You are using an old / unsupported browser.</b><br>
						    <b>We recommend upgrading to <a href="http://channels.netscape.com/ns/browsers/download.jsp" target="_blank">Netscape 7</a> or <a href="http://www.microsoft.com/windows/ie/default.asp" target="_blank">Internet Explorer 6</a>.</b>
				
					    <% Case -1	'Unsuccessfule end of survey - Error:		%>								
						    <p><b>There has been an error</b></p>
						    <p><b>Please <a href="index.html">start over</a> or <a href="mailto:webmaster@menus.com">email</a> the webmaster.</b></p>					
					    <% Case Else 'complete error %>				
						    <p><b>There has been an error</b></p>
						    <p><b>Please <a href="index.html">start over</a> or <a href="mailto:webmaster@menus.com">email</a> the webmaster.</b></p>										
			    <% End Select	
			    End If %>
		    <input type="hidden" name="questionIDs" value="<%= questionIDs %>">

            
            <div class="row footer-navigate">
                <%if surveycur >=0 then %>
                <div class="col-md-4 col-md-offset-8">
                    <button class="btn btn-info btn-block btn-lg" type="submit"><span class="glyphicon glyphicon-ok"></span> Next</button>
                    
                </div>
                <%end if %>

                <div id='tipDiv' style='position:absolute; visibility:hidden; z-index:100;'></div>

                 <script type="text/javascript">
                     /*Mouse over
                       image id must match image name.
                       enlarge version is scaled to 500px width no need to resize image unless the file size is too large. 
                     */
                     $(function()
                     {
		     
                         $("img").each(function(){		               
                             assign_tool_tip($(this).prop("id"), $(this).prop("id") + '.jpg');		               
                         });	       

		        
                     });
            
		    </script>

            </div>
           
    </div>    
            </div>
        </div>
    </div>

<footer class="text-center">All rights reserved <%=Year(Date) %></footer>
<%Response.write("<iframe name=""defibrillator"" width=""1%"" height=""1%"" src=""defibrillator.asp"" style=""display:none""></iframe>")%>
</form>
</body>
</html>
		