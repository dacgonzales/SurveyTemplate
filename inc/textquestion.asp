<%
Function TEXTQUESTION_HTML(var_name, q_text, e_text, options, t_size, t_max)
	privateTEXTQUESTION_HTML var_name, q_text, e_text, options, t_size, t_max, ""
End Function

Function TEXTQUESTION_DISABLED_HTML(var_name, q_text, e_text, options, t_size, t_max, labelString)
	privateTEXTQUESTION_HTML var_name, q_text, e_text, options, t_size, t_max, labelString
End Function

Function privateTEXTQUESTION_HTML( var_name, q_text, e_text, options, t_size, t_max, labelString ) %>
		<% DIV_HTML "q"&var_name, q_text, "question" %>
	
	<div id="diva<%=var_name%>" class="answer">
	<table cellpadding="5" cellspacing="0" border="0" class="textquestionTable" id="table<%=var_name%>">
	<% 	if isNull(options) or options = "" then
			dim ary(0)
			ary(0) = " "
		else
			ary = split(options, "|")
		end if
		if isNull(e_text) or e_text = "" then
			dim ary_e(0)
			ary_e(0) = ""
			no_error = true
		else
			ary_e = split(e_text,"|")
			no_error = false
		end if
				
		For i = 0 to ubound(ary)
			If Session(var_name&"_"&(i+1)) <> "false" Then
				if ubound(ary) <> 0 Then
					val = var_name & "_" & (i+1)
				Else
					val = var_name
				End If
				If no_error <> true then
					vale = ary_e(i)
					If vale <> "" Then 
						valeArray = split(vale,"^") 
						validationType = valeArray(0)
						errorStr = valeArray(1) 
						%>
						<script language="javascript" type="text/javascript">
							function validate_<%=val%>(){
								returnStr = "";
								// validate only if e_text is passed
								item2check = formVar.<%=val%>.value;
								<%
								If labelString <> "" Then
									radioVal = "&& (!(formVar." & val & "_r.checked))"
								Else
									radioVal = ""
								End If
								Select Case validationType %>
									<%	Case "Text" %>		if ( ("<%=errorStr%>"!='') && (!checkInput(item2check)) <%= radioVal %>) returnStr += "<%=errorStr%>\n\n";
									<%	Case "Email" %>		if ( ("<%=errorStr%>"!='') && (!checkEmail(item2check)) <%= radioVal %>) returnStr += "<%=errorStr%>\n\n";
									<%	Case "Phone" %>		if ( ("<%=errorStr%>"!='') && (!checkPhone(item2check)) <%= radioVal %>) returnStr += "<%=errorStr%>\n\n";
									<%	Case "Number" %>	if ( ("<%=errorStr%>"!='') && (!checkNumber(item2check) || (item2check == "")) <%= radioVal %>) returnStr += "<%=errorStr%>\n\n";
									<%	Case "Zip" %>	if ( ("<%=errorStr%>"!='') && (!checkZip(item2check)) <%= radioVal %>) returnStr += "<%=errorStr%>\n\n";
									<%	Case "WholeNumber"%> 
                                        if ( ("<%=errorStr%>"!='') && (!checkWholeNumber(item2check) || (item2check == "")) <%= radioVal %>) returnStr += "<%=errorStr%>\n\n";
                                       
									<%	Case "RealNumber"%> if ( ("<%=errorStr%>"!='') && (!checkRealNumber(item2check) || (item2check == "")) <%= radioVal %>) returnStr += "<%=errorStr%>\n\n";
									<% 	Case Else %>	if ( ("<%=errorStr%>"!='') && (<%=validationType%>) <%= radioVal %>) returnStr += "<%=errorStr%>\n\n";
								<% 
								End Select %>
								if(returnStr != ""){
								    $("#r-<%=val%>").addClass("validation-error");
									return returnStr;
								} else {
									$("#r-<%=val%>").removeClass("validation-error");
									<%If labelString <> "" Then%>
										if(formVar.<%=val%>_r.checked){
											formVar.<%=val%>.disabled = false;
											formVar.<%=val%>.value = formVar.<%=val%>_r.value;
										}
										<%End If%>
                                    
									return false;
								}
							}
							validateArray[arraycounter] = "validate_<%=val%>();";
							arraycounter++;
							function disableInput_<%=val%>(item){
								if(item.checked){
									formVar.<%=val%>.value = "";
									formVar.<%=val%>.disabled = true;
								} else {
									formVar.<%=val%>.value = "";
									formVar.<%=val%>.disabled = false;
								}
							}
						</script>	
					<% 
					End If
				End If
				textArray = split(ary(i),"^")
				If i mod 2 = 0 Then
					rowClass = "textquestionTableRow1"			
				Else
					rowClass = "textquestionTableRow2"
				End If 
				%>
				<tr class="<%= rowClass %>"  id="r-<%=val%>">
					<td class="textquestionTableCellOption" id="cello<%=val%>"><%=textArray(0)%></td>
					<td class="textquestionTableCellInput" id="celli<%=val%>">
						<%if var_name = "VARNAME" then%>
							<textarea name="<%=val%>" id="<%=val%>" class="textquestionInput" value="" rows="5"></textarea>
						<%else%>
							<input type="text" name="<%=val%>" id="<%=val%>" size="<%=t_size%>" maxlength="<%=t_max%>" class="textquestionInput" value=""><% If ubound(textArray) > 0 Then%><span class="textquestionInputUnit"><%=textArray(1)%></span><% End If %>
						<%end if%>
					</td>
					<%If labelString <> "" Then%>
						<td class="textquestionTableCellRadio"><input type="checkbox" name="<%=val%>_r" id="<%=val%>_r" value="<%=labelString%>" onChange="disableInput_<%=val%>(this);"><%= labelString %></td>
					<%End If%>
				</tr>
				<% questionIDs = questionIDs & val &"|" %>
				<%
			End If
		Next	
	%>
	</table>
	</div>
<% End Function %>