<% 
Function RADIO_HTML( var_name, q_text, e_text, options, is_random, is_other, is_none)
	privateRADIOBUILD_HTML var_name, q_text, e_text, options, is_random, is_other, is_none, "list", false
End Function

Function RADIO_VALUES_HTML( var_name, q_text, e_text, options, is_random, is_other, is_none)
	privateRADIOBUILD_HTML var_name, q_text, e_text, options, is_random, is_other, is_none, "list", true
End Function

Function RADIOCOLUMN_HTML(var_name, q_text, e_text, options, is_random, is_other, is_none)
	privateRADIOBUILD_HTML var_name, q_text, e_text, options, is_random, is_other, is_none, "column", false
End Function

Function RADIOCOLUMN_VALUES_HTML(var_name, q_text, e_text, options, is_random, is_other, is_none)
	privateRADIOBUILD_HTML var_name, q_text, e_text, options, is_random, is_other, is_none, "column", true
End Function

Function privateRADIOBUILD_HTML( var_name, q_text, e_text, options, is_random, is_other, is_none, displaytype, is_values ) 
	Dim radioOtherCustom
	radioOtherCustom = false
	select case var_name
		case "VARNAME"
			radioOtherCustom = true
	end select
%>
	<script language="javascript" type="text/javascript">
		function validate_<%=var_name%>(){
			var returnstr = "";
			if ( !checkRadio(formVar.<%=var_name%>) ){
				returnstr = returnstr + "<%=e_text%>\n\n";
			}
			<% if is_other then %>
			if((getRadio(formVar.<%=var_name%>) == formVar.otherValue_<%=var_name%>.value) && (!checkInput(formVar.<%=var_name%>_other.value))){
				returnstr = returnstr + "Please specify where you selected 'Other'\n\n";
			}
			<% end if %>
			if (returnstr != ""){
			    $("#table<%=var_name%> tr").addClass("validation-error");
			    return returnstr;
			} else {
			    $("#table<%=var_name%> tr").removeClass("validation-error");
			    return false;
			}
		}
		<% if is_other then %>
			function enableOther_<%=var_name%>(otherVal){
				if(getRadio(formVar.<%=var_name%>)==otherVal){
					formVar.<%=var_name%>_other.disabled = false;
					formVar.<%=var_name%>_other.focus(); 
					<%if radioOtherCustom then%>
						formVar.<%=var_name%>_other_2.value = "";
						formVar.<%=var_name%>_other_2.disabled = true;
					<%end if%>
				} 
			}
		<%if radioOtherCustom then%>
			function enableOther_<%=var_name%>_2(otherVal)
			{
				clearOther_<%=var_name%>();
				if(getRadio(formVar.<%=var_name%>)==otherVal)
				{
					formVar.<%=var_name%>_other_2.disabled = false;
					formVar.<%=var_name%>_other_2.focus(); 
				}
			}
		<%end if%>
			function clearOther_<%=var_name%>(){
				formVar.<%=var_name%>_other.value = "";
				formVar.<%=var_name%>_other.disabled = true;
				<%if radioOtherCustom then%>
					formVar.<%=var_name%>_other_2.value = "";
					formVar.<%=var_name%>_other_2.disabled = true;
				<%end if%>
			}
		<% end if %>
		validateArray[arraycounter] = "validate_<%=var_name%>();";
		arraycounter++;
	</script>
	
	<% DIV_HTML var_name, q_text, "question" %>
	
	<div id="diva<%=var_name%>" class="answer">
	<table width="100%"  border="0" cellspacing="0" cellpadding="5" class="radioTable" id="table<%=var_name%>">
		<% 	If displaytype = "column" Then %><tr><% End If %>
		<% 	ary = Split(options,"|")
			If is_random Then
				randArray = buildRandArrayReturn(UBound(ary)+1)
			End If  
			displayCounter = 0
			For i = 1 To UBound(ary)+1
				text = ary(i-1)
				If is_values Then
					val = text
				Else
					val = i
				End If
				If is_random Then
					text = ary(randArray(i)-1)
					If is_values Then
						val = text
					Else
						val = randArray(i)
					End If
				End If 
				If Session(var_name & "|" & i) <> "false" Then
					displayCounter = displayCounter + 1
					Response.Write ("<!--" & displayCounter & "-->")
				End If
				if displayCounter mod 2 = 0 then
					rowClass = "radioTableRow1"
				else
					rowClass = "radioTableRow2"
				end if
					%> 	    
				<% 	If Session(var_name & "|" & val) <> "false" Then %>
					<% 	If displaytype = "list" Then %><tr class="<%= rowClass %>"  id="rowo<%=var_name&"|"&val%>"><% End If %>
						<td valign="top" class="radioTableCellInput" id="celli<%=var_name & "|" & i%>" <% 	If displaytype = "column" Then %>align="center"<% End If %>><input type="radio" value="<%=val%>" id="<%=var_name & "|" & val%>" name="<%=var_name%>" <% if is_other then %> onClick="clearOther_<%=var_name%>();" <% end if %> class="radioInput" />
						<% If displaytype = "list" Then %></td><td valign="top" class="radioTableCellOption" id="cello<%=var_name & "|" & val%>"><% ElseIf displaytype = "column" Then %><br><% End If %>
						<label for="<%=var_name & "|" & val%>"><%=text%></label></td>
					<% 	If displaytype = "list" Then %></tr><% End If %>
				<% 	End If %>
		<% 	Next %>
		<% 	questionIDs = questionIDs & var_name &"|" %>
		<% 	rowCount = displayCounter + 1 %>
		<% 	if is_other then %> 
		<%		if rowCount mod 2 = 0 then
					rowClass = "radioTableRowOther1"
				else
					rowClass = "radioTableRowOther2"
				end if	%>     
				<% 	If displaytype = "list" Then %><tr class="<%= rowClass %>"  id="rowo<%=var_name&"_"&rowCount%>"><% End If %>
					<td valign="top" class="radioTableCellInputOther" id="celli<%=var_name&"_"&rowCount%>"><input type="radio" value="<%=rowCount%>" id="<%=var_name & "|" & rowCount%>" name="<%=var_name%>"  onClick="enableOther_<%=var_name%>(this.value);" class="radioInput" /><input type="hidden" value="<%=rowCount%>" id="otherValue_<%=var_name%>">
					<% If displaytype = "list" Then %></td><td valign="top" class="radioTableCellOptionOther" id="cello<%=var_name&"_"&rowCount%>"><% ElseIf displaytype = "column" Then %><br><% End If %>
					<label for="<%=var_name & "|" & rowCount%>">Other: <input type="text" id="<%=var_name%>_other" name="<%=var_name%>_other" value="" disabled class="radioInputOther"> (please specify)</label></td>
				<% 	If displaytype = "list" Then %></tr><% End If %>
				<% questionIDs = questionIDs & var_name &"_other|" %>
				<% rowCount = rowCount + 1 %>
				<%if radioOtherCustom then%>
				<%		if rowCount mod 2 = 0 then
							rowClass = "radioTableRowOther1"
						else
							rowClass = "radioTableRowOther2"
						end if	%>
						<tr class="<%= rowClass %>">
							<td valign="top" class="radioTableCellInputOther">
								<input type="radio" value="<%=rowCount%>" id="<%=var_name & "|" & rowCount%>" name="<%=var_name%>"  onClick="enableOther_<%=var_name%>_2(this.value);" class="radioInput" />
								<input type="hidden" value="<%=rowCount%>" id="otherValue_<%=var_name%>_2" NAME="otherValue_<%=var_name%>_2">
							</td>
							<td valign="top" class="radioTableCellOptionOther">
								<label for="<%=var_name & "|" & rowCount%>">
									Other <input type="text" id="<%=var_name%>_other_2" name="<%=var_name%>_other_2" value="" disabled class="radioInputOther"> (please specify): </label>
							</td>
						</tr>
					<% questionIDs = questionIDs & var_name &"_other_2|" %>
					<% 	rowCount = displayCounter + 1 %>
				<%end if%>
		<% end if %>
		<% if is_none then %>
		<%		if rowCount mod 2 = 0 then
					rowClass = "radioTableRowNone1"
				else
					rowClass = "radioTableRowNone2"
				end if	%>  
			<!-- none --> 	     
				<% 	If displaytype = "list" Then %><tr class="<%= rowClass %>"  id="rowo<%=var_name&"_"&rowCount%>"><% End If %>
					<td valign="top" class="radioTableCellInputNone" id="celli<%=var_name&"_"&rowCount%>"><input type="radio" value="<%=rowCount%>" id="<%=var_name & "|" & rowCount%>" name="<%=var_name%>" <% if is_other then %> onClick="clearOther_<%=var_name%>();" <% end if %> class="radioInput" />
					<% If displaytype = "list" Then %></td><td valign="top" class="radioTableCellOptionNone" id="cello<%=var_name&"_"&rowCount%>"><% ElseIf displaytype = "column" Then %><br><% End If %>
					<label for="<%=var_name & "|" & rowCount%>">None of these</label></td>
				<% 	If displaytype = "list" Then %></tr><% End If %>
		<% end if %>
		<% 	If displaytype = "column" Then %></tr><% End If %>
		</table>
	</div>
<% End Function %>