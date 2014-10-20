<% ' 08.04.04 ejohnson <-- Added Randomization and None of the Above functionality %>
<% ' 06.28.05 ejohnson <-- Added Checkbox ROWS function %>
<% 
Function CHECKBOX_HTML( var_name, q_text, e_text, options, is_random, is_other, is_none ) 
	privateCHECKBOX_HTML var_name, q_text, e_text, options, is_random, is_other, is_none, "", false, "", "", "", "" 
End Function

Function CHECKBOX_INPUT_HTML( var_name, q_text, e_text, options, is_random, is_other, is_none, labelstring )
	privateCHECKBOX_HTML var_name, q_text, e_text, options, is_random, is_other, is_none, "", false, "", "", "" , labelstring
End Function

Function CHECKBOX_DISABLED_HTML( var_name, q_text, e_text, options, is_random, is_other, is_none, labelstring ) 
	privateCHECKBOX_HTML var_name, q_text, e_text, options, is_random, is_other, is_none, "", false, "", "", labelstring, "" 
End Function

Function CHECKBOX_COLUMNS_HTML( var_name, q_text, e_text, options, is_random, is_other, is_none, values ) 
	privateCHECKBOX_HTML var_name, q_text, e_text, options, is_random, is_other, is_none, "", false, values , "columns", "", ""
End Function

Function CHECKBOX_ROWS_HTML( var_name, q_text, e_text, options, is_random, is_other, is_none, values ) 
	privateCHECKBOX_HTML var_name, q_text, e_text, options, is_random, is_other, is_none, "", false, values, "rows" , "", ""
End Function

Function CHECKBOX_MAX_HTML( var_name, q_text, e_text, options, is_random, is_other, is_none, endNum ) 
	privateCHECKBOX_HTML var_name, q_text, e_text, options, is_random, is_other, is_none, endNum, false, "", "" , "", ""
End Function

Function CHECKBOX_VALUES_HTML( var_name, q_text, e_text, options, is_random, is_other, is_none ) 
	privateCHECKBOX_HTML var_name, q_text, e_text, options, is_random, is_other, is_none, "", true, "", "", "", "" 
End Function

Function CHECKBOX_VALUES_MAX_HTML( var_name, q_text, e_text, options, is_random, is_other, is_none, endNum )
	privateCHECKBOX_HTML var_name, q_text, e_text, options, is_random, is_other, is_none, endNum, true, "", "" , "", ""
End Function

Function privateCHECKBOX_HTML( var_name, q_text, e_text, options, is_random, is_other, is_none, endNum, is_values, values, columns_or_rows, none_labelstring, inputlabel ) 
	Dim customCheckboxOther, other_labelString, customNone
    other_labelString = "Other, please specify: "
    customNone = false
	customCheckboxOther = false
	select case var_name
		case "VARNAME"
			customCheckboxOther = true
        
	end select
	%>
	<!-- note: validation is below -->
	<% DIV_HTML "q"&var_name, q_text, "question" %>
	<!-- answer -->
	<div id="diva<%=var_name%>" class="answer">
		<table width="100%"  border="0" cellspacing="0" cellpadding="5" class="checkboxTable" id="table<%=var_name%>">
		<% 
			If values <> "" Then
				columnArray = Split(values, "|")
				widthVar = ubound(columnArray) + 1
				If columns_or_rows = "rows" Then
					If is_none Then widthVar = widthVar + 1 End If
					If is_other Then widthVar = widthVar + 1 End If
				End If
				%>
				<tr>
					<td class="checkboxTableCellColumnHead" width="50%">&nbsp;</td>
					<% For columnNum = 1 to Ubound(columnArray)+1 %>
					<td class="checkboxTableCellColumnHead" id="cello<%=var_name&"_"&columnNum%>" width="<%=50/widthVar%>%" align="center"><%=columnArray(columnNum-1)%></td>
					<% Next %>					
					<% If columns_or_rows = "rows" And is_other Then %>
						<td class="checkboxTableCellColumnHead" id="cello_Other" align="center" width="<%=50/widthVar%>%">Other</td>
					<% End If %>
					<% If columns_or_rows = "rows" And is_none Then %>
						<td class="checkboxTableCellColumnHead" id="cello_NONE" align="center" width="<%=50/widthVar%>%"><%	If none_labelstring = "" Then %>None<% Else %><%=none_labelstring%><% End If%></td>
					<% End If %>			
				</tr>
				<%
			Else
				Dim columnArray(0)
				columnArray(0) = ""
				If inputlabel <> "" Then
					%><tr>
						<td>&nbsp;
							<script language="javascript" type="text/javascript">
							<!-- 
							function enableInput<%=var_name%>(checkbox,target){
								if(checkbox.checked){
									target.disabled = false;
									target.focus();
								} else {
									target.value = "";
									target.disabled = true;
								}
							}
							 -->
							</script>
						</td>
						<td>&nbsp;</td>
						<td align="center"><%=inputlabel%></td>
					  </tr>
					<%
				End If
			End If
		%>
		<% ary = Split(options,"|") 
			If is_random Then
			   randArray = buildRandArrayReturn(UBound(ary)+1)
			End If  
			rowCount = 0
			for i = 1 to UBound(ary)+1
				text = ary(i-1)
				val = i
				If is_random Then
					val = randArray(i)
					text = ary(randArray(i)-1)
				End If 
				v = var_name & "_" & (val) 
				if rowCount mod 2 = 0 then
					rowClass = "checkboxTableRow1"
				else
					rowClass = "checkboxTableRow2"
				end if
				If is_values Then
					checkBoxVal = text
				Else
					checkBoxVal = 1
				End If
				If Session(var_name&"_"&val) <> "false" Then
					rowCount = rowCount + 1
				%> 	  
			<tr class="<%=rowClass %>" id="row<%=var_name&"_"&val%>">
				<td class="checkboxTableCellOption" id="cello<%=var_name&"_"&val%>"><%=rowCount%>. <label for='<%=v%>'><%=text%></label></td>

				<% If values <> "" Then %>
					<% If columns_or_rows = "columns" Then %>

						<% For x = 0 to Ubound(columnArray)
							columnNum = x+1                            
							v = var_name & "_" & columnNum & "_" & val %>
							<td class="checkboxColumnsTableCellInput" id="celli<%=v%>" align="center"><input type="checkbox" id='<%=v%>' name='<%=v%>' value="<%= checkBoxVal %>" <% if is_none Then%>onclick="clearNone_<%=var_name&"_"&columnNum%>();"<% end if %> class="checkboxInput" /></td>  
							<% questionIDs = questionIDs & v &"|" %>
							<script language="javascript">
							function validateSub_<%=var_name%>_<%=columnNum%>(){
								anyChecked = 0;
								for(i = 1; i <= <%= UBound(ary)+1%>; i++){		
								    $("#celli<%=var_name & "_" & columnNum & "_"%>" + i).removeClass("validation-error");    
								    <% If Session(var_name&"_"&val) <> "false" Then %>
                                        
										if(eval("formVar.<%=var_name%>_<%=columnNum%>_" + i + ".checked") == true)anyChecked++;				
									<% End If %>
								}
								<% If is_other Then %>
									if(eval("formVar.<%=var_name%>_<%=columnNum%>_otherbox.checked") == true){									      
									    anyChecked++;
									}
								<% End If %>
								<% If is_none Then %>
									if(eval("formVar.<%=var_name%>_<%=columnNum%>_none.checked") == true)	anyChecked++;
									<% End If %>

								<% if endNum <> "" then %>
								    if(anyChecked != <%=endNum%>){
									    returnStr += "  <%=columnNum%>";
								    }	
							    <% else %>
                                    
								if(anyChecked == 0)
							    {							        
								    for(ii = 1; ii <= <%= UBound(ary)+1%>; ii++){
								        $("#celli<%=var_name & "_" & columnNum & "_"%>" + ii).addClass("validation-error");        
								    }
								    returnStr += "  <%=columnNum%>";	
								}
								<% end if %>
							}
							</script>
						<% Next %>
					<% ElseIf columns_or_rows = "rows" Then %>
						<% For x = 0 to Ubound(columnArray)
							columnNum = x+1
							v = var_name & "_" & val & "_" & columnNum %>
							<td class="checkboxColumnsTableCellInput" id="celli<%=v%>" align="center"><input type="checkbox" id='<%=v%>' name='<%=v%>' value="<%= checkBoxVal %>" <% if is_none Then%>onclick="clearNone_<%=var_name&"_"&val%>();"<% end if %> class="checkboxInput" /></td>  
							<% questionIDs = questionIDs & v &"|" %>
						<% Next %>
						<% If is_other Then %>
							<% v = var_name & "_" & val %>
							<td class="checkboxRowsTableCellInputOther" id="celli<%=v%>_other" nowrap align="center"><input type="checkbox" id='<%=v%>_otherbox' name='<%=v%>_otherbox' value="1" onClick="enableOther_<%=v%>(this);" class="checkboxInput" /><input type="text" id="<%=v%>_other" name="<%=v%>_other" value="" disabled class="checkboxOther"></td>
							<% questionIDs = questionIDs & v &"_other|" %>
							<script language="javascript">
							function enableOther_<%=v%>(item){
								if(item.checked){
									formVar.<%=v%>_other.disabled = false;
									formVar.<%=v%>_other.focus(); 
									<% if is_none Then%>clearNone_<%=v%>();<% end if %> 
								} else {
									formVar.<%=v%>_other.value = "";
									formVar.<%=v%>_other.disabled = true;
								}
							}
							function checkOther_<%=v%>(){
								if((formVar.<%=v%>_otherbox.checked) && (!checkInput(formVar.<%=v%>_other.value))) returnStr += "Please specify other in row <%=val%>.\n\n";
							}
							</script>
						<% End If %>
						<% If is_none Then %>
							<% v = var_name & "_" & val & "_none" %>
							<td class="checkboxRowsTableCellInputNone" id="celli<%=v%>" align="center"><input type="checkbox" id="<%=v%>" name="<%=v%>" value="1" onClick="clearAll_<%=var_name&"_"&val%>(this);" class="checkboxInput" /></td>
							<% questionIDs = questionIDs & v & "|" %>
							<script language="javascript">					
							function clearNone_<%=var_name&"_"&val%>(){
								formVar.<%=var_name&"_"&val%>_none.checked = false;
							}	
							function clearAll_<%=var_name&"_"&val%>(item){
								if(item.checked){
									for(i=1; i <= <%=(uBound(columnArray)+1)%>;i++){
										eval("formVar.<%=var_name&"_"&val%>_" + i + ".checked = false");
										
									}
									<% if is_other Then %>
									formVar.<%=var_name&"_"&val%>_otherbox.checked = false;
									enableOther_<%=var_name&"_"&val%>(formVar.<%=var_name&"_"&val%>_otherbox);
									<% end if %>
								}
							}
							</script>
						<% End If %>
						<script language="javascript">
						function validateSub_<%=var_name%>_<%=val%>(){
							anyChecked = 0;
							for(i = 1; i <= <%= UBound(columnArray)+1%>; i++){	
									if(eval("formVar.<%=var_name%>_<%=val%>_" + i + ".checked") == true)anyChecked++;				
							}
							<% If is_other Then %>
								if(eval("formVar.<%=var_name%>_<%=val%>_otherbox.checked") == true)	anyChecked++;
							<% End If %>
							<% If is_none Then %>
								if(eval("formVar.<%=var_name%>_<%=val%>_none.checked") == true)	anyChecked++;
							<% End If %>
							<% if endNum <> "" then %>
							if(anyChecked != <%=endNum%>){
								returnStr += "  <%=rowCount%>";
							}	
							<% else %>
							if(anyChecked == 0)	returnStr += "  <%=rowCount%>";		
							<% end if %>
						}
						</script>
					<% End If %>
				<% Else %>
					<td class="checkboxTableCellInput" id="celli<%=v%>"><input type="checkbox" id='<%=v%>' name='<%=v%>' value="<%= checkBoxVal %>" onclick="<% If inputlabel <> "" Then %>enableInput<%=var_name%>(formVar.<%=v%>,formVar.<%=v%>input);<% End If %><% if is_none Then%>clearNone_<%=var_name%>();<% end if %>" class="checkboxInput" /></td>  
					<% questionIDs = questionIDs & v &"|" %>
					<% If inputlabel <> "" Then %>
						<td><input type="text" id="<%=v & "input"%>" name="<%=v & "input"%>" value="" disabled></td>
						<% questionIDs = questionIDs & v & "input" & "|" %>
					<% End If %>
				<% End If %>
			</tr>
			<% End If %>
		<% next %>
		<% 
			'rowCount = i
			if is_other and columns_or_rows <> "rows" then 
				v = var_name & "_other"
				if rowCount mod 2 = 0 then
					rowClass = "checkboxTableRowOther1"
				else
					rowClass = "checkboxTableRowOther2"
				end if
				%> 	    
				<% rowCount = rowCount + 1 %>
			<tr class="<%= rowClass %>" id="row<%=var_name&"_"&rowCount%>">
				<td <%if values = "" then %>colspan = "2"<%end if %> class="checkboxTableCellOptionOther" id="cello<%=var_name&"_"&rowCount%>">
                    <label for="<%=var_name%>_otherbox"> <%= rowCount%>. <%if customCheckboxOther then%>Other manufacturer brand (please specify):<%else%><%=other_labelString %><%end if%></label>
                   

                  <%if values = "" then %>                    
                    <div style="display:inline;float:right;padding:3px 0 3px 0;">
                        <input type="checkbox" id='<%=var_name%>_otherbox' name='<%=var_name%>_otherbox' value="1" onClick="enableOther_<%=var_name%>(this);" class="checkboxInput" style="padding-right:10px"/>
                        <input type="text" id="<%=v%>" name="<%=v%>" value="" disabled style="padding-top:5px" class="checkboxOther">
                     </div> 
                    <%end if %>
				</td>

				<% If values <> "" Then %>
					<% For x = 0 to Ubound(columnArray)
						columnNum = x+1
						v = var_name & "_" & columnNum & "_other" %>
						<td class="checkboxTableCellInputOther" id="celli<%=v%>" nowrap><input type="checkbox" id='<%=var_name&"_"&columnNum%>_otherbox' name='<%=var_name&"_"&columnNum%>_otherbox' value="1" onClick="enableOther_<%=var_name&"_"&columnNum%>(this);" class="checkboxInput" /><input type="text" id="<%=v%>" name="<%=v%>" value="" disabled class="checkboxOther"></td>
						<% questionIDs = questionIDs & v &"|" %>
						<script language="javascript">
						function enableOther_<%=var_name&"_"&columnNum%>(item){
							if(item.checked){
								formVar.<%=v%>.disabled = false;
								formVar.<%=v%>.focus(); 
								<% if is_none Then%>clearNone();<% end if %> 
							} else {
								formVar.<%=v%>.value = "";
								formVar.<%=v%>.disabled = true;
							}
						}
						function checkOther_<%=var_name&"_"&columnNum%>(){
						    if((formVar.<%=var_name&"_"&columnNum%>_otherbox.checked) && (!checkInput(formVar.<%=v%>.value))){
						        $("#celli<%=var_name & "_" & columnNum & "_other"%>").addClass("validation-error");  
						        returnStr += "Please specify other in column <%=columnNum%>.\n\n";
						    }
						        
						}
						</script>
					<% Next %>
				<% Else %>
                    <!--<td></td>-->
					<!--<td class="checkboxTableCellInputOther" id="celli<%=v%>" nowrap>
                        <input type="checkbox" id='<%=var_name%>_otherbox' name='<%=var_name%>_otherbox' value="1" onClick="enableOther_<%=var_name%>(this);" style="display:inline" class="checkboxInput" />  
                        <input type="text" id="<%=v%>" name="<%=v%>" value="" disabled style="padding-top:5px;display:inline-block" class="checkboxOther">
					</td>-->
					<% questionIDs = questionIDs & v &"|" %>
				<% End If %>
			</tr>
			
			<%if customCheckboxOther then%>
			<%
				rowCount = rowCount + 2
				'response.write rowCount
				'response.end
				if rowCount mod 2 = 0 then
					rowClass = "checkboxTableRowOther1"
				else
					rowClass = "checkboxTableRowOther2"
				end if
				v = var_name & "_other_2"
			%>
				<tr class="<%= rowClass %>">
					<td class="checkboxTableCellOptionOther">
						<%=rowCount-1%>. <%if var_name = "qQD1" then%>Other<%end if%> Distributor brand (please specify):
					</td>
					<td class="checkboxTableCellInputOther" id="celli<%=v%>" nowrap>
						<input type="checkbox" id='<%=var_name%>_otherbox_2' name='<%=var_name%>_otherbox_2' value="1" onClick="enableOther_<%=var_name%>_2(this);" class="checkboxInput" />
						<input type="text" id="<%=v%>" name="<%=v%>" value="" disabled class="checkboxOther">
						</td>
					<% questionIDs = questionIDs & v &"|" %>
				</tr>
			<%end if%>
			
		<% end if %>
		<% if is_none and columns_or_rows <> "rows" then 
				if rowCount mod 2 = 0 then
					rowClass = "checkboxTableRowNone1"
				else
					rowClass = "checkboxTableRowNone2"
				end if
				rowCount = rowCount + 1
				%> 	    
			<tr class="<%= rowClass %>" id="row<%=var_name&"_"&rowCount%>">
				<td class="checkboxTableCellOptionNone" id="cello<%=var_name&"_"&rowCount%>"><label for="<%=var_name%>_none"><%= rowCount%>. <%	If none_labelstring = "" Then %>None of the above<% Else %><%=none_labelstring%><% End If%></label><%=columns_or_rows%></td>
				<% If values <> "" Then %>
					<% For x = 0 to Ubound(columnArray)
						columnNum = x+1
						v = var_name & "_" & columnNum & "_none" %>
						<td class="checkboxTableCellInputNone" id="celli<%=v%>"><input type="checkbox" id="<%=v%>" name="<%=v%>" value="1" onClick="clearAll_<%=var_name&"_"&columnNum%>(this);" class="checkboxInput" /></td>
						<% questionIDs = questionIDs & v & "|" %>
						<script language="javascript">					
						function clearNone_<%=var_name&"_"&columnNum%>(){
							formVar.<%=var_name&"_"&columnNum%>_none.checked = false;
						}	
						function clearAll_<%=var_name&"_"&columnNum%>(item){
							if(item.checked){
								for(i=1; i < <%=i%>;i++){
									eval("formVar.<%=var_name&"_"&columnNum%>_" + i + ".checked = false");
								}
								<% if is_other Then %>
								formVar.<%=var_name&"_"&columnNum%>_otherbox.checked = false;
								enableOther_<%=var_name&"_"&columnNum%>(formVar.<%=var_name&"_"&columnNum%>_otherbox);
								<% end if %>
							}
						}
						</script>
					<% Next %>
				<% Else %>
					<td class="checkboxTableCellInputNone" id="celli<%=var_name&"_"&rowCount%>"><input type="checkbox" id="<%=var_name%>_none" name="<%=var_name%>_none" value="1" onClick="clearAll_<%=var_name%>(this);" class="checkboxInput" /></td>
					<% questionIDs = questionIDs & var_name &"_none|" %>
				<% End If %>
			</tr>
		<% end if %>

        <%if customNone then
            rowCount = rowCount + 1
            %>
            <tr><td><label for="<%=var_name%>_none2"><%= rowCount%>. <%	If none_labelstring = "" Then %>None of the above<% Else %><%=none_labelstring2%><% End If%></label></td>
                <td>                    
                    <input type="checkbox" id="<%=var_name%>_none2" name="<%=var_name%>_none2" value="1" onClick="clearAll2_<%=var_name%>(this);" class="checkboxInput" />
                </td>
                <% questionIDs = questionIDs & var_name &"_none2|" %>
            </tr>
        <% end if %>
		</table>
	</div>
	
	<script language="javascript" type="text/javascript">
		function validate_<%=var_name%>(){
		    returnStr = "";
		    $("#table<%=var_name%> tr").removeClass("validation-error");
			<% if e_text <> "" then %>
				anyChecked = 0;
				<% If values <> "" Then %>
					<% If columns_or_rows = "columns" Then %>
						<% For columnNum = 1 to ubound(columnArray) + 1 %>                            
							validateSub_<%=var_name%>_<%=columnNum%>();
						<% Next %>
						if(returnStr != ""){
							returnStr = "<%=e_text%>\n" + returnStr + "\n\n";
						}
					<% ElseIf columns_or_rows = "rows" Then %>
						<% For rowNum = 1 to ubound(ary) + 1 
								val = rowNum
								If is_random Then
									val = randArray(rowNum)
								End If %>
							<% If Session(var_name & "_" & val) <> "false" Then %>
								validateSub_<%=var_name%>_<%=val%>();
							<% End If %>
						<% Next %>
						if(returnStr != ""){
							returnStr = "<%=e_text%>\n" + returnStr + "\n\n";
						}
					<% End If %>
				<% Else %>

					<% For i = 1 to uBound(ary) + 1 %>
                         
						<% If Session(var_name&"_"&i) <> "false" Then %>
							if(eval("formVar.<%=var_name%>_<%= i %>.checked") == true)anyChecked++;	
						<% End If %>
					<% Next %>
					<% If is_other Then %>
						if(eval("formVar.<%=var_name%>_otherbox.checked") == true)	anyChecked++;
					<% End If %>
					<%if customCheckboxOther then%>
						if(eval("formVar.<%=var_name%>_otherbox_2.checked") == true)	anyChecked++;
					<%end if%>
					<% If is_none Then %>
						if(eval("formVar.<%=var_name%>_none.checked") == true)	anyChecked++;
						<% End If %>
                    <% If customNone Then %>
						if(eval("formVar.<%=var_name%>_none2.checked") == true)	anyChecked++;
                        <% End If %>
					<% if endNum <> "" then %>
                    
					if(anyChecked > <%=endNum%> || anyChecked == 0){
						returnStr = "<%=e_text%>\n\n";
					}
					<% else %>
					if(anyChecked == 0)	returnStr = "<%=e_text%>\n\n";		
					<% end if %>
				<% End If %>
			<% end if %>
			<% if is_other then %>
				<% If values <> "" Then %>
					<% For columnNum = 1 to ubound(columnArray)+1 %>
						checkOther_<%=var_name&"_"&columnNum%>();					
					<% Next %>
				<% Else %>
					<%if customCheckboxOther then%>
						if((formVar.<%=var_name%>_otherbox.checked) && (!checkInput(formVar.<%=var_name%>_other.value))) returnStr += "Please specify other manufacturer.\n\n";
						if((formVar.<%=var_name%>_otherbox_2.checked) && (!checkInput(formVar.<%=var_name%>_other_2.value))) returnStr += "Please specify other Distributor.\n\n";
		            <%else%>                        
						if((formVar.<%=var_name%>_otherbox.checked) && (!checkInput(formVar.<%=v%>.value))){
						    $("#cello<%=var_name&"_"&UBound(ary)+2%>").addClass("validation-error");  
						    returnStr = "Please specify other.\n\n";
						}
					<%end if%>
				<% End If %>
			<% end if %>
			if(returnStr != ""){
			    
			    if("<%=columns_or_rows%>" != "columns"){
			        $("#table<%=var_name%> tr").addClass("validation-error");
			    }
			    else{
			       
			        return returnStr;
			    }
			} else {
			    $("#table<%=var_name%> tr").removeClass("validation-error");
			    return false;
			}
		}
		validateArray[arraycounter] = "validate_<%=var_name%>();";
		arraycounter++;
		
		<% If values = "" Then %>
			<% If is_other Then %>
			function enableOther_<%=var_name%>(item){
				if(item.checked){
					formVar.<%=v%>.disabled = false;
					formVar.<%=v%>.focus(); 
					<% if is_none Then%>clearNone_<%=var_name%>();<% end if %> 
				} else {
					formVar.<%=v%>.value = "";
					formVar.<%=v%>.disabled = true;
				}
			}
				<%if customCheckboxOther then%>
					function enableOther_<%=var_name%>_2(item)
					{
						if(item.checked)
						{
							formVar.<%=var_name%>_other_2.disabled = false;
							formVar.<%=var_name%>_other_2.focus(); 
							<% if is_none Then%>clearNone_<%=var_name%>();<% end if %> 
						} 
						else
						{
							formVar.<%=var_name%>_other_2.value = "";
							formVar.<%=var_name%>_other_2.disabled = true;
						}
					}	
				<%end if%>
			<% End If %>
			<% If is_none Then %>
			function clearNone_<%=var_name%>(){
			    formVar.<%=var_name%>_none.checked = false;
			    <%if customNone then%>formVar.<%=var_name%>_none2.checked = false;<%end if%>
			}	
		

			function clearAll_<%=var_name%>(item){
				if(item.checked){
					<% For z = 1 to i-1 %>
						<%If Session(var_name&"_"&z) <> "false" Then%>
							formVar.<%=var_name%>_<%=z%>.checked = false;
						<%end if%>
						<% If inputlabel <> "" Then %>
								enableInput<%=var_name%>(formVar.<%=var_name%>_<%=z%>,formVar.<%=var_name%>_<%=z%>input);
						<% End If %>
					<% Next %>
					<% if is_other Then %>
					formVar.<%=var_name%>_otherbox.checked = false;
					enableOther_<%=var_name%>(formVar.<%=var_name%>_otherbox);
					<% end if %>
                    <%if customNone then%>formVar.<%=var_name%>_none2.checked = false;<%end if%>
				}
			}
			 <% End If %>

             <%if customNone then%>
                function clearAll2_<%=var_name%>(item){
                    if(item.checked){
                        <% For z = 1 to i-1 %>
                            <%If Session(var_name&"_"&z) <> "false" Then%>
                                formVar.<%=var_name%>_<%=z%>.checked = false;
                        <%end if%>
						<% If inputlabel <> "" Then %>
								enableInput<%=var_name%>(formVar.<%=var_name%>_<%=z%>,formVar.<%=var_name%>_<%=z%>input);
                        <% End If %>
					<% Next %>
					<% if is_other Then %>
					formVar.<%=var_name%>_otherbox.checked = false;
                        enableOther_<%=var_name%>(formVar.<%=var_name%>_otherbox);
                        <% end if %>
                        <%if customNone then%>formVar.<%=var_name%>_none.checked = false;<%end if%>
				}
                }
             <% End If %>
			function checkAll_<%=var_name%>(item){
				if(item.checked){
					<% For z = 1 to i-2 %>
						formVar.<%=var_name%>_<%=z%>.checked = true;
						<% If inputlabel <> "" Then %>
								enableInput<%=var_name%>(formVar.<%=var_name%>_<%=z%>,formVar.<%=var_name%>_<%=z%>input);
						<% End If %>
					<% Next %>
				}
			}
		<% End If %>
		
		
	</script>
<% End Function %>