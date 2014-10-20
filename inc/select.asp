<% Function SELECT_HTML( var_name, q_text, e_text, options ) %>
	<% privateBUILD_SELECT var_name, q_text, e_text, options, "", ""  %>
<% End Function %>

<% Function SELECT_VALUES_HTML( var_name, q_text, e_text, options) %>
	<% privateBUILD_SELECT var_name, q_text, e_text, options, "", "VALUES"  %>
<% End Function %>

<% Function SELECT_REPEAT_HTML( var_name, q_text, e_text, options, labelString) %>
	<% privateBUILD_SELECT var_name, q_text, e_text, options, labelString, "REPEAT" %>
<% End Function %>

<% Function privateBUILD_SELECT(var_name, q_text, e_text, options, labelString, valuetype) %>
	<script language="javascript" type="text/javascript">
	function validate_<%=var_name%>(){
		<% If e_text <> "" Then %>
			<% If valuetype <> "REPEAT" Then %>
				if ( !checkInput(formVar.<%=var_name%>.value) )
						return "<%=e_text%>\n\n";
				return false
			<% Else %>
				returnstr = "";
				<% 	optionArray = split(options, "|") %>
				<% For i = 1 to ubound(optionArray)+1 %>
					 if ( !checkInput(formVar.<%=var_name&"_"&i%>.value) ) returnstr = returnstr + "     <%=optionArray(i-1)%>\n";
				<% Next %>
				if(returnstr != ""){
					returnstr = "<%=e_text%>\n" + returnstr + "\n\n";
					return returnstr;
				} else {
					return false;
				}
			<% End If %>
		<% End If %>
	}
	validateArray[arraycounter] = "validate_<%=var_name%>();";
	arraycounter++;
	</script>	
	<% DIV_HTML "q"&var_name, q_text, "question" %>	
	<div id="diva<%=var_name%>" class="answer">
	<% If valuetype = "REPEAT" Then  	
		optionArray = split(options, "|")
		valuesArray = split(labelString, "|") %>
		<table cellpadding="2" cellspacing="0" width="100%" border="0">
			<% 	
				For i = 1 to Ubound(optionArray)+1 
				if i  mod 2 = 0 then
					rowClass = "selectTableRow1"
				else
					rowClass = "selectTableRow2"
				end if
				%>				
			<tr class="<%=rowClass%>">
				<td><%=optionArray(i-1)%></td>
				<td><select id="<%=var_name&"_"&i%>" name="<%=var_name&"_"&i%>" class="selectInput">
					<option value="" class="selectOption">-</option>
					<%	For j = 1 to Ubound(valuesArray)+1 %>
					<option value="<%=j%>" class="selectOption"><%=valuesArray(j-1)%></option>
					<%	Next %>
				</select>
				<% questionIDs = questionIDs & var_name&"_"&i &"|" %>
				</td>
			</tr>
			<% 	Next %>
		</table>
	<% Else %>
		<select id="<%=var_name%>" name="<%=var_name%>" class="selectInput">
			<option value="" class="selectOption">--select--</option>
		<% ary = Split(options,"|")
		   For i = 0 To UBound(ary)
			If valuetype = "VALUES" Then
				valary = Split(ary(i),"^")
				If ubound(valary) > 0 Then
					val = valary(0)
					text = valary(1)
				Else
					val = valary(0)
					text = valary(0)
				End If
			Else
				val = i + 1
				text = ary(i)
			End If
			%>
			<option value="<%=val%>" id="<%=var_name%>|<%=i%>" name="<%=var_name%>|<%=i%>" class="selectOption"><%=text%></option>
		<% Next %>
		</select>
		<% questionIDs = questionIDs & var_name &"|" %>
	<% End If %>
	</div>
<% End Function %>

