<% 
Function RANK_HTML( var_name, q_text, e_text, options, is_random )
	private_RANK_HTML var_name, q_text, e_text, options, is_random, false, ""
End Function

Function RANK_SELECT_HTML( var_name, q_text, e_text, options, is_random )
	private_RANK_HTML var_name, q_text, e_text, options, is_random, true, ""
End Function

Function RANK_SELECT_LIMITED_HTML( var_name, q_text, e_text, options, is_random, endNum)
	private_RANK_HTML var_name, q_text, e_text, options, is_random, true, endNum
End Function

%>
<% Function private_RANK_HTML( var_name, q_text, e_text, options, is_random, is_select, endNum) %>
	<%
		'MARC April 21, 2010 set max number for endNum
		if var_name = "VARNAME" then
			endNum = Session("MAX_RANK_NUMBER")
		end if	
	%>
	<script language="javascript" type="text/javascript">	
	function validate_<%=var_name%>(){
		returnStr = "";
		<% 
		ary = Split(options,"|")
		If is_random Then
		   randArray = buildRandArrayReturn(UBound(ary)+1)
		End If  
		If endNum <> "" Then
			%> valuesCheck = 0; <%
		End If
		For i = 1 To UBound(ary)+1 
			text = ary(i-1)
			val = i
			If is_random Then
				val = randArray(i)
				text = ary(randArray(i)-1)
			End If 
			%>
			<% If e_text <> "" Then %>
				<% If endNum = "" Then %>
					if(!(checkInput(eval("formVar.<%=var_name%>_<%=val%>.value")))) returnStr += "    <%=text%>\n";
				<% Else %>
					if((checkInput(eval("formVar.<%=var_name%>_<%=val%>.value")))) valuesCheck++;
				<% End If %>
			<% End If %>
		<% Next %>	
		<% If endNum <> "" Then %>
			if(valuesCheck < <%=endNum%>) returnStr += ("\n");
		<% End If %>
		if(returnStr != ""){
			return "<%=e_text%>\n" + returnStr + "\n";
		} else {
			return false;
		}
	}
	validateArray[arraycounter] = "validate_<%=var_name%>();";
	arraycounter++;
	
	//'MARC April 21, 2010 
	<%if var_name = "VARNAME" then%>
		function checkRank2(item, controlLen, qNum){
			// controlLen = parseInt(eval("document.theform."+controlLen+".value"));
			if((isNaN(item.value))||(parseInt(item.value) > controlLen)||(parseInt(item.value) <= 0)){
				alert('you can only use numbers\nbetween 1 and '+controlLen+'\nto rank these items');
				item.value = "";
				item.focus();
			}
			for(var i = 1; i <= controlLen; i++){
				var obj = document.getElementById(qNum+i);
				if(obj != null)
				{
					if(item.name != qNum+i){
						if(item.value == eval("document.theform."+qNum+i+".value")){
							eval("document.theform."+qNum+i+".value = ''");
						}			
					}
				}
			}
		}		
	<%end if%>
	</script>	
	<% DIV_HTML "q"&var_name, q_text, "question" %>	
	<div id="diva<%=var_name%>" class="answer">
		<table width="100%"  border="0" cellspacing="0" cellpadding="5" class="rankTable" id="table<%=var_name%>">
		<%
		'MARC April 21, 2010
		Dim customColor
		customColor = 0
		
		For i = 1 To UBound(ary)+1
			text = ary(i-1)
			val = i 
			If is_random Then
				text = ary(randArray(i)-1)
				val = randArray(i)
			End If
				'MARC April 21, 2010 row color if RANK items are hidden
				if customColor mod 2 = 0 then
					rowClass = "rankTableRow1"
				else
					rowClass = "rankTableRow2"
				end if
			%>
			<%'MARC April 21, 2010 added hiding of RANK
			if Session(var_name&"_"&i) <> "false" then
				customColor = customColor + 1
			%>
				<tr class="<%=rowClass%>" id="row<%=var_name&"_"&val%>">
					<td class="rankTableCellOption" id="cello<%=var_name&"_"&val%>"><%=text%></td>
					<td class="rankTableCellInput" id="celli<%=var_name&"_"&val%>"><%
					if is_select = true then %><select name="<%=var_name%>_<%= val %>" id="<%=var_name%>_<%= val %>" onChange="checkRank2(this,<%=UBound(ary)+1%>,'<%=var_name%>_');">
						<option value="">-</option>
						<% If endNum = "" Then %>
							<% For x = 1 to (Ubound(ary) + 1) %>
							<option value="<%=x%>">#<%=x%></option>
							<% Next %></select>
						<% Else %>
							<% For x = 1 to endNum %>
							<option value="<%=x%>">#<%=x%></option>
							<% Next %></select>
						<% End If %>
					<% else %> <input type="text" name="<%=var_name%>_<%= val %>" id="<%=var_name%>_<%= val %>" size="5" maxlength="2" onChange="checkRank2(this,<%=UBound(ary)+1%>,'<%=var_name%>_');" class="rankInput">
				<% end if %></td>
				</tr>
				<% questionIDs = questionIDs & var_name&"_"&val &"|" %>
			<%end if%>
		<% Next %>
		</table>
	</div>
<% End Function %>