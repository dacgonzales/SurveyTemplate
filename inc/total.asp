<% Function TOTALCONTROL_HTML( var_name, q_text, e_text, options, is_random, is_none, beginNum, endNum, values)
	privateTOTALCONTROL_HTML var_name, q_text, e_text, options, beginNum, endNum, is_random, is_none, "list", values
End Function

Function TOTALCONTROLCOLUMN_HTML(var_name, q_text, e_text, options, is_random, is_none, beginNum, endNum, values)
	privateTOTALCONTROL_HTML var_name, q_text, e_text, options, beginNum, endNum, is_random, is_none, "column", values
End Function

Function privateTOTALCONTROL_HTML( var_name, q_text, e_text, options, beginNum, endNum, is_random, is_none, displaytype, values ) %>
	<% 	

    dim requiresZero
    requiresZero = false

    Select case var_name
        case "qQJ1"
            requiresZero = true
    End select
        
	If displaytype = "list" Then
		columnArray = split(values,"|")
		rowArray = split(options,"|")
		If is_random Then
			randArray = buildRandArrayReturn(UBound(rowArray)+1)
		End If  
		optionLength = (uBound(rowArray)+1)
		valueLength = (uBound(columnArray)+1)
	Else
		columnArray = split(options,"|")
		rowArray = split(values,"|")
		If is_random Then
			randArray = buildRandArrayReturn(UBound(columnArray)+1)
		End If  
		optionLength = (uBound(columnArray)+1)
		valueLength = (uBound(rowArray) + 1)
	End If
	%>
	<script language="javascript" type="text/javascript">
	    function validate_<%=var_name%>(){
	        var returnStr = "";
	        <% If is_none Then %>
			    if(formVar.<%=var_name%>_none.checked == false) {
			<% End If %>
                <% For i = 1 to valueLength %>
                    if((Number(formVar.<%=var_name%>_<%=i%>_total.value) > <%=endNum%>)||(Number(formVar.<%=var_name%>_<%=i%>_total.value) < <%=beginNum%>)||!(checkInput(formVar.<%=var_name%>_<%=i%>_total.value))) returnStr += "     -";
			    <% Next %>
                emptyCell = 0;
			    if(returnStr == "")
			    {
			        <%if requiresZero then%>
                 
                        for(x = 1; x <= <%=optionLength%>; x++)
                        {
                            $("#rowo<%=var_name%>_" + x).removeClass("validation-error");
			                if(eval("formVar.<%=var_name%>_1_" + x + ".value") == ""){
			                    $("#rowo<%=var_name%>_" + x).addClass("validation-error");
			                    emptyCell++;
			                }
                        
			        }

			        if(emptyCell > 0 && returnStr == "") returnStr = "If you do not have service during a certain day part, please enter '0'.";
			        <%end if%>
			        }
          

			    <% If is_none Then %>
			    }
	        <% End If %>
			if(returnStr != ""){  
			    <%if requiresZero then %>
                        if(emptyCell > 0){
                            return returnStr + "\n\n";
                        }
			else
			{
                $(".totalTableRowTotal").addClass("validation-error");
                return "<%=e_text%>\n" + returnStr + "\n\n";
			}
	        <%else%>    
                $(".totalTableRowTotal").addClass("validation-error");
                return "<%=e_text%>\n" + returnStr + "\n\n";
	        <%end if%>
				
	        } else {
                $(".totalTableRowTotal").removeClass("validation-error");
				return false;
	    }				
	    }
	    function totalValues<%=var_name%>(item, optionValue) {
	        if(isNaN(item.value)){
	            alert("You may only enter a number in this field.");
	            item.value = "";
	            item.focus();
	        } 
	        else  if(!checkWholeNumber(item.value) ){
	            alert("You may only enter a whole number in this field.");
	            item.value = "";
	            item.focus();
	        }
	        else {
	            tmpTotal = 0;
	            for(x = 1; x <= <%=optionLength%>; x++){
	                tmpTotal = tmpTotal + Number(eval("formVar.<%=var_name%>_" + optionValue + "_" + x + ".value"));
	            }
	            if(tmpTotal > <%=endNum%>){
					alert("The total cannot exceed <%=endNum%>.");
	            item.value = "";
	        }
	        eval("formVar.<%=var_name%>_" + optionValue + "_total.value = tmpTotal;");
	    }
	    }
	    function toggleQuestion<%=var_name%>(){
	        if(formVar.<%=var_name%>_none.checked) {
	            //disable all
	            <% For i = 1 to optionLength %>
					<% For j= 1 to valueLength %>
						formVar.<%=var_name%>_<%=i%>_<%=j%>.disabled = true;
	            <% Next %>
            <% Next %>
	            } else {
	            //enable all
	            <% For i = 1 to optionLength %>
					<% For j= 1 to valueLength %>
						formVar.<%=var_name%>_<%=i%>_<%=j%>.disabled = false;
	            <% Next %>
            <% Next %>
	            }
	    }
	    validateArray[arraycounter] = "validate_<%=var_name%>();";
	    arraycounter++;
	</script>
	
	<% DIV_HTML var_name, q_text, "question" %>
	
	<div id="diva<%=var_name%>" class="answer">
	<table width="100%"  border="0" cellspacing="0" cellpadding="5" class="totalTable" id="table<%=var_name%>">
		<tr class="totalTableRowHead" id="rowh<%=var_name%>">
			<td class="totalTableCellHeadBlank" id="cellhb<%=var_name%>">&nbsp;</td>
			<% 
			For i = 1 to Ubound(columnArray) + 1
				ColumnText = columnArray(i-1)				
				if displaytype = "list" then
					if i mod 2 = 0 then
						tdClass = "totalTableCellHead1"
					else
						tdClass = "totalTableCellHead2"
					end if
				else
					If is_random Then
						ColumnText = columnArray(randArray(i)-1)
					End If
					tdClass = "totalTableCellHead"
				end If
				%>
				<td class="<%=tdClass%>" id="cellh<%=var_name%>_<%=i%>"><%=ColumnText%></td>
			<%
			Next
			If displaytype = "column" Then
				%>
				<td class="totalTableCellHeadTotal" id="cellht<%=var_name%>">Total:</td>
				<%
			End If
			%>
		</tr>
		<%
		For i = 1 to Ubound(rowArray) + 1
			RowText = rowArray(i-1)
			RowValue = i
			if displaytype = "list" then
				If is_random Then
					RowText = rowArray(randArray(i)-1)
					RowValue = randArray(i)
				End If
				rowClass = "totalTableRow"
			else
				if i mod 2 = 0 then
					rowClass = "totalTableRow1"
				else
					rowClass = "totalTableRow2"
				end if
			end if
			%>
			<tr class="<%=rowClass%> totalTableRow" id="rowo<%=var_name%>_<%=i%>">
				<td class="totalTableCellItem" id="cello<%=var_name%>_<%=i%>" style="width:80%"><%=RowText%></td>
				<%
				For j = 1 to Ubound(columnArray) + 1
					ColumnValue = j
					If displaytype = "list" Then
						OptionValue = ColumnValue
						ValueValue = RowValue
						if j mod 2 = 0 then
							tdClass = "totalTableCell1"
						else
							tdClass = "totalTableCell2"
						end If
					Else
						If is_random Then
							ColumnValue = randArray(j)
						End If
						OptionValue = RowValue
						ValueValue = ColumnValue
						tdClass = "totalTableCell"
					End If
					%>
					<td class="<%= tdClass %>" id="celli<%=var_name%>_<%=OptionValue%>_<%=ValueValue%>"><input type="text" id="<%=var_name%>_<%=OptionValue%>_<%=ValueValue%>" size="5" name="<%=var_name%>_<%=OptionValue%>_<%=ValueValue%>" value="" class="totalInput" onChange="totalValues<%=var_name%>(this,<%=OptionValue%>);"> %</td>
					<% questionIDs = questionIDs & var_name & "_" & OptionValue & "_" & ValueValue & "|" %>
					<%
				Next
				If displaytype = "column" Then
					%>
					<td class="totalTableCellTotal" id="cellt<%=var_name%>_<%=i%>"><input type="text" id="<%=var_name%>_<%=OptionValue%>_total" name = "<%=var_name%>_<%=OptionValue%>_total" value="" class="totalTotalInput" disabled>(<%=OptionValue%>)</td>
					<%
				End If
				%>
			</tr>
			<%
		Next
		If displaytype = "list" Then
		%>
		<tr class="totalTableRowTotal" id="rowt<%=var_name%>">
			<td class="totalTableCellTotalItem" id="cellt<%=var_name%>">Total:</td>
			<%
			For i = 1 to Ubound(columnArray) + 1
				ColumnText = columnArray(i-1)
				ColumnValue = i
				If i mod 2 = 0 Then
					tdClass = "totalTableCellTotal1"
				Else
					tdClass = "totalTableCellTotal2"
				End If
				%>
				<td class="<%= tdClass %>" id="cellti<%=var_name%>_<%=i%>"><input type="text" id="<%=var_name%>_<%=ColumnValue%>_total" size="5" name="" value="" class="totalTotalInput" disabled> %</td>
			<%
			Next
			%>
		</tr>		
		<%
		End If
		If is_none Then
		%>
		<tr class="totalTableRowNone" id="rown<%=var_name%>">
			<td colspan="<%=ubound(columnArray)+2%>" class="totalTableCellNone" id="celln<%=var_name%>"><input type="checkbox" id="<%=var_name%>_none" name="<%=var_name%>_none" value="1" onClick="toggleQuestion<%=var_name%>();" class="totalNoneInput"><label for="<%=var_name%>_none">Don't Know / Unsure</label></td>
		</tr>
		<%
		End If
		%>
	</table>	
	</div>
	
<% End Function %>