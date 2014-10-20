<% 
Function RADIOSLIDESCALE_HTML(var_name, q_text, e_text, options, is_random, beginNum, endNum, bottomLabel, topLabel, offsetLabelString) 
	privateRADIOSCALE_HTML var_name, q_text, e_text, options, is_random, "slidescale", beginNum, endNum, bottomLabel, topLabel, "", "", offsetLabelString
End Function 
	
Function RADIOSLIDESCALESELECT_HTML(var_name, q_text, e_text, options, is_random, beginNum, endNum, bottomLabel, topLabel, initSelect,  offsetLabelString)
	privateRADIOSCALE_HTML var_name, q_text, e_text, options, is_random, "slidescaleselect", beginNum, endNum, bottomLabel, topLabel, initSelect, "", offsetLabelString
End Function 

Function RADIOSLIDESCALELABEL_HTML( var_name, q_text, e_text, options, is_random, labelString, offsetLabelString)
	privateRADIOSCALE_HTML var_name, q_text, e_text, options, is_random, "slidescalelabel", "", "", "", "", "", labelString,  offsetLabelString
End Function 

%>

<% Function privateRADIOSCALE_HTML( var_name, q_text, e_text, options, is_random, scaletype, beginNum, endNum, bottomLabel, topLabel, initSelect, labelString, offsetLabelString  ) %>
	<% 	
	'MARC April 22, 2010 scale with checkbox to select
	Dim scalewithBox
	Dim customSlide 'scale with OTHER
	scalewithBox = false
	customSlide = false
	select case var_name
		case "VARNAME"
			scalewithBox = true
		case "qQB6"
			customSlide = true
            scalewithBox = true
	end select
	
	if isNull(options) or options = "" then
		dim ary(0)
		ary(0) = ""
	else
		ary = Split(options,"|")
	end if
	If is_random Then
	   randArray = buildRandArrayReturn(UBound(ary)+1)
	End If  
	%>
	<script language="javascript" type="text/javascript">
	function validate_<%=var_name%>(){
		<%if scalewithBox then%>
			var unCheck_counts = 0;
		<%end if%>
		returnStr = "";
		<% 	
			displayCount = 0
			For i = 1 To UBound(ary)+1
			text = ary(i-1)
			val = i
			If is_random Then
				val = randArray(i)
				text = ary(randArray(i)-1)
			End If  
			%>
			<% If Session(var_name&"_"&val) <> "false" Then 
				displayCount = displayCount + 1 %>
				<% If scaletype <> "slidescaleselect" Then %>
					<%if scalewithBox then%>
						//if ( !checkRadio(formVar.<%=var_name&"_"&val%>) )	unCheck_counts++;
					<%else%>
						if ( !checkRadio(formVar.<%=var_name&"_"&val%>) )	{						    
						    $("#row<%=var_name&"_"&val%>").addClass("validation-error");
						    returnStr += "  <%=displayCount%>";
						}
                        else{						    
						    $("#row<%=var_name&"_"&val%>").removeClass("validation-error");
                }
						<%end if%>
                    <% Else %>
                        if (getRadio(formVar.<%=var_name&"_"&val%>) == <%=initSelect%>) {						    
						    $("#row<%=var_name&"_"&val%>").addClass("validation-error");
						    returnStr += "  <%=displayCount%>";
						}
						else{						    
						    $("#row<%=var_name&"_"&val%>").removeClass("validation-error");
						}
				<% End If %>
			<% End If %>

            <%if scalewithBox then%>
                $("#row<%=var_name&"_"&val%>").removeClass("validation-error");
                if(formVar.<%=var_name&"_"&val%>_box.checked){
                    if ( !checkRadio(formVar.<%=var_name&"_"&val%>) ){
                        $("#row<%=var_name&"_"&val%>").addClass("validation-error");
                        returnStr += "<br/>Please rate <%=text%>";
                    }
		    	}
	            else{
                    unCheck_counts++;
	            }

				if(unCheck_counts==<%= UBound(ary)+1 %>){
                    returnStr += "<br/>Please answer at least one";
				}
		    <%end if%>
		<% Next %>
		

		<%if customSlide then%>
            $("#rowo<%=var_name&"_"&UBound(ary)+2%>").removeClass("validation-error");			    
		    $("#row<%=var_name&"_other"%>").removeClass("validation-error");

			if (formVar.<%=var_name%>_otherbox.checked && !checkInput(formVar.<%=var_name%>_othertext.value)){
			    $("#rowo<%=var_name&"_"&UBound(ary)+2%>").addClass("validation-error");			    
			    returnStr += "  <br/>Please specify other";
			}
			if ( !checkRadio(formVar.<%=var_name&"_other"%>) && (formVar.<%=var_name%>_otherbox.checked && checkInput(formVar.<%=var_name%>_othertext.value)) ){
			    $("#row<%=var_name&"_other"%>").addClass("validation-error");
			    returnStr += "  <br/>Please specify an answer for other";
			}
		<%end if%>

		var qLength;
		qLength = <%= UBound(ary)+1 %>
		if(returnStr != ""){
			if(qLength == 1){
				return "<%=e_text%><br/><br/>";
			} else {
				return "<%=e_text%><br/><br/>" + returnStr;
			}
			//$("#rowo<%=var_name&"_"&val%>").addClass("validation-error");
		} else {
			return false;
		}
	}
	validateArray[arraycounter] = "validate_<%=var_name%>();";
	arraycounter++;
	
		<%if scalewithBox then%>
			function <%=var_name%>_enableCheckRad(obj,val)
			{
				if(obj.checked)
				{
					for(var x = <%=beginNum%>;x<=<%=endNum%>;x++)
					{
						var o_rad = document.getElementById('<%=var_name%>_'+val+'|'+x);
						o_rad.disabled = false;
					}
				}
				else
				{
				    for(var x = <%=beginNum%>;x<=<%=endNum%>;x++)
					{
						var o_rad = document.getElementById('<%=var_name%>_'+val+'|'+x);
						o_rad.disabled = true;
						o_rad.checked = false;
					}
				}
			}
		<%end if%>
		
		<%if customSlide then%>
			function <%=var_name%>_enableRadios(obj)
			{
				if(obj.checked)
				{
					formVar.<%=var_name%>_othertext.disabled = false;
					formVar.<%=var_name%>_othertext.focus();
					for(var x = 1;x<=10;x++)
					{
						var o_rad = document.getElementById('<%=var_name%>_other|'+x);
						o_rad.disabled = false;
					}
				}
				else
				{
					formVar.<%=var_name%>_othertext.disabled = true;
					formVar.<%=var_name%>_othertext.value = '';
					for(var x = 1;x<=10;x++)
					{
						var o_rad = document.getElementById('<%=var_name%>_other|'+x);
						o_rad.disabled = true;
						o_rad.checked = false;
					}
				}
			}
		<%end if%>
	</script>
	
		<% DIV_HTML "q"&var_name, q_text, "question" %>
	
	<div id="diva<%=var_name%>" class="answer table-responsive">
		<table width="100%"  border="0" cellspacing="0" cellpadding="5" class="radioScaleTable table" id="table<%=var_name%>">
		<% 	displayCounter = 0
			For i = 1 To UBound(ary)+1
			text = ary(i-1)
			val = i
			If is_random Then
				val = randArray(i)
				text = ary(randArray(i)-1)
			End If 
			if displayCounter mod 2 = 0 then
				rowClass = "radioScaleTableRow1"
			else
				rowClass = "radioScaleTableRow2"
			end if
				%> 	    
			<% If Session(var_name&"_"&val) <> "false" Then
				displayCounter = displayCounter + 1 %>
				<% if  (ary(0) <> "") then %>
			<tr class="<%= rowClass %>" id="rowo<%=var_name&"_"&val%>">
				<td id="cello<%=var_name&"_"&val%>" class="radioScaleOption" <% If offsetLabelString <> "" Then %>colspan="2"<%End If%>><span class="radioScaleNumber"><%=displayCounter%>.</span>
				<%if scalewithBox then%>
					<input type="checkbox" id="<%=var_name%>_<%=val%>_box" name="<%=var_name%>_<%=val%>_box" value="1" onclick="<%=var_name%>_enableCheckRad(this,<%=val%>);">
				<%end if%>
				<%=text%>
				</td>
			</tr>
				<% end if %>
			<tr class="<%= rowClass %>" id="rows<%=var_name&"_"&val%>">
				<td align="center" id="cells<%=var_name&"_"&val%>" class="radioScaleScale">
				<%if scalewithBox then%>
					<%privateBuildSlideScale_2 beginNum, endNum, bottomLabel, topLabel, var_name&"_"&val%>
				<%else%>
					<% Select Case scaletype
							Case "slidescale"  privateBuildSlideScale beginNum, endNum, bottomLabel, topLabel, var_name&"_"&val 
							Case "slidescaleselect"  privateBuildSlideScaleSelect beginNum, endNum, bottomLabel, topLabel, var_name&"_"&val , initSelect 
							Case "slidescalelabel"  privateBuildSlideScaleLabel labelString, var_name &"_"&val  
						End Select%>
				<%end if%>
				</td>
				<%	If offsetLabelString <> "" Then %>
				<td align="center" id="cells<%=var_name&"_"&val%>" class="radioScaleScaleOffset">
					<% 
					If endNum <> "" Then
						val_offset = endNum
					Else
						tmpArray = split(labelString, "|")
						val_offset = ubound(tmpArray)+1
					End If
					privateBuildSlideScaleLabelOffset offsetLabelString, var_name&"_"&val , val_offset  %>
				</td>
				<% End If %>
			</tr>
				<% questionIDs = questionIDs & var_name&"_"&val &"|" %>
			<% End If %>
		<% Next %>
			<%if customSlide then
				if displayCounter mod 2 = 0 then
					rowClass = "radioScaleTableRow1"
				else
					rowClass = "radioScaleTableRow2"
				end if
				displayCounter = displayCounter + 1
			%>
				<tr class="<%= rowClass %>" id="rowo<%=var_name&"_"&displayCounter%>">
					<td id="cello<%=var_name&"_"&displayCounter%>" class="radioScaleOption"><span class="radioScaleNumber"><%=displayCounter%>.</span> 
						Other
						<input type="checkbox" id="<%=var_name%>_otherbox" name="<%=var_name%>_otherbox" value="1" onclick="<%=var_name%>_enableRadios(this);">
						<input type="text" id="<%=var_name%>_othertext" name="<%=var_name%>_othertext" value="" disabled>
						<% questionIDs = questionIDs & var_name&"_othertext" &"|" %>
					</td>
				</tr>
				<tr class="<%= rowClass %>" id="rows<%=var_name&"_"&displayCounter%>">
					<td align="center" id="cells<%=var_name&"_"&displayCounter%>" class="radioScaleScale">
						<% Select Case scaletype
							Case "slidescale"  privateBuildSlideScale_2 beginNum, endNum, bottomLabel, topLabel, var_name&"_other" 
						End Select%>
					</td>
				</tr>
					<% questionIDs = questionIDs & var_name&"_other" &"|" %>	
			<%end if%>
		</table>
	</div>
<% End Function %>

<%
Function RADIOSLIDESCALE_NONUM_GRID_HTML( var_name, q_text, e_text, options, is_random, labelString, offsetLabelString)  
	%>
	<% 	
	'MARC April 22, 2010 GRID with other
	Dim custom_other,isColumn
	custom_other = false
    custom_otherLabel = "Other"
    custom_other2 = false
    custom_otherLabel2 = "Other"
    isColumn = false
	select case var_name
		case "qQB2"
			custom_other = true
            custom_otherLabel = "Other MANUFACTURER brand (please specify): "
            custom_other2 = true
            custom_otherLabel2 = "Other DISTRIBUTOR brand (please specify): "
        case "qQB5"
			isColumn = true
	end select
	
	if isNull(options) or options = "" then
		dim ary(0)
		ary(0) = ""
	else
		options2split = Replace(options, vbNewLine, "")
		ary = Split(options2split,"|")
	end if
	If is_random Then
	   randArray = buildRandArrayReturn(UBound(ary)+1)
	End If  
	%>
	<script language="javascript" type="text/javascript">
	function validate_<%=var_name%>(){
		returnStr = "";
		<% 	For i = 1 To UBound(ary)+1
			text = ary(i-1)
			val = i
			If is_random Then
				val = randArray(i)
				text = ary(randArray(i)-1)
			End If  
			%>
			 <% If Session(var_name&"_"&val) <> "false" Then %>
                <%if NOT isColumn then%>
                    $("#rowo<%=var_name&"_"&val%>").removeClass("validation-error");
                    if ( !checkRadio(formVar.<%=var_name&"_"&val%>) ){
                        $("#rowo<%=var_name&"_"&val%>").addClass("validation-error");
                        returnStr += "    <%=text%>\n";	           
                    }
                <% End If %>
            <% End If %>
		<% Next %>

         
         <%if isColumn then%>			
             <%labelStringArray = split(labelString,"|")
         %>
         
            <%For ii = 1 To UBound(labelStringArray)+1 %>
                    if (!checkRadio(formVar.<%=var_name&"_"&ii%>) )	
                    {         
                        <% 	For j = 1 To UBound(ary)+1
                        val2 = i
                        If is_random Then
                        val2 = randArray(j)                            
                        End If  
                    %>
                            $("#cell<%=var_name &"_"&val2&"-"&ii-1%>").addClass("validation-error");
                        <%next%>
                        returnStr += "    <%=labelStringArray(ii-1)%>\n";

                    }
	            else{
                    <% 	For j = 1 To UBound(ary)+1
                    val2 = i
                    If is_random Then
                    val2 = randArray(j)                            
                    End If  
                %>
                    $("#cell<%=var_name &"_"&val2&"-"&ii-1%>").removeClass("validation-error");
                    <%Next%>

                    }
	               
                        
                   
	                
         <%next%>
        <% End If %>

		<%if custom_other then%>
            $("#rowo<%=var_name&"_"& UBound(ary)+2%>").removeClass("validation-error");
			if ( formVar.<%=var_name%>_otherbox.checked && !checkInput(formVar.<%=var_name%>_othertext.value) ){
			    $("#rowo<%=var_name&"_"& UBound(ary)+2%>").addClass("validation-error");
			    returnStr += "    Please specify OTHER\n";
			}
			if ( !checkRadio(formVar.<%=var_name&"_other"%>) && ( formVar.<%=var_name%>_otherbox.checked && checkInput(formVar.<%=var_name%>_othertext.value) ) ){
			    $("#rowo<%=var_name&"_"& UBound(ary)+2%>").addClass("validation-error");
			    returnStr += "    Please select an answer for OTHER\n";
			}
			<%end if%>
        <%if custom_other2 then%>
            $("#rowo<%=var_name&"_"& UBound(ary)+3%>").removeClass("validation-error");
			if ( formVar.<%=var_name%>_otherbox2.checked && !checkInput(formVar.<%=var_name%>_othertext2.value) ){
			    $("#rowo<%=var_name&"_"& UBound(ary)+3%>").addClass("validation-error");
			    returnStr += "    Please specify OTHER\n";
			}
			if ( !checkRadio(formVar.<%=var_name&"_other2"%>) && ( formVar.<%=var_name%>_otherbox2.checked && checkInput(formVar.<%=var_name%>_othertext2.value) ) ){
			    $("#rowo<%=var_name&"_"& UBound(ary)+3%>").addClass("validation-error");
			    returnStr += "    Please select an answer for OTHER Distributor\n";
			}
            <%end if%>
		var qLength;
		qLength = <%= UBound(ary)+1 %>;
		hasCurrently = false;
		if(returnStr != ""){
			if(qLength == 1){
				return "<%=e_text%>\n\n";
			} else {
				return "<%=e_text%>\n" + returnStr + "\n\n";
			}
		} else {
		    <%if var_name="qQB2" then
                For i = 1 To UBound(ary)+1		        
                val = i
                If is_random Then
                val = randArray(i)		        
                End If  
                %>
                
            if (getRadio(formVar.<%=var_name&"_"&val%>) == 1)
                {
                    hasCurrently  = true;
                }
		        
                <% Next%>

                <%if custom_other then%>
                    if (formVar.<%=var_name%>_otherbox.checked && checkInput(formVar.<%=var_name%>_othertext.value) && getRadio(formVar.<%=var_name&"_other"%>) == 1) hasCurrently  = true;
                <%end if%>
                <%if custom_other2 then%>
                    if (formVar.<%=var_name%>_otherbox2.checked && checkInput(formVar.<%=var_name%>_othertext2.value) && getRadio(formVar.<%=var_name&"_other2"%>) == 1) hasCurrently  = true;
                <%end if%>

                if(!hasCurrently)
                {
                    return "Must select at least one brand you currently use.";
                }
                else
                {
                    return false;
                }
                <%Else%>
                    return false;
                <%end if%>

            

		  
		}
    }
	validateArray[arraycounter] = "validate_<%=var_name%>();";
	arraycounter++;
		<%if custom_other then %>
			function <%=var_name%>_CenableOthers(obj)
			{
				if(obj.checked)
				{
					formVar.<%=var_name%>_othertext.disabled=false;
					formVar.<%=var_name%>_othertext.focus();
					for(var x = 0;x<=<%=UBOUND(split(labelString, "|"))%>;x++)
					{
						var o_rad = document.getElementById('<%=var_name%>_other|'+x);
						o_rad.disabled = false;
					}
				}
				else
				{
					formVar.<%=var_name%>_othertext.disabled=true;
					formVar.<%=var_name%>_othertext.value='';
					for(var x = 0;x<=<%=UBOUND(split(labelString, "|"))%>;x++)
					{
						var o_rad = document.getElementById('<%=var_name%>_other|'+x);
						o_rad.disabled = true;
						o_rad.checked = false;
					}					
				}
				
			}
			<%end if%>

                 <%if custom_other2 then %>
			function <%=var_name%>_CenableOthers2(obj)
			    {
			        if(obj.checked)
			        {
			            formVar.<%=var_name%>_othertext2.disabled=false;
			            formVar.<%=var_name%>_othertext2.focus();
			            for(var x = 0;x<=<%=UBOUND(split(labelString, "|"))%>;x++)
			            {
			                var o_rad = document.getElementById('<%=var_name%>_other2|'+x);
			                o_rad.disabled = false;
			            }
			        }
			        else
			        {
			            formVar.<%=var_name%>_othertext2.disabled=true;
			            formVar.<%=var_name%>_othertext2.value='';
			            for(var x = 0;x<=<%=UBOUND(split(labelString, "|"))%>;x++)
			            {
			                var o_rad = document.getElementById('<%=var_name%>_other2|'+x);
			                o_rad.disabled = true;
			                o_rad.checked = false;
			            }					
			        }
				
			    }
			    <%end if%>

                  <%if isColumn then%>
                    function eventhere(colIdx, rowIdx)
                    {
                        for(col=1; col <= 4; col++)
                        {
                            if(col == colIdx)
                            {
                                continue;
                            }
                            else
                            {				
                                var x = document.getElementById('<%=var_name%>' + '_' + col + '|' +rowIdx);
                                //x.checked=false;
                            }
                        }
                    }
                 <%end if%>
	</script>
	
		<% DIV_HTML "q"&var_name, q_text, "question" %>
	
	<div id="diva<%=var_name%>" class="answer table-responsive">
		<table width="100%"  border="0" cellspacing="0" cellpadding="2" class="radioScaleTable table" id="table<%=var_name%>">
			<tr class="radioScaleRowRankOption">
				<td width="40%">&nbsp;</td>
				<% rankHeaders = split(labelString, "|")
					numberOfRankOptions = ""
					For i = 0 to ubound(rankHeaders)
						Response.Write "<td align=""center"" valign=""middle"" width=""" & cInt(60/(ubound(rankHeaders)+1)) & "%"" class=""radioScaleRankOption"">" & rankHeaders(i) & "</td>" & vbNewLine
						numberOfRankOptions = numberOfRankOptions & "|"
					Next
					numberOfRankOptions = Left(numberOfRankOptions,Len(numberOfRankOptions)-1)
				%>
			</tr>
		<% 	displayCounter = 0
			For i = 1 To UBound(ary)+1
			If Session(var_name&"_"&val) <> "false" Then
				displayCounter = displayCounter + 1
			End If
			text = ary(i-1)
			val = i
			If is_random Then
				val = randArray(i)
				text = ary(randArray(i)-1)
			End If 
			if displayCounter mod 2 = 0 then
				rowClass = "radioScaleTableRow1"
			else
				rowClass = "radioScaleTableRow2"
			end if
				%> 	    
			<% If Session(var_name&"_"&val) <> "false" Then %>
			<tr class="<%= rowClass %>" id="rowo<%=var_name&"_"&val%>">
				<td id="cello<%=var_name&"_"&val%>" class="radioScaleOption" <% If offsetLabelString <> "" Then %>colspan="2"<%End If%>><%=text%></td>
				<%
				labelStringAry = Split(numberOfRankOptions,"|")		
				%>

                <%if NOT isColumn then%>
				    <% For labelStringCounter = 0 to ubound(labelStringAry) %>
					    <td align="center" valign="top" class="radioScaleOption" id="cell<%=var_name &"_"&val%>|<%=labelStringCounter%>"><input type="radio" name="<%=var_name &"_"&val%>" id="<%=var_name &"_"&val%>|<%=labelStringCounter%>" value="<%=(labelStringCounter + 1)%>" class="slideScaleLabelInput"><br><label for="<%=var_name &"_"&val%>|<%=labelStringCounter%>"><%=labelStringAry(labelStringCounter)%></label></td>
				    <% Next %>
                <%else %>
                    <% For labelStringCounter = 0 to ubound(labelStringAry) %>
							<td align="center" valign="top" class="radioScaleOption" id="cell<%=var_name &"_"&val%>-<%=labelStringCounter%>"><input type="radio" name="<%=var_name &"_"&labelStringCounter + 1%>" id="<%=var_name &"_"&labelStringCounter + 1%>|<%=val%>" value="<%=val%>" onclick="eventhere('<%=labelStringCounter + 1%>', '<%=val%>');" class="slideScaleLabelInput"><br><label for="<%=var_name &"_"&labelStringCounter + 1%>|<%=i%>"><%=labelStringAry(labelStringCounter)%></label></td>
						<% Next %>		
                <%end if %>
			</tr>
                <%if NOT isColumn then%>
				    <% questionIDs = questionIDs & var_name&"_"&val &"|" %>
                <% End If %>
			<% End If %>
		<% Next %>

        	<%if isColumn then%>
			<% For labelStringCounter = 0 to ubound(labelStringAry) %>
				<% questionIDs = questionIDs & var_name&"_"&labelStringCounter+1 &"|" %>
			<% Next %>	
		<% End If %>
			<%
				'if (custom_other) then
				'	displayCounter = displayCounter + 1
				'	if displayCounter mod 2 = 0 then
				'		rowClass = "radioScaleTableRow1"
				'	else
				'		rowClass = "radioScaleTableRow2"
				'	end if				
				'end if
			%>
			<%if custom_other then
                displayCounter = displayCounter + 1
                %>
                
				<tr class="<%= rowClass %>" id="rowo<%=var_name&"_"&displayCounter%>">
					<td id="cello<%=var_name&"_"&displayCounter%>" class="radioScaleOption">
					<%=custom_otherLabel %><br />
					<input type="checkbox" id="<%=var_name%>_otherbox" name="<%=var_name%>_otherbox" value="1" onclick="<%=var_name%>_CenableOthers(this);">
					<input type="text" id="<%=var_name%>_othertext" name="<%=var_name%>_othertext" value="" disabled>
					<% questionIDs = questionIDs & var_name&"_othertext" &"|" %>
					</td>
						<%for x = 0 to UBOUND(rankHeaders) %>
						<td align="center" valign="top" class="radioScaleOption" id="cell<%=var_name &"_other"%>|<%=x%>"><input type="radio" name="<%=var_name & "_other"%>" id="<%=var_name & "_other"%>|<%=x%>" value="<%=x+1%>" class="slideScaleLabelInput" disabled><br></td>
						
					<%next%>
					<% questionIDs = questionIDs & var_name&"_other" &"|" %>
				</tr>
			<%end if%>		

             <%if custom_other2 then
                 displayCounter = displayCounter + 1
                 %>
				<tr class="<%= rowClass %>" id="rowo<%=var_name&"_"&displayCounter%>">
					<td id="cello<%=var_name&"_"&displayCounter%>" class="radioScaleOption">
					<%=custom_otherLabel2 %><br />
					<input type="checkbox" id="<%=var_name%>_otherbox2" name="<%=var_name%>_otherbox2" value="1" onclick="<%=var_name%>_CenableOthers2(this);">
					<input type="text" id="<%=var_name%>_othertext2" name="<%=var_name%>_othertext2" value="" disabled>
					<% questionIDs = questionIDs & var_name&"_othertext2" &"|" %>
					</td>
					<%for x = 0 to UBOUND(rankHeaders) %>
						<td align="center" valign="top" class="radioScaleOption" id="cell<%=var_name &"_other"%>|<%=x%>"><input type="radio" name="<%=var_name & "_other2"%>" id="<%=var_name & "_other2"%>|<%=x%>" value="<%=x+1%>" class="slideScaleLabelInput" disabled><br></td>
						
					<%next%>
					<% questionIDs = questionIDs & var_name&"_other2" &"|" %>
				</tr>
			<%end if%>	
		</table>
	</div>
<% End Function 

%>
<%	Sub privateBuildSlideScale(beginNum,endNum,bottomLabel,topLabel,var_name) %>
	<table cellpadding="5" cellspacing="0" border="0" width="100%" class="slideScaleTable" id="slideTable<%=var_name%>">
		<tr class="slideScaleTableRow" id="row<%=var_name%>">
		<%	if bottomLabel <> "" then %>
			<td class="slideScaleTableCellBottomLabel" id="cell<%=var_name%>slidescalebot" width="20%">(<%=bottomLabel%>)</td>
		<%	end if %>
		<%	For scalecount = beginNum to endNum %>
			<td align=""center"" class="slideScaleTableCell" id="cell<%=var_name%>|<%=scalecount%>"><input type="radio" name="<%=var_name%>" id="<%=var_name%>|<%=scalecount%>" value="<%=scalecount%>" class="slideScaleInput"><br><label for="<%=var_name%>|<%=scalecount%>"><%=scalecount%></label></td>
		<% Next %>
		<%if topLabel <> "" then %>
			<td class="slideScaleTableCellTopLabel" id="cell<%=var_name%>slidescaletop" width="20%">(<%=topLabel%>)</td>
		<% end if %>
		</tr>
	</table>
<%	End Sub %>
	
<%	Sub privateBuildSlideScaleSelect(beginNum,endNum,bottomLabel,topLabel,var_name,initSelect) %>
	<table cellpadding="5" cellspacing="0" border="0" width="100%" class="slideScaleSelectTable" id="slideTable<%=var_name%>">
		<tr class="slideScaleSelectTableRow" id="row<%=var_name%>">
		<%	if bottomLabel <> "" then %>
			<td class="slideScaleSelectTableCellBottomLabel" id="cell<%=var_name%>slidescaleselectbot" width="50%">(<%=bottomLabel%>)</td>
		<%	end if %>
		<%	For scalecount = beginNum to endNum %>
			<td align="center" class="slideScaleSelectTableCell" id="cell<%=var_name%>|<%=scalecount%>"><input type="radio" name="<%=var_name%>" id="<%=var_name%>|<%=scalecount%>" value="<%=scalecount%>" <% If initSelect = scalecount Then %>checked disabled<%End If%>  class="slideScaleSelectInput"><br><label for="<%=var_name%>|<%=scalecount%>"><%=scalecount%></label></td>
		<% Next %>
		<%if topLabel <> "" then %>
			<td class="slideScaleSelectTableCellTopLabel" id="cell<%=var_name%>slidescaleselecttop" width="50%">(<%=topLabel%>)</td>
		<% end if %>
		</tr>
	</table>
<%	End Sub %>
	
<%	Sub privateBuildSlideScaleLabel(labelString,var_name)%>
	<table cellpadding="5" cellspacing="0" border="0" width="100%" class="slideScaleLabelTable" id="slideTable<%=var_name%>">
		<tr class="slideScaleLabelTableRow" id="row<%=var_name%>">
		<%
		labelStringAry = Split(labelString,"|")
		Dim labelStringCounter, widthVar
		widthVar = CInt(100/(ubound(labelStringAry)+1))				
		%>
		<% For labelStringCounter = 0 to ubound(labelStringAry) %>
			<td align="center" valign="top" width="<%=widthVar%>%" class="slideScaleLabelTableCell" id="cell<%=var_name%>|<%=labelStringCounter%>"><input type="radio" name="<%=var_name%>" id="<%=var_name%>|<%=labelStringCounter%>" value="<%=(labelStringCounter + 1)%>" class="slideScaleLabelInput"><br><label for="<%=var_name%>|<%=labelStringCounter%>"><%=labelStringAry(labelStringCounter)%></label></td>
		<% Next %>
		</tr>
	</table>
<%	End Sub %>
	
<%	Sub privateBuildSlideScaleLabelOffset(labelString, var_name, val_offset )%>
	<table cellpadding="0" cellspacing="0" border="0" class="slideScaleLabelOffsetTable" id="slideTable2<%=var_name%>">
		<tr class="slideScaleLabelOffsetTableRow" id="row<%=var_name%>">
		<%
		labelStringAry = Split(labelString,"|")
		Dim labelStringCounter, widthVar
		widthVar = CInt(100/(ubound(labelStringAry)+1))					
		%>
		<% For labelStringCounter = 0 to ubound(labelStringAry) %>
			<td align="center" valign="top" width="<%=widthVar%>%" class="slideScaleLabelOffsetTableCell" id="cell<%=var_name%>|<%=(val_offset+labelStringCounter + 1)%>"><input type="radio" name="<%=var_name%>" id="<%=var_name%>|<%=(val_offset+labelStringCounter + 1)%>" value="<%=(val_offset+labelStringCounter + 1)%>" class="radioSlideScaleLabelOffsetInput"><br><label for="<%=var_name%>|<%=(val_offset+labelStringCounter + 1)%>"><%=labelStringAry(labelStringCounter)%></label></td>
		<% Next %>
		</tr>
	</table>
<%	End Sub %>

<%	Sub privateBuildSlideScale_2(beginNum,endNum,bottomLabel,topLabel,var_name) %>
	<table cellpadding="5" cellspacing="0" border="0" width="100%" class="slideScaleTable" id="Table1">
		<tr class="slideScaleTableRow" id="row<%=var_name%>">
		<%	if bottomLabel <> "" then %>
			<td class="slideScaleTableCellBottomLabel" id="cell<%=var_name%>slidescalebot" width="20%">(<%=bottomLabel%>)</td>
		<%	end if %>
		<%	For scalecount = beginNum to endNum %>
			<td align=""center"" class="slideScaleTableCell" id="cell<%=var_name%>|<%=scalecount%>"><input type="radio" name="<%=var_name%>" id="<%=var_name%>|<%=scalecount%>" value="<%=scalecount%>" class="slideScaleInput" disabled><br><label for="<%=var_name%>|<%=scalecount%>"><%=scalecount%></label></td>
		<% Next %>
		<%if topLabel <> "" then %>
			<td class="slideScaleTableCellTopLabel" id="cell<%=var_name%>slidescaletop" width="20%">(<%=topLabel%>)</td>
		<% end if %>
		</tr>
	</table>
<%	End Sub %>