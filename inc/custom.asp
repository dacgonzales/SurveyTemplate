<% Function CUSTOM_HTML (var_name, e_text) %>
	<% Execute(e_text) %>
<% End Function %>

<%
Function CUSTOM_DATA_HTML (var_name, q_text, e_text)
	'in this function "e_text" stands for "execution text" or code to execute
	DIV_HTML "q"&var_name, q_text, "question"%>
	<div id="diva<%=var_name%>" class="answer">
		<% Execute(e_text) %>
	</div><%
End Function
%>
<% 
Function CUSTOM_DATA_REPEATING_HTML(var_name, q_text, e_text, options)
	'in this function "e_text" stands for "execution text" or code to execute
	ary = split(options, "|")	
	DIV_HTML "q"&var_name, q_text, "question"
	%>
	<div id="diva<%=var_name%>" class="answer">
		<table width="100%"  border="0" cellspacing="0" cellpadding="5" class="customTable" id="table<%=var_name%>">
		<%
		For i = 1 to ubound(ary) + 1
			if i mod 2 = 0 then
				rowClass = "customTableRow1"
			else
				rowClass = "customTableRow2"
			end if
			%>
			<tr class="<%= rowClass %>" id="rowo<%=var_name&"_"&i%>">
				<td id="cello<%=var_name&"_"&i%>" class="customOption"><span class="customNumber"><%=i%>.</span> <%=ary(i-1)%></td>
				<td><% eval(e_text)%></td>
			</tr>
			<%
		Next
		%>
		</table>
	</div>
<%
End Function
%>		

