<% Function FREE_TEXT_HTML( var_name, q_text, e_text ) %>
<script language="javascript" type="text/javascript">
function validate_<%=var_name%>(){
	// validate only if e_text is passed
    item2check = formVar.<%=var_name%>.value;
    $("#<%=var_name%>").removeClass("validation-error");
	if ( ("<%=e_text%>" != "") && (!checkInput(item2check)) ){
	    $("#<%=var_name%>").addClass("validation-error");
			return "<%=e_text%>\n";
	}
	else{
	    
	    return false;
	}
	
}
validateArray[arraycounter] = "validate_<%=var_name%>();";
arraycounter++;
</script>

<% DIV_HTML "q"&var_name, q_text, "question" %>
<div id="diva<%=var_name%>" class="answer">
	<textarea class="freetextInput" id="<%=var_name%>" name="<%=var_name%>" rows=5 style="width:100%; max-width:100%;"></textarea>
	<% questionIDs = questionIDs & var_name &"|" %>
</div>
<% End Function %>