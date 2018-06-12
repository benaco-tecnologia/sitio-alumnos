<%@  language="VBScript" %>
<%
  strParame="SELECT dbo.Fn_ValorParame('PERSONALIZAPA')Parame"
	set rstParame= Session("Conn").Execute(strParame)
	
	if not rstParame.eof then
		salirnuevo=rstParame("Parame")
	else
		salirnuevo=""
	end if 
  	
	'if salirnuevo="SI" then
	''	pagina="alumnosfinning.asp"
	'else
	    strParame="SELECT dbo.Fn_ValorParame('PORTADA_PROPIA_PORTAL_ALUMNO')Parame"
	    set rstParame= Session("Conn").Execute(strParame)
	    if not rstParame.eof then
		        salirnuevo=rstParame("Parame")
	    else
		        salirnuevo=""
	    end if 
	    if(salirnuevo <> "" and not isNull(salirnuevo)) then
	            pagina=rstParame("Parame")
	    else
	             pagina="alumnos.asp"
	    end if 
		
	'end if 
  
  
  Session("Conn").close()
  Set Session("Conn") = Nothing
   
  session.Abandon
  Session("MiValor")=0

%>
<html>

<script>
    window.top.location = "<%=pagina%>";
</script>

</html>
