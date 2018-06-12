<!--#INCLUDE FILE="include/conexion.inc" -->
<%
strM ="sp_validaSubsecciones '"& session("codcli") &"',"& session("ano") &","& session("periodo")&""
if BCL_ADO(strM,rstM) then   
	DEBE = rstM("DEBE")
end if

if DEBE = "SI" then 
%><script languaje="javascript"> 
	alert("debe seleccionar las subsecciones asociada a la secci\u00f3n t\u00e9orica");
	window.location="asignatura-seccion.asp"; 
</script><%
end if
%>
 
<%
strR ="sp_validaRequisitosAprobInscrito '"& session("codcli") &"',"& session("ano") &","& session("periodo")&"" 'SEGUIR !!!
if BCL_ADO(strR,rstR) then   
	Codramo = rstR("codramo")
	Requisito = rstR("Requisito")
end if

if not rstR.eof then 
%><script languaje="javascript"> 
	alert("Debe seleccionar la asigntura <%=Requisito%> para poder inscribir la asignatura <%=Codramo%>.\nSe quitar\u00e1 marca de selecci\u00f3n para la asignatura <%=Codramo%>.");  
	window.location="asignatura-seccion.asp"; 
</script><%
end if
%>

<script languaje="javascript"> 
	window.location="inscrip-asigna.asp"; 
</script>