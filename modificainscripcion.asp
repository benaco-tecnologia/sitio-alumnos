<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%
str ="sp_elimna_inscripcion '"& session("codcli") &"',"& session("ano") &","& session("periodo")&" "
Session("Conn").execute(str)

Audita 470, "Acepta Modificar Inscripción" 
%>
<script languaje="javascript"> 
	window.location="frame-inscrip-asigna.asp"; 
</script>