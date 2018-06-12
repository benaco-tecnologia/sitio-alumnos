<!--#INCLUDE FILE="include/conexion.inc" -->
<%

rut = replace(replace(request.Form("rut"),".",""),"-","")
nombre = request.Form("nombre")
apellido = request.Form("apellido")
direccion = request.Form("direccion")
email = request.Form("email")
'confirmaEmail = request.Form("confirmaEmail")
telFijo = request.Form("telFijo")
telMovil = request.Form("telMovil")
nombreTutor = request.Form("nombreTutor")
telTutor = request.Form("telTutor")
comentarios = request.Form("comentarios")
if telFijo ="" then
telFijo=0
end if 

	str="TUTORES_in '"& rut &"','"& nombre &"','"& apellido &"','"& direccion &"' ,'"& email &"',"& telFijo &","& telMovil &",'"& nombreTutor &"',"& telTutor &",'"& comentarios &"'"

	Session("Conn").execute(str)
%>
<script languaje="javascript"> 
	alert("La operaci\u00f3n ha finalizado exitosamente.");
	window.top.location.href = "formularioinscripcion.asp"; 
</script>
