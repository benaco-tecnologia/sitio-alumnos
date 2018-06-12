 
<%
    Response.Expires = -1
    Session.Timeout = 999
    Response.Buffer = false
    Server.ScriptTimeout = 99999
%>
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<% 
  session("CodProf")= Decode(request("CodProfe"))
  session("TipoDoc")=request("Tipo")
  session("Nombre")= request("Nom")
%>
<html>
<head>
<title>evaluaci&oacute;n Docente</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset rows="190,*" cols="*" framespacing="0" frameborder="NO" border="0">
  <frame name="bottomFrame2" scrolling="NO" src="pregunta-opc.asp">
  <frame name="bottomFrame3" src="list-preguntas.asp" scrolling="AUTO">
</frameset>
<noframes> 
 </noframes> 
</html>
<!--#INCLUDE file="include/desconexion.inc" -->