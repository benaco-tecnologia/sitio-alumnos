<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title>Portal de Alumnos en L&iacute;nea</title>
<body topmargin="0" rightmargin="0" leftmargin="0">
<table border="0" cellpadding="0" cellspacing="0" align="left">
	<td colspan="2" valign="top" align="left" ><% CargarTop()%><% 
		if trim(Session("NomAlum"))="" then
			CargarMenu() 			
		else
			CargarMenu2() 
		end if
		%></td>
</table>
</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->