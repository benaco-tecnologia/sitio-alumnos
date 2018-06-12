<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
  response.End()
end if
%>

<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title>Portal Docente en Línea </title>
<body >
<table border="0" cellpadding="0" cellspacing="0"  align="left">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right" ><% CargarTop()%><% 
				if Session("CodCli") = "" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td valign="top" >
			<% CargarTop1()%><% SubMenu()%>
			<table width="100%" border="0" cellspacing="0" cellpadding="20" height="550" bgcolor="#FFFFFF">
			  <tr> 
				<td valign="top"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#4a5da1"><b>El Alumno a&uacute;n no acepta la Carga Acad&eacute;mica; para los Alumnos Nuevos, la generaci&oacute;n de Horarios se encuentra actualmente en proceso.</b></font> 

		        <p>&nbsp;</p></td>
			  </tr>
			</table>
			</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>
</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->