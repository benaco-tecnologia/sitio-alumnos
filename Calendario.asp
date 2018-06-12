<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->


<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title>Portal Alumnos en Línea </title>
<table border="0" cellpadding="0" cellspacing="0"  align="left">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right" ><% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td valign="top" >
			<% CargarTop1()%><% SubMenu()%>
			<table width="100%" border="0" cellspacing="0" cellpadding="20" height="525" bgcolor="#FFFFFF">
			  <tr> 
				<td height="58" valign="top"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#4a5da1"><b>Usted Puede realizar la inscripci&oacute;n en los siguientes d&iacute;as.</b></font> </td>
			  </tr>
              <tr>
              <td valign="top"><img src="Imagenes/Calendario/Calendario.jpg" ></td>
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