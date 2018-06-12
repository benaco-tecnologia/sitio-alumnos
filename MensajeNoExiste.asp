<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%
msg=request("msg")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/css/parrafo.css" rel="stylesheet" type="text/css">
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<body >
<table border="0" cellpadding="0" cellspacing="0"  align="left">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right"><% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td valign="top" >
			<% CargarTop1()%><% SubMenu()%>
				  <table width="92%" border="0" cellpadding="0" cellspacing="15" >
				  <form name="form1" method="post" action="">
					<tr> 
					  <td ><img src="Imagenes/titulos/T-atencion.gif" width="271" height="38"></td>
					</tr>
					<tr> 
					  <td><%=msg%></td>
					</tr>
					<tr> 
					  <td><p class="parrafo"><font color="#CC3333" size="2" face="Verdana, Arial, Helvetica, sans-serif" id="lblMensaje">Lo sentimos el Rut que ha ingresado no se encuentra en nuestros registros. Vuelva a Intentar</font></p></td>
					</tr>
					<tr> 
					  <td>&nbsp;</td>
					</tr>
					<tr> 
					  <td>&nbsp;</td>
					</tr>
					<tr> 
					  <td>&nbsp;</td>
					</tr>
					<tr> 
					  <td>&nbsp;</td>
					</tr>
					<tr>
					  <td>&nbsp;</td>
					</tr>
					</form>
				  </table>
				  <p>&nbsp;</p>
				<p>&nbsp;</p>
				<p>&nbsp;</p>
			  </tr>
			</table>
			</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>
</body>
<%ObjetosLocalizacion("mensajenoexiste.asp")%>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
<!--javascript:parent.mainframe.history.go(0)-->