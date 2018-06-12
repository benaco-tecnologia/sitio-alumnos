<%
msg=request("msg")
%>
<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenuExterno.asp" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE file="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Portal de Alumnos en Línea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<body background="ima/iso-uhc.gif">
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
				  <table width="92%" border="0" cellpadding="0" cellspacing="15">
				  <form name="form1" method="post" action="">
					<tr> 
					  <td><img src="Imagenes/titulos/T-atencion.gif" name="fder_r2_c1" width="271" height="38" border="0"></td>
					</tr>
					<tr> 
					  <td><%=msg%></td>
					</tr>
					<tr> 
					  <td><p><font color="#CC3333" size="2" face="Verdana, Arial, Helvetica, sans-serif">Clave 
						  incorrecta ..!!! Vuelva a Intentar</br></br>(Recuerde que el campo password es sensible a mayúsculas y minúsculas)</font></p></td>
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
			
				<p>&nbsp; </p>
				<p>&nbsp;</p>
				<p>&nbsp;</p>
			</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>
</body>
</html>
<!--javascript:parent.mainframe.history.go(0)-->