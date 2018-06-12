<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Portal de Alumnos en L&iacute;nea</title>
<style type="text/css">
<!--
.cssBloque {	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-weight: bold;
}
.cssPregunta {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; }
-->
</style>
</head>

<body>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<table border="0" cellpadding="0" cellspacing="0" align="left">
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
				<table width="100%" border="0" cellpadding="5" cellspacing="3">
				  <tr class="cssBloque">
					<td align="center"><img src="Imagenes/titulos/T-encuesta_contestada.gif"></td>
				  </tr>
				
				  <tr class="cssPregunta">
					<td align="center" valign="middle" bgcolor="#DBECF2"><span style="font-weight: bold">USTED YA RESPONDI&Oacute; LA ENCUESTA    </span></td>
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
