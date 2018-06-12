<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->

<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title>Portal Docente en Línea </title>
<table border="0" cellpadding="0" cellspacing="0"  align="center">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" ><% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td valign="top" align="center">
			<% CargarTop1()%><% SubMenu()%>
			<table width="100%" border="0" cellspacing="0" cellpadding="20" height="550" bgcolor="#FFFFFF">
			  <tr> 
				<td valign="top"><p><img src="imagenes/T-asistencia.gif" width="136" height="38"></p>
				  <table width="684" border="0" cellpadding="0" cellspacing="0" bgcolor="#333333">
					<tr valign="top" bgcolor="#FFFFFF">
					  <td width="671" height="100"><a href="">
					    <p class="text-menu">Inscripci&oacute;n Normal de Asignaturas </p>
					  </a> 
				        <p class="text-normal-celdas">Lorem ipsum dolor sit amet, consectetuer adipiscing elit. Integer nibh ligula, scelerisque non, elementum et, malesuada fringilla, nunc. Curabitur augue. Nullam id </p></td>
					  <td width="13">&nbsp;</td>
				    </tr>
					<tr valign="top" bgcolor="#FFFFFF">
					  <td height="100"><a href="">
					    <p class="text-menu">Solicitud Especial de Inscripci&oacute;n de Asignaturas </p>
					  </a>
                        <p class="text-normal-celdas">Augue. Fusce massa enim, egestas a, vehicula id, tempor sed, metus. Fusce euismod gravida velit. Aliquam ut odio non massa placerat aliquam</p></td>
					  <td>&nbsp;</td>
				    </tr>
					<tr valign="top" bgcolor="#FFFFFF">
					  <td height="100"><a href="">
					    <p class="text-menu">Resumen de Inscripcion de Asignaturas </p>
					  </a>
                        <p class="text-normal-celdas">Augue. Fusce massa enim, egestas a, vehicula id, tempor sed, metus. Fusce euismod gravida velit. Aliquam ut odio non massa placerat aliquam</p></td>
					  <td>&nbsp;</td>
				    </tr>
					<tr valign="top" bgcolor="#FFFFFF">
					  <td height="100"><a href="">
					    <p class="text-menu">Malla Curricular </p>
					  </a>
                        <p class="text-normal-celdas">Augue. Fusce massa enim, egestas a, vehicula id, tempor sed, metus. Fusce euismod gravida velit. Aliquam ut odio non massa placerat aliquam</p></td>
					  <td>&nbsp;</td>
				    </tr>
					<tr valign="top" bgcolor="#FFFFFF">
					  <td height="100">&nbsp;</td>
					  <td>&nbsp;</td>
				    </tr>
			    </table>      
			      <table width="738">
                    <tr>
                    </tr>
                    <tr>
                      </tr>
                  </table>
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