<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Portal de Alumnos de L&iacute;nea</title>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
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
			  <table width="" border="0" cellspacing="15" >
			  		<form name="form1" method="post" action="">
					<tr> 
					  <td><img name="fder_r2_c1" src="Imagenes/titulos/T-atencion.gif" width="271" height="38" border="0"></td>
					</tr>
					<tr> 
					  <td></td>
					</tr>
					<tr> 
					  <td><p><font face="Verdana, Arial, Helvetica, sans-serif"><font color="#CC3333" size="2"> 
						  <%=session("MensajeBloqueosVarios")%></font></font></p></td>
					</tr>
					<tr> 
					  <td>&nbsp;</td>
					</tr>
					<tr> 
					  <td>&nbsp;</td>
					</tr>
			 	 </form>
			  </table>
			
				<p>&nbsp; </p>
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
</html>
<%
'     "javascript:parent.history.go(0)"
%>
<script languaje="javascript">
function Terminar()
{
<% if session("logueado")="1" then %>
      window.location = "salir.asp";
        //window.top.location = "alumn-udd.asp";
  
<% end if%>   
}
</script>