<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->

<%
msg=request("msg")
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<body background="ima/iso-uhc.gif">
<table border="0" cellpadding="0" cellspacing="0"  align="center">
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
				<form name="form1" method="post" action="">
				  <p>&nbsp;</p>
				  <table width="92%" border="0">
					<tr> 
					  <td><img name="fder_r2_c1" src="Imagenes/titulos/T-atencion.gif" width="271" height="38" border="0"></td>
					</tr>
					<tr> 
					  <td><%=msg%></td>
					</tr>
					<tr> 
					  <td><font face="Verdana, Arial, Helvetica, sans-serif"><font color="#CC3333"></font></font></td>
					</tr>
					<tr> 
					  <td>&nbsp;</td>
					</tr>
					<tr> 
					  <td>&nbsp;</td>
					</tr>
					<tr> 
					  <td><p><font face="Verdana, Arial, Helvetica, sans-serif"><font color="#CC3333" size="2">Usted 
						  no puede entrar al sitio ya que los periodos de tu plan de estudio no 
						  estan aperturados</font></font></p></td>
					</tr>
					<tr> 
					  <td>&nbsp;</td>
					</tr>
					<tr> 
					  <td>&nbsp;</td>
					</tr>
				  </table>
				  <p>&nbsp;</p>
				</form>
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
