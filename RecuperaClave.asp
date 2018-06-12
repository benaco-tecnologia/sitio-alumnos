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
                  <%
				  TieneArroa=0
					for i=1 to len(session("RecuperaClaveMail"))
						letra=mid(session("RecuperaClaveMail"),i,1)
						if letra="@" then
							TieneArroa=1
						end if
					next
					
					if TieneArroa =1 then%>
						  <td><p class="parrafo"><font color="#CC3333" size="2" face="Verdana, Arial, Helvetica, sans-serif">Su clave de acceso al Portal de Alumnos ha sido enviada al E-Mail: <%
						  
						  largoNombre=InStr(session("RecuperaClaveMail"), "@")					  
						  dominio=Mid(session("RecuperaClaveMail") ,largoNombre,len(session("RecuperaClaveMail")))	
	
						  PNombre=Mid(session("RecuperaClaveMail") ,1,(largoNombre-1)/2)	
						  
						  'response.Write(dominio)
						  'response.Write(PNombre)
						  
						  for i=1 to (largoNombre-1)/2
							x=x &"*"
						  next 
						  response.Write(PNombre & "" & x & "" & dominio)
					  %> 
                      </font></p></td>
                      <%else%>
                      <td><p class="parrafo"><font color="#CC3333" size="2" face="Verdana, Arial, Helvetica, sans-serif">Su cuenta de correo electrónico asociado a su cuenta no es válido, Comuníquese con la unidad académica para actualizar sus datos.</font></p></td>
                      <%end if%>
                      
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
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
<!--javascript:parent.mainframe.history.go(0)-->