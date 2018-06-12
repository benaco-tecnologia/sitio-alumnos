<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->

<%
	sql = " select COALESCE (dbo.Fn_ValorParame('ACTIVACALENDARIOALUMNO'), 'N') AS ACTIVACALENDARIOALUMNO"

	Set Rst = Session("Conn").Execute(sql)
		ACTIVACALENDARIOALUMNO = Valnulo(Rst("ACTIVACALENDARIOALUMNO"), str_)
	Rst.close()
	
	if ACTIVACALENDARIOALUMNO = "S" then 
		response.Redirect("Calendario.asp")
		response.End()
	end if
	
%>	

<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title><%=Session("NombrePestana")%></title>
<body >
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
			<table width="100%" border="0" cellspacing="0" cellpadding="20" height="550" bgcolor="#FFFFFF">
			  <tr> 
				<td valign="top"><font size="3" face="Verdana, Arial, Helvetica, sans-serif" color="#4a5da1"><b id="lblMensaje">No hay informaci&oacute;n, ya que usted no es un alumno vigente.</b></font> 
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
<%ObjetosLocalizacion("noaccess.asp")%>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->