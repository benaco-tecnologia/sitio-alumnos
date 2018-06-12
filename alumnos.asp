<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top12.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenuExternoPrincipal.asp" -->
<!--#INCLUDE file="cuerpo.asp" -->
<%
	Session("RutCli")=""
	Session("RutCliente")= ""
	Session("MiClave")=""
	Session("ClaveEncriptada")=""
	Session("MiValor")=0
	Session("NomAlum") = ""
	Session("RutAlum") = ""
	Session("Rut") = ""
	Session("CodCli") = ""
	session("Codsede")=""
	session("Codcarr")=""
	Session("codpestud")=""
	Session("anoalumno") = ""
	Session("Jornada") = ""
	Session("Nivel") = ""
    Session("Logueado") = ""
	Session("BloqueoA") = ""
	Session("BloqueoF") = ""
	Session("estado") = ""
	Session("EstadoMatriculado") = ""
	session("Cerrada") = ""
	Session("GetCarr") = ""
	session("CarreraAlumno") = ""
	session("NumeroMaximoPrueba") = ""
	Session("destino")= Request("destino")
    Session("pagar")= Request("pagar")

if session("TIPOVALIDACIONRUT") ="PERUANA" then 
	Titulo="Portal de estudiantes en L&iacute;nea"
else
	Titulo="Portal de Alumnos en L&iacute;nea"
end if 
%>


<!DOCTYPE html>
<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title><%=Titulo%></title>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<body>
<table width="1024" border="0" align="left" cellpadding="0" cellspacing="0">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0">
		  <tr>
			<td width="0" valign="top" align="left" ><% CargarTop()%><% CargarMenu()%></td>
		    <td valign="top" align="left" ><% CargarTop1()%><% CargarCuerpo()%></td>
		  </tr>
		</table>
	</td>
  </tr>
</table>
</body>
</html>

