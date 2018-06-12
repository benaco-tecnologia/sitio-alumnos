<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%

if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if

ano =AnoAcad()
periodo=PeriodoAcad()

str="INSERT INTO ra_Documentos_web( codcli ,documento ,ANO ,PERIODO ,fecha ,estado,tipo)VALUES('"& session("codcli") &"','PAGARE',"& ano &","& periodo &",GETDATE(),'EN PROCESO','ALUMNO') "
Session("Conn").execute(str)

strident= "SELECT @@IDENTITY codigo " 
set rst = Session("Conn").execute(strident)
 
strexec="sp_genera_documentos_web "& rst("codigo") &""
Session("Conn").execute(strexec)
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
				  <table width="100%" border="0" cellspacing="0" cellpadding="15" height="550" bgcolor="#FFFFFF">
			  <tr> 
				<td valign="top"><p style="font-size:25px" class="text-menu">Pagar&eacute;</p>
				  <table width="800" border="0" cellpadding="0" cellspacing="0">
                 <td height="34">
                      	<a  target="_blank"href="Documentos/<%=session("codcli")& rst("codigo")%>.pdf">
					    <p class="text-menu">Descargar Pagar&eacute;</p></a></td>
					<tr valign="top" bgcolor="#FFFFFF">
					  <td height="15">&nbsp;</td>
					  <td height="15">&nbsp;</td>
				      <td height="15">&nbsp;</td>
					</tr>
					<tr valign="top" bgcolor="#FFFFFF">
					  <td height="100">
					    <p class="text-menu">&nbsp;</p>
					  </td>
					  <td>&nbsp;</td>
				      <td>&nbsp;</td>
					</tr>
                    <tr valign="top" bgcolor="#FFFFFF">
					  <td height="100">
					    <p class="text-menu">&nbsp;</p>
					  </td>
					  <td>&nbsp;</td>
				      <td></td>
					</tr>
                    <tr valign="top" bgcolor="#FFFFFF">
					  <td height="100">
					    <p class="text-menu">&nbsp;</p>
					  </td>
					  <td>&nbsp;</td>
				      <td></td> 
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
	</td>
  </tr>
</table>
</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
<!--javascript:parent.mainframe.history.go(0)-->