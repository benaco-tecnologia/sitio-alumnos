<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE file="analytics.asp" -->

<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if

strParame="SELECT coalesce(dbo.Fn_ValorParame('USAANALYTICS'),'NO')Parame"
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
	USAANALYTICS=rstParame("Parame")
end if

if USAANALYTICS="SI" then
	analitycs() 
end if 

if EstaHabilitadaNW (762)="S" then
	if GetPermisoNW(762) ="N" then
		'response.Redirect("MensajesBloqueos.asp")	
	else
		strParame="SELECT dbo.Fn_ValorParame('BLOQUEAPAENCUESTAS')Parame"
		set rstParame= Session("Conn").Execute(strParame)		
		if not rstParame.eof then
				BLOQUEAPAENCUESTAS=rstParame("Parame")
			else
				BLOQUEAPAENCUESTAS="" 
		end if
		
		if BLOQUEAPAENCUESTAS="SI" then
			'valida si contesta o no la encuesta
			if TomadeRamo(session("Codcarr"),session("Codcli"),session("ano_ed"),session("periodo_ed"))=0 then
				session("MensajeBloqueosVarios") ="Para acceder a esta opción debe responder sus Encuestas Pendientes."
				'response.Redirect("MensajeBloqueo.asp")
			end if 
		end if 
		
	end if 
else
	response.Redirect("MensajeBloqueoHabilita.asp")
end if 

Audita 762,"Ingresa a Cupon de arancel"
%>

<%
ano =AnoAcad()
periodo=PeriodoAcad()

str="INSERT INTO ra_Documentos_web (RUT,CODCARR,DOCUMENTO,ANO,PERIODO,FECHA,ESTADO,TIPO) VALUES('"& session("rutcli") &"','"& session("codcarr") &"','CUPONCUOTA',"& ano &","& periodo &",GETDATE(),'EN PROCESO','ALUMNO') "

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
				<td valign="top"><p style="font-size:25px" class="text-menu">Cup&oacute;n de Arancel.</p>
				  <table width="800" border="0" cellpadding="0" cellspacing="0">
                 <td height="34">
                      	<a  target="_blank"href="Documentos/<%=session("codcli")& rst("codigo")%>.pdf">
					    <p class="text-menu">Descargar Cup&oacute;n.</p></a></td>
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