<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=RamosDebe name=RamosDebe></OBJECT> 
<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if
%>
<%

ramo = request.form("CR")
nombreramo = request.form("NR")
seccion = request.form("CS")
carrera = request.form("CC")
sede = request.form("SE")
ano = request.form("PE")
periodo = request.form("AN")
codprof = request.form("CP")
nombreprof = request.form("NP")
asistencia = request.form("AS")

if EstaHabilitadaNW (896)="S" then 
	if GetPermisoNW(896) ="N" then
		response.Redirect("MensajesBloqueos.asp")	
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
				session("MensajeBloqueosVarios") ="Para acceder a esta opci�n debe responder sus Encuestas Pendientes."
				response.Redirect("MensajeBloqueo.asp")
			end if 
		end if 
		
	end if
else
	response.Redirect("MensajeBloqueoHabilita.asp")
end if 
'Audita 896,"Ingresa a detalle de asistencia por fechas" 

strRamosDebe = "SP_LISTA_ASISTENCIA_DET_PORTAL '" & session("codcli") & "','"& sede &"','"& codprof &"','"& ramo & "',"& seccion &","& ano & ","& periodo &" "
RamosDebe.Open strRamosDebe, Session("Conn")
'response.Write(strRamosDebe)
	%>
	
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
<link rel="stylesheet" href="file://///Gotzilla/uddweb/admalumnos/css/tex-normales.css" type="text/css">
<link href="css/tex-normales.css" rel="stylesheet" type="text/css">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1"> 
<style type="text/css">
<!--
.Estilo19 {
	color: #FFFFFF;
	font-weight: bold;
}
.Estilo20 {color: #FFFFFF; font-weight: bold; font-size: 16px; }
-->
</style>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<link href="css/textos.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.Estilo21 {color: #000000}
.Estilo22 {font-size: 10px}
-->
</style>
<script language="JavaScript" type="text/javascript">


function CargarReporte()
{
	var opciones="toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=no, resizable=no, width=1024, height=768, top=85, left=125";
window.open("cuponera.asp","",opciones);
}


</script>
</head>
<body>
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
                 <table border="0" cellpadding="0" cellspacing="15" width="757">
					  <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
					  <tr valign="top" > 
						<td > <span style="font-size: 14px">
                                                <p style="font-size: 25px" class="text-menu">
                                               Detalle de asistencia por Asignaturas</p>
                                                </span></tr>
					  <tr valign="top"> 
						<td colspan="3" height="80">  
                        
                          <table width="801" cellspacing="1" cellpadding="0" height="59" border="0" bordercolor="#FFFFFF">
                            <tr>
                              <td height="30" background="imagenes/fdo-cabecera-cel.jpg" class="text-cabecera-celda"><span class="nnEstilo1 Estilo4"><strong>Profesor:</strong></span></td>
                             <td width="88%" height="30" colspan="6" bgcolor="#DBECF2" class="text-normal-celdas"><%=UCASE(nombreprof)%></td>
                            </tr>
                            <tr>
                              <td width="12%" height="30" background="imagenes/fdo-cabecera-cel.jpg" class="text-cabecera-celda"><span class="nnEstilo1 Estilo3">Asignatura:</span></td>
                              <td height="30" colspan="6" bgcolor="#DBECF2" class="text-normal-celdas"><%=ramo & " - "& UCASE(nombreramo)%></td>
                            </tr>
                              <tr>
                              <td width="12%" height="30" background="imagenes/fdo-cabecera-cel.jpg" class="text-cabecera-celda"><span class="nnEstilo1 Estilo3">Secci&oacute;n:</span></td>
                              <td height="30" colspan="6" bgcolor="#DBECF2" class="text-normal-celdas"><%=seccion%></td>
                            </tr>
                            </tr>
                              <tr>
                              <td width="12%" height="30" background="imagenes/fdo-cabecera-cel.jpg" class="text-cabecera-celda"><span class="nnEstilo1 Estilo3">Asistencia</span></td>
                              <td height="30" colspan="6" bgcolor="#DBECF2" class="text-normal-celdas"><%=asistencia%></td>
                            </tr>
                          </table>
                            
						  <table width="801" cellspacing="1" cellpadding="0" height="59" border="0" bordercolor="#FFFFFF">
							<td height="20" colspan="11">&nbsp;</td>
							</tr>
							<tr background="imagenes/fdo-cabecera-cel22.jpg"> 
							  <td width="190" height="30"  background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center" ><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Fecha</font></b></font></div></td>
							  <td width="190" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">M&oacute;dulo</font></b></font></div></td>
							  <td width="190" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Sala</font></b></font></div></td>  
                              <td width="226" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Estado</font></b></font></div></td>    
							</tr>
							<% 
						  If RamosDebe.Eof Then
						  %>
							<%
						  Else
						   While Not RamosDebe.Eof
 
						   %>
							<tr bgcolor="#DBECF2" height="18"> 
                            <%
								if RamosDebe("estado") ="Ausente" then
									fondo = "background='Imagenes/fondo_amarillo.gif'"
								else
									fondo = ""
								end if 
							%>
							  <td width="190" height="15" align="center" <%=fondo%>><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("fecha")%></font></td>
							  <td width="190" height="15" align="center" <%=fondo%>><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("codmod")%></font></td>
							  <td width="190" height="15" align="center" <%=fondo%>><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("sala")%></font></td>
							  <td width="226" height="15" align="center" <%=fondo%>><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("estado")%></font></td>
							</tr> 
							<%
							 RamosDebe.MoveNext
						   Wend
						 End If %>
                         

						  </table>
						 
						  <p></p></td> 
					  </tr>
                      <tr></tr><tr></tr>
                       <tr>
    <td width="727" align="left"><A onMouseOver="MM_swapImage('Image7','','Imagenes/botones/volver-on.gif',1)" onmouseout=MM_swapImgRestore() href="javascript: window.top.location.href ='asistenciasalumnos.asp'"><IMG src="Imagenes/botones/volver-of.gif" border="0"  name="Image7"></A></td>
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
