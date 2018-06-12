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
if EstaHabilitadaNW (488)="S" then 
	if GetPermisoNW(488) ="N" then
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
'Audita 840,"Ingresa a Programas de asignaturas" 

strANOP="select ANO,PERIODO from MT_CARRERAPERIODO WHERE CODCARR='"& session("codcarr") &"' AND codpestud='"& session("codpestud") &"' "
if bcl_ado(strANOP,rstANOP) then
	ano =valnulo(rstANOP("ANO"),NUM_)
	periodo=valnulo(rstANOP("PERIODO"),NUM_)
end if

strRamosDebe = "sp_listaramosalumnoportal '" & session("codcli") & "','" & ano & "','" & periodo & "'"
RamosDebe.Open strRamosDebe, Session("Conn")
'response.Write(strRamosDebe)
	%>
	
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%=Session("NombrePestana")%></title>
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
                                                <p style="font-size: 25px" class="text-menu" id="lblTitulo">
                                               Programa de Asignaturas</p>
                                                </span></tr>
					  <tr valign="top"> 
						<td colspan="3" height="80">  
						  <table width="801" cellspacing="1" cellpadding="0" height="59" border="0" bordercolor="#FFFFFF">
							<td height="20" colspan="11">&nbsp;</td>
							</tr>
							<tr background="imagenes/fdo-cabecera-cel22.jpg"> 
							  <td width="63" height="30"  background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center" ><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">C&oacute;digo</font></b></font></div></td>
							  <td width="144" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1" id="lblColumnaRamo">Asignatura</font></b></font></div></td>
							  <td width="50" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Secci&oacute;n</font></b></font></div></td>                           
							  <td width="104" height="30" background="imagenes/fdo-cabecera-cel22.jpg" NOWRAP> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Programa</font></b></font></div></td>							  
							</tr>
							<% 
						  If RamosDebe.Eof Then
						  %>
							<%
						  Else
						   While Not RamosDebe.Eof
 
						   %>
							<tr bgcolor="#DBECF2" height="25"> 
							  <td width="63" height="30" align="center" ><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("ramoequiv")%></font></td>
							  <td width="144" height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("Nombre")%></font></td>
							  <td width="50" height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("codsecc")%></font></td>
                 
							 <% 
							 strPr="SELECT dbo.Fn_ValorParame('RUTA_ARCH_PROGR_ASIGNATURAS')Parame"
							 if bcl_ado(strPr,rstPr) then
								Ruta =valnulo(rstPr("Parame"),STR_)
							 end if
							 
							 strTPr="  SELECT COALESCE(ADJUNTO,'')ADJUNTO FROM dbo.RA_RAMO WHERE CODRAMO='"& RamosDebe("ramoequiv") &"'"
							 if bcl_ado(strTPr,rstTPr) then
								Adjunto =rstTPr("ADJUNTO")
							 end if
							 
							 if Ruta <> "" and Adjunto <> "" then%>     
                              	<td width="104" height="30" align="center" nowrap="nowrap">
                                <a href="<%=Ruta+Adjunto%>" target="_blank"><IMG src="Imagenes/editar.gif" border="0" name="conArchivo">                                </td>
							 <%else%>                            
                             	<td width="104" height="30" align="center" nowrap="nowrap"><IMG src="Imagenes/cruz.gif" border="0"  name="SinArchivo">
                                </td>
							 <%end if %> 			  
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
    <td width="727" align="left"><A onMouseOver="MM_swapImage('Image7','','Imagenes/botones/volver-on.gif',1)" onmouseout=MM_swapImgRestore() href="javascript: window.top.location.href ='menu_material.asp'"><IMG src="Imagenes/botones/volver-of.gif" border="0"  name="Image7"></A></td>
  </tr>
	</table>
			</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>  

</body>
<%ObjetosLocalizacion("programasasignatura.asp")%>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
