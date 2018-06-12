<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE file="analytics.asp" -->

<%
strParame="SELECT coalesce(dbo.Fn_ValorParame('USAANALYTICS'),'NO')Parame"
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
	USAANALYTICS=rstParame("Parame")
end if

if USAANALYTICS="SI" then
	analitycs() 
end if  
%>

<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if

if EstaHabilitadaNW (899)="S" then 
	if GetPermisoNW(899) ="N" then
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
				session("MensajeBloqueosVarios") ="Para acceder a esta opción debe responder sus Encuestas Pendientes."
				response.Redirect("MensajeBloqueo.asp")
			end if 
		end if 
		
	end if
else
	response.Redirect("MensajeBloqueoHabilita.asp")
end if
Audita 899,"Ingresa a Listado de Solicitudes"
%>

<html>
<head>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title><%=Session("NombrePestana")%></title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1"> 
<body >
<%
strsql="SP_LISTADO_SOLICITUDES_PA '" & session("codcli") & "' "  
Set Rst = Session("Conn").Execute(StrSql) 

strPD="SELECT dbo.Fn_ValorParame('verIngSolicitudes')RutaLogicaDoctos,dbo.Fn_ValorParame('rutaIngSolicitudes')RutaDoctos"
set rstPD= Session("Conn").Execute(strPD)		
if not rstPD.eof then
	RutaDoctos = rstPD("RutaDoctos")
	RutaLogicaDoctos = rstPD("RutaLogicaDoctos")
else
	RutaDoctos="" 
	RutaLogicaDoctos=""	
end if
		


%>
<table border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
	<td>
		<table width="800" border="0" cellpadding="0" cellspacing="0" align="center">
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
                                               Listado de solicitudes</p>
                                                </span></tr>
					  <tr valign="top"> 
						<td colspan="3" height="80">  
						  <table width="801" cellspacing="1" cellpadding="0" height="59" border="0" bordercolor="#FFFFFF">
							<td height="20" colspan="11">&nbsp;</td>
							</tr>
							<tr background="imagenes/fdo-cabecera-cel22.jpg"> 
							  <td width="144" height="30"  background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center" ><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Tipo de solicitud</font></b></font></div></td>
							  <td width="60" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">N&uacute;mero de Solicitud</font></b></font></div></td>
							  <td width="60" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Motivo</font></b></font></div></td>  
                              <td width="60" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Fecha de creaci&oacute;n</font></b></font></div></td>    
                               <td width="60" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Fecha de Resoluci&oacute;n</font></b></font></div></td>    
                               <td width="60" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Estado</font></b></font></div></td>  
                               <td width="130" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Observaci&oacute;n</font></b></font></div></td>                        
							  <td width="70" height="30" background="imagenes/fdo-cabecera-cel22.jpg" NOWRAP> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Adjuntos</font></b></font></div></td>							  
							</tr>
							<% 
						   While Not Rst.Eof
 
						   %>
							<tr bgcolor="#DBECF2" height="25"> 
							  <td width="40" height="30" align="center" ><font face="Arial, Helvetica, sans-serif" size="1"><%=Rst("TIPO")%></font></td>
							  <td width="40" height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=Rst("NUMSOLI")%></font></td>
							  <td width="40" height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=Rst("MOTIVO")%></font></td>
                              <td width="40" height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=Rst("FEC_ASIG")%></font></td>
                              <td width="40" height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=Rst("FEC_APROB")%></font></td>    
                              <td width="40" height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=Rst("ESTADO")%></font></td> 
                              <td width="40" height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=Rst("OBSERVACION")%></font></td>                     
							  <td width="40" height="30" align="center" nowrap="nowrap" align="center">
                              <%if Rst("ADJUNTO1")<>"" then%>                              
                              <%if ExisteArchivo(RutaDoctos & Rst("ADJUNTO1")) = true then%>
                              	<a href="<%=RutaLogicaDoctos & Rst("ADJUNTO1")%>"><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                              <%end if %>
                              <%end if %>
                              
                              <%if Rst("ADJUNTO2")<>"" then%>                              
                              <%if ExisteArchivo(RutaDoctos & Rst("ADJUNTO2")) = true then%>
                              	<a href="<%=RutaLogicaDoctos & Rst("ADJUNTO2")%>"><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                              <%end if %>
                              <%end if %>
                              
                              <%if Rst("ADJUNTO3")<>"" then%>                              
                              <%if ExisteArchivo(RutaDoctos & Rst("ADJUNTO3")) = true then%>
                              	<a href="<%=RutaLogicaDoctos & Rst("ADJUNTO3")%>"><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                              <%end if %>
                              <%end if %>
                              </td>
							</tr> 
							<%
							 Rst.MoveNext
						   Wend %>
                         

						  </table>
						 
						  <p></p></td> 
					  </tr>
                      <tr></tr><tr></tr>
                       <tr>
    <td width="727" align="left"><A onMouseOver="MM_swapImage('Image7','','Imagenes/botones/volver-on.gif',1)" onmouseout=MM_swapImgRestore() href="javascript: window.top.location.href ='menu_consultas.asp'"><IMG src="Imagenes/botones/volver-of.gif" border="0"  name="Image7"></A></td>
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

