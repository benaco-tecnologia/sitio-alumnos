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

if EstaHabilitadaNW (483)="S" then 
	if GetPermisoNW(483) ="N" then
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
Audita 483,"Ingresa a Consulta Curricular"
%>
<%

dim strsql,strsql2,sql 
codcli=session("codcli")
%>

<html>
<head>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title><%=Session("NombrePestana")%></title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1"> 
<body >
<%
strsql=" SP_CONSULTA_CURRICULAR_PA '"& session("codcli") &"' " 

'response.Write(strsql)
'SEGUIR ra_Seccio pal tipocurso yunion para ra_notaactividad!!! 
'response.End()
Set Rst = Session("Conn").Execute(StrSql) 
%>
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
			<table width="100%" border="0" cellspacing="0" cellpadding="20" height="493" bgcolor="#FFFFFF">
			  <tr> 
				<td height="38" valign="top"><p><img src="imagenes/titulos/T-consulta-curricular.gif" width="400" height="38"></p>
				  <table width="603" border="0" cellpadding="0" cellspacing="0" >
				    <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
				    <tr valign="bottom">
				      <td colspan="2" class="tex-normales"><font id="lblTextAsign" color="#333333" size="2" face="Arial, Helvetica, sans-serif" class="datos-informe">Asignaturas Inscritas sin nota final.</font></td>
			        </tr>
				    <tr valign="top">
				      <td colspan="2" height="8">&nbsp;</td>
			        </tr>
				    <tr valign="top">
				      <td colspan="2" height="70"><table width="604" cellspacing="2" cellpadding="3" height="72" border="0" bordercolor="#FFFFFF">
				        <tr class="text-cabecera-celda" background="imagenes/fdo-cabecera-cel22.jpg" >
				          <td width="104" height="30" height="30" background="imagenes/fdo-cabecera-cel22.jpg" class="textos"><div align="center" class="Estilo2"><font face="Verdana, Arial, Helvetica, sans-serif"><font id="lblCodigoRamo" size="1">C&oacute;digo 
				            Ramo </font></font></div></td>
				          <td width="208" height="30"height="30" background="imagenes/fdo-cabecera-cel22.jpg" class="textos"><div align="center" class="Estilo2"><font face="Verdana, Arial, Helvetica, sans-serif"><font id="lblNombreRamo" size="1">Nombre 
				            del Ramo</font></font></div></td>
				          <td width="59" height="30" height="30" background="imagenes/fdo-cabecera-cel22.jpg" class="textos"><div align="center" class="Estilo2"><font face="Verdana, Arial, Helvetica, sans-serif"><font id="lblPeriodo" size="1">Per&iacute;odo</font></font></div></td>
				          <td width="59" height="30" height="30" background="imagenes/fdo-cabecera-cel22.jpg" class="textos"><div align="center" class="Estilo2"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1">A&ntilde;o</font></font></div></td>
                          <td width="59" height="30" height="30" background="imagenes/fdo-cabecera-cel22.jpg" class="textos"><div align="center" class="Estilo2"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1">Asistencia</font></font></div></td>
				          <td width="82" height="30" height="30" background="imagenes/fdo-cabecera-cel22.jpg" class="textos"><div align="center" class="Estilo2"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1">Tipo Curso</font></font></div></td>
				          </tr>
				        <%do while not rst.eof %>
				        <%
							if not rst.eof then
								codigoramo=rst("Codramo")
								Nombre=rst("Nombre")
								tipocurso=rst("tipocurso")
								periodo=rst("periodo")
								ano=rst("ano")                                    
								Asistencia=rst("asist_fecha")  
							end if
                         %>
				        <tr bgcolor="ffc172">
				          <td width="104" height="32" bgcolor="#DBECF2" class="textos"><div align="left" class="Estilo4"><font face="Arial, Helvetica, sans-serif" size="1"><%=codigoramo%></font></div></td>
                          <td width="104" height="32" bgcolor="#DBECF2" class="textos"><div align="left" class="Estilo4"><font face="Arial, Helvetica, sans-serif" size="1"><%=Nombre%></font></div></td>
                          <td width="59" height="32" bgcolor="#DBECF2" class="textos"><div align="right" class="Estilo4"><font face="Arial, Helvetica, sans-serif" size="1"><%=periodo%></font></div></td>
				          <td width="59" height="32" bgcolor="#DBECF2" class="textos"><div align="right" class="Estilo4"><font face="Arial, Helvetica, sans-serif" size="1"><%=ano%></font></div></td>
                          <td width="59" height="32" bgcolor="#DBECF2" class="textos"><div align="right" class="Estilo4"><font face="Arial, Helvetica, sans-serif" size="1"><%=Asistencia%></font></div></td>
				          <td width="82" height="32" bgcolor="#DBECF2" class="textos"><div align="left" class="Estilo4"><font face="Arial, Helvetica, sans-serif" size="1"><%=tipocurso%></font></div></td>
				          <%  rst.movenext %>
				          </tr>
				        <% loop%>
			          </table></td>
			        </tr>
				    <tr valign="top"  align="left">
						<td valign="top"><a href="menu_consultas.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image15','','Imagenes/botones/volver-on.gif',1)"><img src="Imagenes/botones/volver-of.gif" width="178" height="45" name="Image15" border="0"></a></td>
			        </tr>
		        </table></td>
			  </tr>
              </table>
            </td>
		  </tr>
		</table>
	</td>
  </tr>
</table>
</body>
<%ObjetosLocalizacion("ConsultaCurricular.asp")%>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
