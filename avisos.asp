<!--#INCLUDE file="top.asp" --> 
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#include file="addon/includes/libreria.asp" -->
<!--#include file="vars.inc.asp" -->
<%

Exe = cInt(Request.QueryString("Cmd"))
if Exe=0 then Exe = 1 end if
%>
 
<% 
Set RsAyP = Session("Conn").Execute("SELECT dbo.Fn_ValorParame('ano') as ano,dbo.Fn_ValorParame('periodo') as periodo ")
if Not RsAyP.Bof or Not RsAyP.Eof then
	RsAyP.MoveFirst
	While Not RsAyP.Eof
	  ano=RsAyP(0)
	  periodo=RsAyP(1)
	  RsAyP.MoveNext  
	wend
End if

if request.QueryString("IDCDAUMAS")<>"" _
	then
		Session ("CarreraEnCurso")=Session("RutAlum")
end if

strParame="SELECT dbo.Fn_ValorParame('MATERIAL_SECCION')Parame"
if BCL_ADO(strParame, rstParame) then
	MaterialSeccion=trim(valnulo(rstParame("Parame"),STR_))
else
	MaterialSeccion=""
end if

'Set Rs1 = Session("Conn").Execute("SELECT DISTINCT ano FROM ra_nota WHERE codcli='" & Session ("CarreraEnCurso") & "'  and ano='"&trim(ano)&"'  and periodo='"&trim(periodo)&"' ORDER BY ano")

Set Rs1 = session("conn").Execute("SELECT distinct s.ano FROM dbo.RA_SECCIO s, ra_nota n WHERE n.CODCLI = '" & Session ("CarreraEnCurso") & "' AND n.RAMOEQUIV = s.CODRAMO AND n.codsecc = s.CODSECC AND n.CODSEDE = s.CODSEDE AND n.ano = s.ANO AND n.PERIODO = s.PERIODO AND CERRADA = 'N' order by s.ano")

if  Rs1.bof or Rs1.eof  then 
	response.Redirect("MensajeSinDatos.asp")
	response.end()
end if 

if EstaHabilitadaNW (489)="S" then 
	if GetPermisoNW(489) ="N" then
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
				session("MensajeBloqueosVarios") ="Para acceder a esta opci?n debe responder sus Encuestas Pendientes."
				response.Redirect("MensajeBloqueo.asp")
			end if 
		end if 
		
	end if
else
	response.Redirect("MensajeBloqueoHabilita.asp")
end if 
Audita 489,"Ingresa a Avisos ingresados"  


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%=Session("NombrePestana")%></title> 
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1"> 
<link rel="stylesheet" href="css/tex-normales.css" type="text/css">
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<style type="text/css">
<!--
.tex-normales {  font: normal 9px Verdana, Arial, Helvetica, sans-serif; color: #000000; letter-spacing: -0.5px}
table{font-family:Arial, Helvetica, sans-serif;font-size:11px}
-->
</style>
<script src="addon/includes/enac.js"></script>
<body onLoad="">    
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
			<td valign="top">
			<% CargarTop1()%><% SubMenu()%>
			<p>&nbsp;</p>
            <table width="770" border='1' cellpadding='5' cellspacing='0' align="center" > 
				  
<%
Set Rs2 = Session("Conn").Execute("SELECT a.ANO, a.PERIODO, a.RAMOEQUIV as CODRAMO from ra_nota as a where a.codcli='" & Session ("CarreraEnCurso") & "' and a.RAMOEQUIV<>'' and a.ano='" & Rs1("ANO") & "' and a.periodo='1' order by a.ano, a.periodo")
if Not Rs1.Bof or Not Rs1.Eof then
	Rs1.MoveFirst
	While Not Rs1.Eof%>
	

		
			<tr>
					<td class="caption" background="Imagenes/fdo-cabecera-cel22.jpg" height="30"><span class="text-cabecera-celda"><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs1("ANO")%>');"><b>Avisos Publicados A&ntilde;o <%=Rs1("ANO")%></b></a></span></td>
				  </tr>
				  <tr>
					<td valign='top' id='<%=Rs1("ANO")%>'><table width="100%" border='1' cellpadding='0' cellspacing='0'>
						<tr>
						  <td width='5%' align='center' valign='top' bgcolor="#DBECF2" class="tex-totales-celda"><span class="normalname"><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs1("ANO") & "_I" %>');">I Semestre</a></span></td>
						  <td width='95%' valign='top'><span class="normalname">
							<%					
				TablaRamos=""
				TablaProfesores=""
				TablaDocumentos=""
				%>
							</span>
							  <table width='100%' border='1' cellpadding='0' cellspacing='0' id='<%=Rs1("ANO") & "_I" %>' >
								<tr>
								  <td valign='top'>
									<%					
				Set Rs2 = Session("Conn").Execute("SELECT a.ANO, a.PERIODO, a.RAMOEQUIV as CODRAMO from ra_nota as a where a.codcli='" & Session ("CarreraEnCurso") & "' and a.RAMOEQUIV<>'' and a.ano='" & Rs1("ANO") & "' and a.periodo='1' order by a.ano, a.periodo")
				if Not Rs2.Bof or Not Rs2.Eof _
					then
						Rs2.MoveFirst
						While Not Rs2.Eof
										 
							Set Rs3 = Session("Conn").Execute("SELECT NOMBRE from ra_ramo where codramo='" & Rs2("CODRAMO") & "'")
							
						if MaterialSeccion ="SI" then				
							Set Rs4 = Session("Conn").Execute("select c.CODPROFE codprof from ca_avisos c, ra_nota n,ra_horascont h where n.codcli = '"& session("codcli") &"' and c.CodRamo='" & Rs2("CODRAMO") & "' and c.Ano='"&trim(ano)&"' And c.Periodo='1' and c.codcarr in (select codcarr COLLATE Modern_Spanish_CI_AS from mt_Carrer where sede='"& session("codsede") &"') AND c.CODRAMO=n.RAMOEQUIV COLLATE SQL_Latin1_General_CP1_CI_AS AND c.ano=n.ANO AND c.periodo=n.PERIODO AND c.CODPROFE=h.CODPROF COLLATE SQL_Latin1_General_CP1_CI_AS AND c.CODRAMO =h.CODRAMO COLLATE SQL_Latin1_General_CP1_CI_AS AND n.codsecc = h.codsecc AND c.ano=h.ANO AND c.periodo=h.periodo AND h.CODSEDE='"& session("codsede") &"' GROUP BY c.CODPROFE")
						else
							Set Rs4 = Session("Conn").Execute("select DISTINCT codprofe codprof from ca_avisos where codramo='" & Rs2("CODRAMO") & "' and Ano='"&trim(ano)&"' And Periodo='1' and codcarr in (select codcarr COLLATE Modern_Spanish_CI_AS from mt_Carrer where sede='"& session("codsede") &"')")
						end if
				%>
									  <table width='100%' border='1' cellpadding='3' cellspacing='0' class='smalltext'>
										<tr>
										  <td align="left" valign="middle" bgcolor='#FFFFCC'><span class="normalname"><a style='text-align:left; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs2("ANO") & "_" & Rs2("PERIODO") & "_" & Rs2("CODRAMO")%>');">
										  <%if Not Rs4.Bof or Not Rs4.Eof then response.Write("<img src='addon/imagenes/IconoMensaje[1].gif' width='16' height='16'>") end if%>&nbsp;&nbsp;<%=Rs2("CODRAMO")%>):<%=Rs3("NOMBRE")%></a></span></td>
										</tr>
										<tr>
										  <td valign='top' style='display:none ;' id='<%=Rs2("ANO") & "_" & Rs2("PERIODO") & "_" & Rs2("CODRAMO")%>'>
											<%
				TablaProfesores="" ' 
	
				if Not Rs4.Bof or Not Rs4.Eof _
					then
						Rs4.MoveFirst
						While Not Rs4.Eof
							Set Rs6=Session("Conn").Execute("select * from ra_profes where codprof='" & Rs4("CodProf") & "'")
							if Not Rs6.Bof or Not Rs6.Eof _ 
								then%>
											  <table width='100%' border='1' cellpadding='0' cellspacing='0' class='smalltext'>
												<tr>
												  <td class='maintitle'><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs2("CODRAMO") & "_" & Rs6("CODPROF")%>');"><%=Rs6("NOMBRES") & " " & Rs6("AP_PATER") & " " & Rs6("AP_MATER")%></a></td>
												</tr>
												<tr>
												  <td valign='top' style='display:none ;' id='<%=Rs2("CODRAMO") & "_" & Rs6("CodProf")%>'>
													<%
				TablaDocumentos=""
				HeadDoctos=""
				BodyDoctos=""
				FooterDoctos=""
				Set Rs5=Session("Conn").Execute("select Fecha,Descripcion, IdDocumento from ca_avisos where CodRamo='" & Rs2("CODRAMO") & "' and ano='"&trim(ano)&"' and CodProfe='" & Rs6("CodProf") & "' ")
				if Not Rs5.Bof or Not Rs5.Eof _
					then
						Rs5.MoveFirst%>
													  <table width='100%' border='1' cellpadding='2' cellspacing='0'>
														<tr>
														  <td width='10%'  align='center' valign='middle'><span class="normalname">Fecha Publicaci&oacute;n</span></td>
														  <td  width='30%' ><span class="normalname">T&iacute;tulo</span></td>
														  <td  width='60%' ><span class="normalname">Descripci&oacute;n</span></td>
														</tr>
														<%
						While Not Rs5.Eof%>
														<tr>
														  <td align='center' valign='middle'><span class="normalname"><%=Day(cDate(Rs5("Fecha"))) & "/" & Month(cDate(Rs5("fecha"))) & "/" & year(cDate(Rs5("fecha")))%></span></td>
														  <td width='30%' align="left" valign="middle"><span class="normalname"><%=Rs5("Descripcion")%></span>&nbsp;</td>
														  <td width='60%' align="left" valign="top"><%=Rs5("IdDocumento")%>
														  </td>
														</tr>
														<%
							Rs5.MoveNext
						Wend%>
													  </table>
													  <span class="normalname">
													  <%
					else%>
													  <!--(Sin Datos)//-->
													  <%
				End if%>
													</span></td>
												</tr>
											  </table>
											  <span class="normalname">
											  <%
								else%>
											  <!--(Sin Datos)//-->
											  <%
							end if
							Rs4.MoveNext
						Wend
					else%>
											  <!--(Sin Datos)//-->
											  <%
				End If%>
											</span></td>
										</tr>
									</table>
									  <span class="normalname">
									  <%
							Rs2.MoveNext
						Wend
					else%>
									  <!--(Sin Datos)//-->
									  <%
				End if%>
									</span></td>
								</tr>
							</table></td>
						</tr>
						<tr>
						  <td width='5%' align='center' valign='top' bgcolor="#DBECF2" class="tex-totales-celda"><span class="normalname"><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs1("ANO") & "_II" %>');">II Semestre</a></span></td>
						  <td width='95%' valign='top'><span class="normalname">
							<%
				TablaRamos=""
				TablaProfesores=""
				TablaDocumentos=""
				%>
							</span>
							  <table width='100%' border='1' cellpadding='0' cellspacing='0' id='<%=Rs1("ANO") & "_II"%>' >
								<tr>
								  <td valign='top'><span class="normalname">
									<%
				Set Rs2 = Session("Conn").Execute("SELECT a.ANO, a.PERIODO, a.RAMOEQUIV as CODRAMO from ra_nota as a where a.codcli='" & Session ("CarreraEnCurso") & "' and a.ano='" & Rs1("ANO") & "' and a.periodo='2' order by a.ano, a.periodo")
				if Not Rs2.Bof or Not Rs2.Eof _
					then
						Rs2.MoveFirst
						While Not Rs2.Eof
							Set Rs3 = Session("Conn").Execute("SELECT NOMBRE from ra_ramo where codramo='" & Rs2("CODRAMO") & "'")
							
						if MaterialSeccion ="SI" then							
							Set Rs4 = Session("Conn").Execute("select c.CODPROFE codprof from ca_avisos c, ra_nota n,ra_horascont h where n.codcli = '"& session("codcli") &"' and c.CodRamo='" & Rs2("CODRAMO") & "' and c.Ano='"&trim(ano)&"' And c.Periodo='2' and c.codcarr in (select codcarr COLLATE Modern_Spanish_CI_AS from mt_Carrer where sede='"& session("codsede") &"') AND c.CODRAMO=n.RAMOEQUIV COLLATE SQL_Latin1_General_CP1_CI_AS AND c.ano=n.ANO AND c.periodo=n.PERIODO AND c.CODPROFE=h.CODPROF COLLATE SQL_Latin1_General_CP1_CI_AS AND c.CODRAMO =h.CODRAMO COLLATE SQL_Latin1_General_CP1_CI_AS AND n.codsecc = h.codsecc AND c.ano=h.ANO AND c.periodo=h.periodo AND h.CODSEDE='"& session("codsede") &"' GROUP BY c.CODPROFE")
						else
							Set Rs4 = Session("Conn").Execute("select DISTINCT codprofe codprof from ca_avisos where CodRamo='" & Rs2("CODRAMO") & "' and Ano='"&trim(ano)&"' And Periodo='2' and codcarr in (select codcarr COLLATE Modern_Spanish_CI_AS from mt_Carrer where sede='"& session("codsede") &"')") 
						end if  
				%>
									</span>
									  <table width='100%' border='1' cellpadding='3' cellspacing='0' class='smalltext'>
										<tr>
										  <td bgcolor='#FFFFCC'><span class="normalname"><a style='text-align:left; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs2("CODRAMO")%>');"><%if Not Rs4.Bof or Not Rs4.Eof then response.Write("<img src='addon/imagenes/IconoMensaje[1].gif' width='16' height='16'>") end if%>&nbsp;&nbsp;(<%=Rs2("CODRAMO")%>):<%=Rs3("NOMBRE")%></a>
										  </span></td> 
										</tr>
										<tr>
										  <td valign='top' style='display:none ;' id='<%=Rs2("CODRAMO")%>'>
											<%
				
				
				if Not Rs4.Bof or Not Rs4.Eof _
					then
						Rs4.MoveFirst
						While Not Rs4.Eof
							Set Rs6=Session("Conn").Execute("select * from ra_profes where codprof='" & Rs4("codprof") & "'") 
							if Not Rs6.Bof or Not Rs6.Eof _ 
								then%>
											  <table width='100%' border='1' cellpadding='0' cellspacing='0' class='smalltext'>
												<tr>
												  <td class='maintitle'><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs2("CODRAMO") & "_" & Rs6("CODPROF")%>');"><%=Rs6("NOMBRES") & " " & Rs6("AP_PATER") & " " & Rs6("AP_MATER")%></a></td>
												</tr>
												<tr>
												  <td valign='top' style='display:none ;' id='<%=Rs2("CODRAMO") & "_" & Rs6("CodProf")%>'>
													<%
				TablaDocumentos=""
				HeadDoctos=""
				BodyDoctos=""
				FooterDoctos=""
				Set Rs5=Session("Conn").Execute("select Fecha,Descripcion, IdDocumento from ca_avisos where codramo='" & Rs2("CODRAMO") & "' and ano='"&trim(ano)&"' and CodProfe='" & Rs6("CodProf") & "' ")
				'response.End() 
				if Not Rs5.Bof or Not Rs5.Eof _
					then
						Rs5.MoveFirst%>
						   <table width='100%' border='1' cellpadding='2' cellspacing='0'>
														<tr>
														  <td width='10%'  align='center' valign='middle'><span class="normalname">Fecha Publicaci&oacute;n</span></td>
														  <td  width='30%' ><span class="normalname">T&iacute;tulo</span></td>
														  <td  width='60%' ><span class="normalname">Descripci&oacute;n</span></td>
														</tr>
														<%
						While Not Rs5.Eof%>
														<tr>
														  <td align='center' valign='middle'><span class="normalname"><%=Day(cDate(Rs5("Fecha"))) & "/" & Month(cDate(Rs5("fecha"))) & "/" & year(cDate(Rs5("fecha")))%></span></td>
														  <td width='30%' align="left" valign="middle"><span class="normalname"><%=Rs5("Descripcion")%></span>&nbsp;</td>
														  <td width='60%' align="left" valign="top"><%=Rs5("IdDocumento")%>
														  </td>
														</tr>
														<%
							Rs5.MoveNext
						Wend%>
												    </table>
													  <span class="normalname">
													  <%
					else%>
													  <!--(Sin Datos)//-->
													  <%
				End if%>
													</span></td>
												</tr>
											  </table>
											  <span class="normalname">
											  <%
								else%>
											  <!--(Sin Datos)//-->
											  <%
							end if
							Rs4.MoveNext
						Wend
					else%>
											  <!--(Sin Datos)//-->
											  <%
				End If%>
											</span></td>
										</tr>
									  </table>
									  <span class="normalname">
									  <%
							Rs2.MoveNext
						Wend
				End if
				%>
									</span></td>
								</tr>
							</table></td>
						</tr>
					</table></td>
		  
<%
	
		Rs1.MoveNext
	wend
end if
%>
</tr>
				</table>
			  
<table>
           <tr>
           <br><br><br>
    <td width="83" align="left"><A onMouseOver="MM_swapImage('Image7','','Imagenes/botones/volver-on.gif',1)" onmouseout=MM_swapImgRestore() href="javascript: window.top.location.href ='menu_material.asp'"><IMG src="Imagenes/botones/volver-of.gif" border="0"  name="Image7"></A></td>
  </tr>
  </table>
			  </tr>
			</table>
			</td>
		  </tr>
		</table>

</body>
	</td>
  </tr>
</table>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
