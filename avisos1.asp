<!--#include file="addon/includes/funciones.inc" -->
<!--#include file="addon/includes/libreria.asp" -->
<link rel="stylesheet" href="css/tex-normales.css" type="text/css">
<style type="text/css">
<!--
.tex-normales {  font: normal 9px Verdana, Arial, Helvetica, sans-serif; color: #000000; letter-spacing: -0.5px}
table{font-family:Arial, Helvetica, sans-serif;font-size:11px}
-->
</style>
<script src="addon/includes/enac.js"></script>

<%
//dim TablaRamos
//dim TablaProfesores
//dim TablaDocumentos
//dim HeadDoctos
//dim BodyDoctos
//dim FooterDoctos
//dim Bloque
Set RsAyP = Session("Conn").Execute("SELECT DISTINCT ano,periodo FROM mt_parame ")
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

Set Rs1 = Session("Conn").Execute("SELECT DISTINCT ano FROM ra_nota WHERE codcli='" & Session ("CarreraEnCurso") & "'  and ano='"&trim(ano)&"'  and periodo='"&trim(periodo)&"' ORDER BY ano")

Set Rs2 = Session("Conn").Execute("SELECT a.ANO, a.PERIODO, a.RAMOEQUIV as CODRAMO from ra_nota as a where a.codcli='" & Session ("CarreraEnCurso") & "' and a.CodRamo<>'' and a.ano='" & Rs1("ANO") & "' and a.periodo='1' order by a.ano, a.periodo")
if Not Rs1.Bof or Not Rs1.Eof then
	Rs1.MoveFirst
	While Not Rs1.Eof%>

<body onLoad="expandit('<%=Rs1("ANO")%>');expandit('<%=Rs1("ANO") & "_I" %>');">    
<table width="100%" border='1' cellpadding='5' cellspacing='0' > 
  <tr>
    <td class="caption"><span class="normalname"><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs1("ANO")%>');"><b>Avisos Publicados A&ntilde;o <%=Rs1("ANO")%></b></a></span></td>
  </tr>
  <tr>
    <td valign='top' style='display:none ;' id='<%=Rs1("ANO")%>'><table width="100%" border='1' cellpadding='0' cellspacing='0'>
        <tr>
          <td width='5%' align='center' valign='top'><span class="normalname"><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs1("ANO") & "_I" %>');">I Semestre</a></span></td>
          <td width='95%' valign='top'><span class="normalname">
            <%					
TablaRamos=""
TablaProfesores=""
TablaDocumentos=""
%>
            </span>
              <table width='100%' border='1' cellpadding='0' cellspacing='0' id='<%=Rs1("ANO") & "_I" %>' style='display:none ;'>
                <tr>
                  <td valign='top'>
                    <%						
//Set Rs2 = Session("Conn").Execute("SELECT a.ANO, a.PERIODO, a.RAMOEQUIV as CODRAMO from ra_nota as a where a.codcli='" & Session ("CarreraEnCurso") & "' and a.CodRamo<>'' and a.ano='" & Rs1("ANO") & "' and a.periodo='1' order by a.ano, a.periodo")
if Not Rs2.Bof or Not Rs2.Eof _
	then
		Rs2.MoveFirst
		While Not Rs2.Eof
			Set Rs3 = Session("Conn").Execute("SELECT NOMBRE from ra_ramo where codramo='" & Rs2("CODRAMO") & "'")
%>
                      <table width='100%' border='1' cellpadding='3' cellspacing='0' class='smalltext'>
                        <tr>
                          <td align="left" valign="middle" bgcolor='#FFFFCC'><span class="normalname"><a style='text-align:left; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs2("ANO") & "_" & Rs2("PERIODO") & "_" & Rs2("CODRAMO")%>');">
                          <%if ConterProfesRamos(Rs2("CODRAMO"), 1)>0 then response.Write("<img src='addon/imagenes/IconoMensaje[1].gif' width='16' height='16'>") end if%>&nbsp;&nbsp;<%=Rs2("CODRAMO")%>):<%=Rs3("NOMBRE")%></a></span></td>
                        </tr>
                        <tr>
                          <td valign='top' style='display:none ;' id='<%=Rs2("ANO") & "_" & Rs2("PERIODO") & "_" & Rs2("CODRAMO")%>'>
                            <%
TablaProfesores=""
Set Rs4 = Session("Conn").Execute("select DISTINCT codprof from ca_doctos where codramo='" & Rs2("CODRAMO") & "' And PeriodoPublicacion='1' And Activo='1'")
'response.Write(Rs4)
'response.End()
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
Set Rs5=Session("Conn").Execute("select Fecha,Descripcion, IdDocumento from ca_avisos where CodRamo='" & Rs2("CODRAMO") & "' and CodProfe='" & Rs6("CodProf") & "' ")
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
          <td width='5%' align='center' valign='top'><span class="normalname"><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs1("ANO") & "_II" %>');">II Semestre</a></span></td>
          <td width='95%' valign='top'><span class="normalname">
            <%
TablaRamos=""
TablaProfesores=""
TablaDocumentos=""
%>
            </span>
              <table width='100%' border='1' cellpadding='0' cellspacing='0' id='<%=Rs1("ANO") & "_II"%>' style='display:none ;'>
                <tr>
                  <td valign='top'><span class="normalname">
                    <%
Set Rs2 = Session("Conn").Execute("SELECT a.ANO, a.PERIODO, a.RAMOEQUIV as CODRAMO from ra_nota as a where a.codcli='" & Session ("CarreraEnCurso") & "' and a.ano='" & Rs1("ANO") & "' and a.periodo='2' order by a.ano, a.periodo")
if Not Rs2.Bof or Not Rs2.Eof _
	then
		Rs2.MoveFirst
		While Not Rs2.Eof
			Set Rs3 = Session("Conn").Execute("SELECT NOMBRE from ra_ramo where codramo='" & Rs2("CODRAMO") & "'")%>
                    </span>
                      <table width='100%' border='1' cellpadding='3' cellspacing='0' class='smalltext'>
                        <tr>
                          <td bgcolor='#FFFFCC'><span class="normalname"><a style='text-align:left; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs2("CODRAMO")%>');"><%if ConterProfesRamos(Rs2("CODRAMO"), 2)>0 then response.Write("<img src='addon/imagenes/IconoMensaje[1].gif' width='16' height='16'>") end if%>&nbsp;&nbsp;(<%=Rs2("CODRAMO")%>):<%=Rs3("NOMBRE")%></a>
                          </span></td>
                        </tr>
                        <tr>
                          <td valign='top' style='display:none ;' id='<%=Rs2("CODRAMO")%>'>
                            <%
Set Rs4 = Session("Conn").Execute("select DISTINCT codprof from ca_doctos where CodRamo='" & Rs2("CODRAMO") & "' And PeriodoPublicacion='2' And Activo='1'")
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
Set Rs5=Session("Conn").Execute("select Fecha,Descripcion, IdDocumento from ca_avisos where codramo='" & Rs2("CODRAMO") & "' and CodProfe='" & Rs6("CodProf") & "' ")
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
  </tr>
</table>
<%
		Rs1.MoveNext
	wend
end if
%>
