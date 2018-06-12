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
RutaDoctos = Server.MapPath("Doctos")& "\"
'response.Write(RutaDoctos) 

Set RsAyP = session("conn").Execute("SELECT dbo.Fn_ValorParame('ano') as ano,dbo.Fn_ValorParame('periodo') as periodo ")
if Not RsAyP.Bof or Not RsAyP.Eof then
	RsAyP.MoveFirst
	While Not RsAyP.Eof
	  ano=RsAyP(0)
	  periodo=RsAyP(1)
	  RsAyP.MoveNext  
	wend
End if

if request.QueryString("IDCDAumas")<>"" _
	then
		Session ("CarreraEnCurso")=Session("RutAlum")
end if

Session ("CarreraEnCurso") = Session ("codcli")
'Set Rs1 = session("conn").Execute("SELECT DISTINCT ano FROM ra_nota WHERE codcli='" & Session ("CarreraEnCurso") & "'  and ano='"&trim(ano)&"'  and periodo='"&trim(periodo)&"' ORDER BY ano")

'if  Rs1.bof or Rs1.eof  then 
	'response.Redirect("MensajeSinDatos.asp")
	'response.end()
'end if 
' se cambia la validación por secciones abiertas o cerradas

Set Rs1 = session("conn").Execute("SELECT distinct s.ano FROM dbo.RA_SECCIO s, ra_nota n WHERE n.CODCLI = '" & Session ("CarreraEnCurso") & "' AND n.RAMOEQUIV = s.CODRAMO AND n.codsecc = s.CODSECC AND n.CODSEDE = s.CODSEDE AND n.ano = s.ANO AND n.PERIODO = s.PERIODO AND CERRADA = 'N' order by s.ano")

'Set Rs1 = session("conn").Execute("SELECT 2015 ano union select 2016 ano")

if  Rs1.bof or Rs1.eof  then 
	response.Redirect("MensajeSinDatos.asp")
	response.end()
end if 

strParame="SELECT dbo.Fn_ValorParame('MATERIAL_SECCION')Parame"
if BCL_ADO(strParame, rstParame) then
	MaterialSeccion=trim(valnulo(rstParame("Parame"),STR_))
else
	MaterialSeccion=""
end if

Set Rs2 = session("conn").Execute("SELECT a.ANO, a.PERIODO, a.RAMOEQUIV as CODRAMO from ra_nota as a where a.codcli='" & Session ("CarreraEnCurso") & "' and a.RAMOEQUIV<>'' and a.ano='" & Rs1("ANO") & "' and a.periodo='1' order by a.ano, a.periodo")


	
StrEnac="SELECT dbo.Fn_ValorParame('DOCTOSENACPA')DOCTOSENACPA,dbo.Fn_ValorParame('DOCTOSMPWPA')DOCTOSMPWPA"
if BCL_ADO(StrEnac, rstEnac) then
	DOCTOSENACPA=trim(valnulo(rstEnac("DOCTOSENACPA"),STR_))
	DOCTOSMPWPA = trim(valnulo(rstEnac("DOCTOSMPWPA"),STR_))
else
	DOCTOSENACPA=""
	DOCTOSMPWPA=""
end if	
	%>
	
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%=Session("NombrePestana")%></title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1"> 
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<link rel="stylesheet" href="css/tex-normales.css" type="text/css">
<style type="text/css">
<!--
.tex-normales {  font: normal 9px Verdana, Arial, Helvetica, sans-serif; color: #000000; letter-spacing: -0.5px}
table{font-family:Arial, Helvetica, sans-serif;font-size:11px}
-->
</style>
<script src="addon/includes/enac.js"></script>
<script language="javascript" type="text/javascript" src="include/prototype.js"></script>
<script language="javascript">
function ActualizaVisto(codramo,codprof,codsecc) { 

	var url = 'InsertaDoctoLeido.asp'; 
	var pars = 'CR='+codramo+'&CP='+codprof+'&CS='+codsecc;	
	var wafoAjax = new Ajax.Request(url, {method: 'post', 
										  parameters: pars, 
										  onComplete: function(originalRequest){
														var _objData = eval(originalRequest.responseText);		
													}
						});
						
						
}

</script>

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
			<p>&nbsp;</p>
	<%
	if Not Rs1.Bof or Not Rs1.Eof then
	Rs1.MoveFirst
	While Not Rs1.Eof
	%>
		<table width="770" border='1' cellpadding='5' cellspacing='0' onKeyPress="" align="center" > 
  <tr>
    <td class="caption" background="Imagenes/fdo-cabecera-cel22.jpg" height="30"><span class="normalname" style="font-size: 14"><a class="text-cabecera-celda" style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;'>Documentos Publicados A&ntilde;o <%=Rs1("ANO")%></a></span></td>
  </tr>
  <tr>
    <td valign='top' style='display:none ;' id='<%=Rs1("ANO")%>'><table width="100%" border='1' cellpadding='0' cellspacing='0'>
        <tr> 
          <td width='5%' align='center' valign='top' bgcolor="#DBECF2" class="tex-totales-celda"><span class="normalname"><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs1("ANO") & "_I" %>');">I Semestre</a></span></td>
          <td width='95%' valign='top' class="tex-totales-celda"><span class="tex-totales-celda">
            <%					
TablaRamos=""
TablaProfesores=""
TablaDocumentos=""
%>
            </span>
              <table width='100%' border='1' bordercolor="#E7E4BC" cellpadding='0' cellspacing='0' id='<%=Rs1("ANO") & "_I" %>' style='display:none ;'>
                <tr>
                  <td valign='top'>
                    <%						
Set Rs2 = session("conn").Execute("SELECT a.ANO, a.PERIODO, a.RAMOEQUIV as CODRAMO,a.codsecc from ra_nota as a where a.codcli='" & Session ("CarreraEnCurso") & "' and a.RAMOEQUIV <>'' and a.ano='" & Rs1("ANO") & "' and a.periodo='1' order by a.ano, a.periodo")

if Not Rs2.Bof or Not Rs2.Eof _
	then
		Rs2.MoveFirst
		While Not Rs2.Eof
			Set Rs3 = session("conn").Execute("SELECT NOMBRE from ra_ramo where codramo='" & Rs2("CODRAMO") & "'")
%>
                      <table width='100%' border='1' cellpadding='3' cellspacing='0' class='smalltext'>
                        <tr>
                          <td align="left" valign="middle" bgcolor='#FFFFCC'><span class="normalname"><a style='text-align:left; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs2("ANO") & "_" & Rs2("PERIODO") & "_" & Rs2("CODRAMO")%>');">
                          <%
						  
						  if DOCTOSENACPA="SI" then
								  strD0="select DISTINCT CodProf from ca_doctos where codramo='" & Rs2("CODRAMO") & "' and AnoPublicacion= '"&trim(ano)&"' And PeriodoPublicacion='1' And Activo='1' and codcarr in (select codcarr COLLATE Modern_Spanish_CI_AS from mt_Carrer where sede='"& session("codsede") &"')"
									
								  Set Rs4 = session("conn").Execute(strD0) 
							else	   
								  strD0="select DISTINCT CodProf from ca_doctos where codramo='" & Rs2("CODRAMO") & "' and AnoPublicacion= '"&trim(ano)&"' And PeriodoPublicacion='1' And Activo='1' and codcarr in (select codcarr COLLATE Modern_Spanish_CI_AS from mt_Carrer where sede='"& session("codsede") &"')"
									
								  if MaterialSeccion="SI" then
									strD0=strD0 & "and codsecc='"& Rs2("codsecc") &"'"
								  end if 
								  Set Rs4 = session("conn").Execute(strD0) 
							end if
							
						  	
							if not Rs4.eof then
						  		response.Write("<img src='addon/imagenes/btn_show-all_bg.gif' width='16' height='16'>")						  
							end if
							
						  %>&nbsp;&nbsp;<%=Rs2("CODRAMO")%>:<%=Rs3("NOMBRE")%></a></span></td>
                        </tr>
                        <tr >
                          <td valign='top'  style='display:none ;' id='<%=Rs2("ANO") & "_" & Rs2("PERIODO") & "_" & Rs2("CODRAMO")%>'>
                            <%

TablaProfesores=""

if Not Rs4.Bof or Not Rs4.Eof _
	then
		Rs4.MoveFirst
		While Not Rs4.Eof
			Set Rs6=session("conn").Execute("select * from ra_profes where codprof='" & Rs4("CodProf") & "'")
			if Not Rs6.Bof or Not Rs6.Eof _ 
				then%>
                              <table width='100%' border='1' cellpadding='0' cellspacing='0' class='smalltext'>
                                <tr>
                                  <td class='maintitle' ><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs2("CODRAMO") & "_" & Rs6("CODPROF") & "_" & Rs2("PERIODO")%>');"><%=Rs6("NOMBRES") & " " & Rs6("AP_PATER") & " " & Rs6("AP_MATER")%></a></td>
                                </tr>
                                <tr>
                                  <td valign='top' style='display:none ;' id='<%=Rs2("CODRAMO") & "_" & Rs6("CodProf")& "_" & Rs2("PERIODO")%>'>
                                    <%
TablaDocumentos=""
HeadDoctos=""
BodyDoctos=""
FooterDoctos=""

if DOCTOSENACPA="SI" then 
	strD="select  IDCONTER,CODPROF,CODCARR,CODRAMO,TIPODOCUMENTO,REFERENCIA,DETALLEDOCUMENTO,IDARCHIVO1,IDARCHIVO2,IDARCHIVO3,CONVERT(VARCHAR(10),CONVERT(DATE,FechaUpLoad),105)FECHAUPLOAD,CONVERT(DATE,FECHAUPLOAD)x,ANOPUBLICACION,PERIODOPUBLICACION,FECHAACTUALIZACION,ACTIVO,CODPROFEFOLDER,"
	if MaterialSeccion="SI" then
		strD = strD & "CODSECC,"
	end if 
	strD = strD & "EVALUADO,FECHA_INI,FECHA_FIN,ACEPTA_MAT_FUERA_PLAZO from ca_doctos where codramo='" & Rs2("CODRAMO") & "' and CodProf='" & Rs6("CodProf") & "' and AnoPublicacion='"&trim(ano)&"' "
	strD=strD & "And PeriodoPublicacion='1' And Activo='1' and codcarr in (select codcarr COLLATE Modern_Spanish_CI_AS from mt_Carrer where sede='"& session("codsede") &"')"
	if MaterialSeccion="SI" then
		strD=strD & "and codsecc='"& Rs2("codsecc") &"'"
	end if
	
	if MaterialSeccion="SI" then
		Set Rs5 = session("conn").Execute("SELECT distinct IDCONTER,doc.CODPROF,doc.CODCARR,doc.CODRAMO,Doc.TIPODOCUMENTO,REFERENCIA,DETALLEDOCUMENTO,IDARCHIVO1,IDARCHIVO2,IDARCHIVO3,CONVERT(VARCHAR(10),CONVERT(DATE,FechaUpLoad),105)FECHAUPLOAD,CONVERT(DATE,FECHAUPLOAD)x,ANOPUBLICACION,PERIODOPUBLICACION,FECHAACTUALIZACION,ACTIVO,CODPROFEFOLDER,sec.CODSECC,EVALUADO,FECHA_INI,FECHA_FIN,ACEPTA_MAT_FUERA_PLAZO FROM RA_NOTA NOTA, RA_RAMO RAMO,RA_SECCIO SEC , RA_PROFES PRO, RA_HORPROF HOR, MT_ALUMNO ALUM , CA_Doctos DOC WHERE NOTA.RAMOEQUIV = RAMO.CODRAMO AND NOTA.CODCLI = '" & Session ("CarreraEnCurso") & "' AND NOTA.CODCLI = ALUM.CODCLI AND NOTA.ANO = '" & Rs1("ANO") & "' AND HOR.CODPROF = '" & Rs4("CodProf") & "' AND NOTA.PERIODO = 1 AND NOTA.CODSECC <> 0 AND SEC.ANO = NOTA.ANO AND SEC.PERIODO = NOTA.PERIODO AND SEC.CODSECC = NOTA.CODSECC AND SEC.CODRAMO = NOTA.RAMOEQUIV AND SEC.CODSEDE = ALUM.CODSEDE AND SEC.CODRAMO = HOR.CODRAMO  AND SEC.ANO = HOR.ANO AND SEC.PERIODO = HOR.PERIODO AND SEC.CODSEDE = HOR.CODSEDE AND HOR.CODPROF = PRO.CODPROF and hor.CODPROF = doc.codprof collate Modern_Spanish_CI_AS and sec.CODRAMO = doc.codramo collate Modern_Spanish_CI_AS and sec.CODCARR = doc.codcarr collate Modern_Spanish_CI_AS  and convert (varchar(4),sec.ANO) =  doc.AnoPublicacion and convert (varchar(1),sec.PERIODO) = doc.PeriodoPublicacion and doc.codramo='" & Rs2("CODRAMO") & "' union "& strD & " order by x desc ")
	else
	    Set Rs5 = session("conn").Execute("SELECT distinct IDCONTER,doc.CODPROF,doc.CODCARR,doc.CODRAMO,doc.TIPODOCUMENTO,REFERENCIA,DETALLEDOCUMENTO,IDARCHIVO1,IDARCHIVO2,IDARCHIVO3,CONVERT(VARCHAR(10),CONVERT(DATE,FechaUpLoad),105)FECHAUPLOAD,CONVERT(DATE,FECHAUPLOAD)x,ANOPUBLICACION,PERIODOPUBLICACION,FECHAACTUALIZACION,ACTIVO,CODPROFEFOLDER,EVALUADO,FECHA_INI,FECHA_FIN,ACEPTA_MAT_FUERA_PLAZO FROM RA_NOTA NOTA, RA_RAMO RAMO,RA_SECCIO SEC , RA_PROFES PRO, RA_HORPROF HOR, MT_ALUMNO ALUM , CA_Doctos DOC WHERE NOTA.RAMOEQUIV = RAMO.CODRAMO AND NOTA.CODCLI = '" & Session ("CarreraEnCurso") & "' AND NOTA.CODCLI = ALUM.CODCLI AND NOTA.ANO = '" & Rs1("ANO") & "' AND HOR.CODPROF = '" & Rs4("CodProf") & "' AND NOTA.PERIODO = 1 AND NOTA.CODSECC <> 0 AND SEC.ANO = NOTA.ANO AND SEC.PERIODO = NOTA.PERIODO AND SEC.CODSECC = NOTA.CODSECC AND SEC.CODRAMO = NOTA.RAMOEQUIV AND SEC.CODSEDE = ALUM.CODSEDE AND SEC.CODRAMO = HOR.CODRAMO  AND SEC.ANO = HOR.ANO AND SEC.PERIODO = HOR.PERIODO AND SEC.CODSEDE = HOR.CODSEDE AND HOR.CODPROF = PRO.CODPROF and hor.CODPROF = doc.codprof collate Modern_Spanish_CI_AS and sec.CODRAMO = doc.codramo collate Modern_Spanish_CI_AS and sec.CODCARR = doc.codcarr collate Modern_Spanish_CI_AS  and convert (varchar(4),sec.ANO) =  doc.AnoPublicacion and convert (varchar(1),sec.PERIODO) = doc.PeriodoPublicacion and doc.codramo='" & Rs2("CODRAMO") & "'  union "& strD & " order by x desc ")
	end if

	
else
	strD="select  IDCONTER,CODPROF,CODCARR,CODRAMO,TIPODOCUMENTO,REFERENCIA,DETALLEDOCUMENTO,IDARCHIVO1,IDARCHIVO2,IDARCHIVO3,CONVERT(VARCHAR(10),CONVERT(DATE,FechaUpLoad),105)FECHAUPLOAD,CONVERT(DATE,FECHAUPLOAD)x,ANOPUBLICACION,PERIODOPUBLICACION,FECHAACTUALIZACION,ACTIVO,CODPROFEFOLDER,"
	if MaterialSeccion="SI" then
		strD = strD & "CODSECC,"
	end if 
	strD = strD & "EVALUADO,FECHA_INI,FECHA_FIN,ACEPTA_MAT_FUERA_PLAZO from ca_doctos where codramo='" & Rs2("CODRAMO") & "' and CodProf='" & Rs6("CodProf") & "' and AnoPublicacion='"&trim(ano)&"' "
	
	strD=strD & "And PeriodoPublicacion='1' And Activo='1' and codcarr in (select codcarr COLLATE Modern_Spanish_CI_AS from mt_Carrer where sede='"& session("codsede") &"')"
	if MaterialSeccion="SI" then
		strD=strD & "and codsecc='"& Rs2("codsecc") &"'"
	end if
	strD = strD & " order by x desc"
	
	Set Rs5=session("conn").Execute(strD)

end if




if Not Rs5.Bof or Not Rs5.Eof _
	then
		Rs5.MoveFirst%>
                                      <table width='100%' border='1' cellpadding='2' cellspacing='0'>
                                        <tr>
                                          <td width='7%' align='center' valign='middle'><span class="normalname">Fecha Publicaci&oacute;n</span></td>
                                          <td colspan='4'><span class="normalname">Detalle de Documentaci&oacute;n / Apuntes </span></td>
                                        </tr>
                                        <%
		While Not Rs5.Eof%>
                              <tr>
                         <%if not isnull(Rs5("tipodocumento")) then%>
                         		<%if DOCTOSMPWPA="SI" then%>
                         		    <% if Rs5("IdArchivo1")<>"" then%>
                                    <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\{"& Rs5("codprofefolder") &"}\"& Rs5("anopublicacion") &"\S1\"&Rs5("IdArchivo1")) = true then%>
                                    <td align='center' valign='middle'><span class="normalname"><%=Day(cDate(Rs5("FechaUpLoad"))) & "/" & Month(cDate(Rs5("FechaUpLoad"))) & "/" & year(cDate(Rs5("FechaUpLoad")))%></span></td>
                                    <td width='89%' align="left" valign="middle"><span class="normalname"><%=Rs5("REFERENCIA")%></span>&nbsp;</td>
                                    <td width='4%' align="center" valign="top" onClick="ActualizaVisto('<%=Rs5("CODRAMO")%>','<%=Rs5("CODPROF")%>','<%=Rs5("CODSECC")%>');">
                                    <a href='Doctos/<%=Rs5("codramo")%>/{<%=Rs5("codprofefolder")%>}/<%=Rs5("anopublicacion")%>/S1/<%=Rs5("IdArchivo1")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                    </td>
                                    <%end if%>
                                    <%end if%>
                                     
                                    <% if Rs5("IdArchivo2")<>"" then%>
                                    <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\{"& Rs5("codprofefolder") &"}\"& Rs5("anopublicacion") &"\S1\"&Rs5("IdArchivo2")) = true then%>                                  
                                    <td width='4%' align="center" valign="top" onClick="ActualizaVisto('<%=Rs5("CODRAMO")%>','<%=Rs5("CODPROF")%>','<%=Rs5("CODSECC")%>');">
                                    <a href='Doctos/<%=Rs5("codramo")%>/{<%=Rs5("codprofefolder")%>}/<%=Rs5("anopublicacion")%>/S1/<%=Rs5("IdArchivo2")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                    </td>
                                    <%end if%>
                                    <%end if%>                                  
                                      
                                    <% if Rs5("IdArchivo3")<>"" then%>
                                    <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\{"& Rs5("codprofefolder") &"}\"& Rs5("anopublicacion") &"\S1\"&Rs5("IdArchivo3")) = true then%>                                
                                    <td width='4%' align="center" valign="top" onClick="ActualizaVisto('<%=Rs5("CODRAMO")%>','<%=Rs5("CODPROF")%>','<%=Rs5("CODSECC")%>');">
                                    <a href='Doctos/<%=Rs5("codramo")%>/{<%=Rs5("codprofefolder")%>}/<%=Rs5("anopublicacion")%>/S1/<%=Rs5("IdArchivo3")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                    </td>
                                    <%end if%>
                                    <%end if%>
                         		<%elseif DOCTOSENACPA="SI" then%>
                                
                                    <td align='center' valign='middle'><span class="normalname"><%=Day(cDate(Rs5("FechaUpLoad"))) & "/" & Month(cDate(Rs5("FechaUpLoad"))) & "/" & year(cDate(Rs5("FechaUpLoad")))%></span></td>
                                      <td width='89%' align="left" valign="middle"><span class="normalname"><%=Rs5("REFERENCIA")%></span>&nbsp;</td>
                                      <td width='4%' align="center" valign="top">
                                      
									  <% if Rs5("IdArchivo1")<>"" then%>
                                      <% if ExisteArchivo(RutaDoctos & Rs5("CODCARR") &"\"& Rs5("CODPROF") &"\"& Rs5("CODRAMO") &"\"& Rs5("idconter")&"\"& Rs5("IdArchivo1")) = true then%>
                                      <a href='doctos/<%=Rs5("CODCARR")%>/<%=Rs5("CODPROF")%>/<%=Rs5("CODRAMO")%>/<%=Rs5("idconter")%>/<%=Rs5("IdArchivo1")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                      <%end if%>
                                      <%end if%>
                                      
									  <% if Rs5("IdArchivo2")<>"" then%>
                                      <% if ExisteArchivo(RutaDoctos & Rs5("CODCARR") &"\"& Rs5("CODPROF") &"\"& Rs5("CODRAMO") &"\"& Rs5("idconter")&"\"& Rs5("IdArchivo2")) = true then%>
                                      <a href='doctos/<%=Rs5("CODCARR")%>/<%=Rs5("CODPROF")%>/<%=Rs5("CODRAMO")%>/<%=Rs5("idconter")%>/<%=Rs5("IdArchivo2")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                      <%end if%>
                                      <%end if%>
                                      
                                      <% if Rs5("IdArchivo3")<>"" then%>
                                      <% if ExisteArchivo(RutaDoctos & Rs5("CODCARR") &"\"& Rs5("CODPROF") &"\"& Rs5("CODRAMO") &"\"& Rs5("idconter")&"\"& Rs5("IdArchivo3")) = true then%>
                                      <a href='doctos/<%=Rs5("CODCARR")%>/<%=Rs5("CODPROF")%>/<%=Rs5("CODRAMO")%>/<%=Rs5("idconter")%>/<%=Rs5("IdArchivo3")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                      <%end if%>
                                      <%end if%>		                                      								  											  
                                      </td>                                
                         
                         		<%else%>
									  <% if Rs5("IdArchivo1")<>"" then%>
                                      <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\"& Rs5("codprofefolder") &"\"& Rs5("anopublicacion") &"\S1\"&Rs5("IdArchivo1")) = true then%>
                                      <td align='center' valign='middle'><span class="normalname"><%=Day(cDate(Rs5("FechaUpLoad"))) & "/" & Month(cDate(Rs5("FechaUpLoad"))) & "/" & year(cDate(Rs5("FechaUpLoad")))%></span></td>
                                      <td width='89%' align="left" valign="middle"><span class="normalname"><%=Rs5("REFERENCIA")%></span>&nbsp;</td>
                                      <td width='4%' align="center" valign="top" onClick="ActualizaVisto('<%=Rs5("CODRAMO")%>','<%=Rs5("CODPROF")%>','<%=Rs5("CODSECC")%>');">
                                      <a href='Doctos/<%=Rs5("codramo")%>/{<%=Rs5("codprofefolder")%>}/<%=Rs5("anopublicacion")%>/S1/<%=Rs5("IdArchivo1")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                      </td>
                                      <%end if%>
                                      <%end if%>
                                      
                                      <% if Rs5("IdArchivo2")<>"" then%>
                                      <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\"& Rs5("codprofefolder") &"\"& Rs5("anopublicacion") &"\S1\"&Rs5("IdArchivo2")) = true then%>                                  
                                      <td width='4%' align="center" valign="top" onClick="ActualizaVisto('<%=Rs5("CODRAMO")%>','<%=Rs5("CODPROF")%>','<%=Rs5("CODSECC")%>');">
                                      <a href='Doctos/<%=Rs5("codramo")%>/{<%=Rs5("codprofefolder")%>}/<%=Rs5("anopublicacion")%>/S1/<%=Rs5("IdArchivo2")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                      </td>
                                      <%end if%>
                                      <%end if%>                                  
                                      
                                      <% if Rs5("IdArchivo3")<>"" then%>
                                      <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\"& Rs5("codprofefolder") &"\"& Rs5("anopublicacion") &"\S1\"&Rs5("IdArchivo3")) = true then%>                                  
                                      <td width='4%' align="center" valign="top" onClick="ActualizaVisto('<%=Rs5("CODRAMO")%>','<%=Rs5("CODPROF")%>','<%=Rs5("CODSECC")%>');">
                                      <a href='Doctos/<%=Rs5("codramo")%>/{<%=Rs5("codprofefolder")%>}/<%=Rs5("anopublicacion")%>/S1/<%=Rs5("IdArchivo3")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                      </td>
                                      <%end if%>
                                      <%end if%>
                                <%end if%>               
                         <%else%> 
                              <%if MaterialSeccion="SI" then%> 
                                  <% if Rs5("IdArchivo1")<>"" then%>
                                  <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\"& Rs5("CodProf") &"\"& Rs5("CodProfefolder") &"\"& Rs5("fechaupload") &"\"& Rs5("periodopublicacion") &"\"& trim(Rs5("codsecc")) &"\"& Rs5("IdArchivo1")) = true then%>
                                 <td align='center' valign='middle'><span class="normalname"><%=Day(cDate(Rs5("FechaUpLoad"))) & "/" & Month(cDate(Rs5("FechaUpLoad"))) & "/" & year(cDate(Rs5("FechaUpLoad")))%></span></td>
                              	  <td width='89%' align="left" valign="middle"><span class="normalname"><%=Rs5("REFERENCIA")%></span>&nbsp;</td>
                                  <td width='4%' align="center" valign="top" onClick="ActualizaVisto('<%=Rs5("CODRAMO")%>','<%=Rs5("CODPROF")%>','<%=Rs5("CODSECC")%>');">
                                  <a href='Doctos/<%=Rs5("codramo")%>/<%=Rs5("CodProf")%>/<%=Rs5("CodProfefolder")%>/<%=Rs5("fechaupload")%>/<%=Rs5("periodopublicacion")%>/<%=trim(Rs5("codsecc"))%>/<%=Rs5("IdArchivo1")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                  </td>
                                  <%end if%>
                                  <%end if%>
                                  
                                  <% if Rs5("IdArchivo2")<>"" then%>
                                  <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\"& Rs5("CodProf") &"\"& Rs5("CodProfefolder") &"\"& Rs5("fechaupload") &"\"& Rs5("periodopublicacion") &"\"& trim(Rs5("codsecc")) &"\"& Rs5("IdArchivo2")) = true then%>                                  
                                  <td width='4%' align="center" valign="top" onClick="ActualizaVisto('<%=Rs5("CODRAMO")%>','<%=Rs5("CODPROF")%>','<%=Rs5("CODSECC")%>');">
                                  <a href='Doctos/<%=Rs5("codramo")%>/<%=Rs5("CodProf")%>/<%=Rs5("CodProfefolder")%>/<%=Rs5("fechaupload")%>/<%=Rs5("periodopublicacion")%>/<%=trim(Rs5("codsecc"))%>/<%=Rs5("IdArchivo2")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                  </td>
                                  <%end if%>
                                  <%end if%>
                                  
                                  <% if Rs5("IdArchivo3")<>"" then%>
                                  <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\"& Rs5("CodProf") &"\"& Rs5("CodProfefolder") &"\"& Rs5("fechaupload") &"\"& Rs5("periodopublicacion") &"\"& trim(Rs5("codsecc")) &"\"& Rs5("IdArchivo3")) = true then%>                                 
                                  <td width='4%' align="center" valign="top" onClick="ActualizaVisto('<%=Rs5("CODRAMO")%>','<%=Rs5("CODPROF")%>','<%=Rs5("CODSECC")%>');">
                                  <a href='Doctos/<%=Rs5("codramo")%>/<%=Rs5("CodProf")%>/<%=Rs5("CodProfefolder")%>/<%=Rs5("fechaupload")%>/<%=Rs5("periodopublicacion")%>/<%=trim(Rs5("codsecc"))%>/<%=Rs5("IdArchivo3")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                  </td>
                                  <%end if%>
                                  <%end if%>
                              
                              <%else%>
                                  <% if Rs5("IdArchivo1")<>"" then%>
                                  <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\"& Rs5("CodProf") &"\"& Rs5("CodProfefolder") &"\"& Rs5("fechaupload") &"\"& Rs5("periodopublicacion") &"\"& Rs5("IdArchivo1")) = true then%>
                                  <td align='center' valign='middle'><span class="normalname"><%=Day(cDate(Rs5("FechaUpLoad"))) & "/" & Month(cDate(Rs5("FechaUpLoad"))) & "/" & year(cDate(Rs5("FechaUpLoad")))%></span></td>
                              	  <td width='89%' align="left" valign="middle"><span class="normalname"><%=Rs5("REFERENCIA")%></span>&nbsp;</td>
                                  <td width='4%' align="center" valign="top" onClick="ActualizaVisto('<%=Rs5("CODRAMO")%>','<%=Rs5("CODPROF")%>','0');">
                                  <a href='Doctos/<%=Rs5("codramo")%>/<%=Rs5("CodProf")%>/<%=Rs5("CodProfefolder")%>/<%=Rs5("fechaupload")%>/<%=Rs5("periodopublicacion")%>/<%=Rs5("IdArchivo1")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                  </td>
                                  <%end if%>
                                  <%end if%>
                                  
                                  <% if Rs5("IdArchivo2")<>"" then%>
                                  <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\"& Rs5("CodProf") &"\"& Rs5("CodProfefolder") &"\"& Rs5("fechaupload") &"\"& Rs5("periodopublicacion") &"\"& Rs5("IdArchivo2")) = true then%>                                 
                                  <td width='4%' align="center" valign="top" onClick="ActualizaVisto('<%=Rs5("CODRAMO")%>','<%=Rs5("CODPROF")%>','0');">
                                  <a href='Doctos/<%=Rs5("codramo")%>/<%=Rs5("CodProf")%>/<%=Rs5("CodProfefolder")%>/<%=Rs5("fechaupload")%>/<%=Rs5("periodopublicacion")%>/<%=Rs5("IdArchivo2")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                  </td>
                                  <%end if%>
                                  <%end if%>
                                  
                                  <% if Rs5("IdArchivo3")<>"" then%>
                                  <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\"& Rs5("CodProf") &"\"& Rs5("CodProfefolder") &"\"& Rs5("fechaupload") &"\"& Rs5("periodopublicacion") &"\"& Rs5("IdArchivo3")) = true then%>                                 
                                  <td width='4%' align="center" valign="top" onClick="ActualizaVisto('<%=Rs5("CODRAMO")%>','<%=Rs5("CODPROF")%>','0');">
                                  <a href='Doctos/<%=Rs5("codramo")%>/<%=Rs5("CodProf")%>/<%=Rs5("CodProfefolder")%>/<%=Rs5("fechaupload")%>/<%=Rs5("periodopublicacion")%>/<%=Rs5("IdArchivo3")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                  </td>
                                  <%end if%>
                                  <%end if%>
                                  
                              <%end if %>
                         <%end if %> 
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
          
          
          <td width='95%' valign='top' class="tex-totales-celda"><span class="tex-totales-celda">
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
Set Rs2 = session("conn").Execute("SELECT a.ANO, a.PERIODO, a.RAMOEQUIV as CODRAMO,a.codsecc from ra_nota as a where a.codcli='" & Session ("CarreraEnCurso") & "' and a.ano='" & Rs1("ANO") & "' and a.periodo='2' order by a.ano, a.periodo")
if Not Rs2.Bof or Not Rs2.Eof _
	then
		Rs2.MoveFirst
		While Not Rs2.Eof
			Set Rs3 = session("conn").Execute("SELECT NOMBRE from ra_ramo where codramo='" & Rs2("CODRAMO") & "'")%>
                    </span>
                      <table width='100%' border='1' cellpadding='3' cellspacing='0' class='smalltext'>
                        <tr>
                          <td bgcolor='#FFFFCC'><span class="normalname"><a style='text-align:left; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs2("CODRAMO")&"_" & Rs2("codsecc")& "_" & Rs2("PERIODO")& "_" &Rs2("ANO")%>');"><%
						 
						 	if DOCTOSENACPA="SI" then
								  strD0="select DISTINCT CodProf from ca_doctos where codramo='" & Rs2("CODRAMO") & "' and AnoPublicacion= '"&trim(ano)&"' And PeriodoPublicacion='2' And Activo='1' and codcarr in (select codcarr COLLATE Modern_Spanish_CI_AS from mt_Carrer where sede='"& session("codsede") &"')"
									
								  Set Rs4 = session("conn").Execute(strD0) 
							else	   
								  strD0="select DISTINCT CodProf from ca_doctos where codramo='" & Rs2("CODRAMO") & "' and AnoPublicacion= '"&trim(ano)&"' And PeriodoPublicacion='2' And Activo='1' and codcarr in (select codcarr COLLATE Modern_Spanish_CI_AS from mt_Carrer where sede='"& session("codsede") &"')"
									
								  if MaterialSeccion="SI" then
									strD0=strD0 & "and codsecc='"& Rs2("codsecc") &"'"
								  end if 
								  Set Rs4 = session("conn").Execute(strD0) 
							end if

						  if not Rs4.eof then
						  	response.Write("<img src='addon/imagenes/btn_show-all_bg.gif' width='16' height='16'>")						  
						  end if  
						  
						  %>&nbsp;&nbsp;(<%=Rs2("CODRAMO")%>):<%=Rs3("NOMBRE")%></a>
                          </span></td>
                        </tr>
                        <tr>
                          <td valign='top' style='display:none ;' id='<%=Rs2("CODRAMO")&"_" & Rs2("codsecc")& "_" & Rs2("PERIODO")& "_" &Rs2("ANO")%>'>
                            <%

if Not Rs4.Bof or Not Rs4.Eof _
	then
		Rs4.MoveFirst
		While Not Rs4.Eof
			Set Rs6=session("conn").Execute("select * from ra_profes where codprof='" & Rs4("codprof") & "'")
			if Not Rs6.Bof or Not Rs6.Eof _ 
				then%>
                              <table width='100%' border='1' cellpadding='0' cellspacing='0' class='smalltext'>
                                <tr>
                                  <td class='maintitle'><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs2("CODRAMO") & "_" & Rs6("CODPROF")& "_" & Rs2("PERIODO")&"_" & Rs2("codsecc")& "_" & Rs2("PERIODO")& "_" &Rs2("ANO")%>');"><%=Rs6("NOMBRES") & " " & Rs6("AP_PATER") & " " & Rs6("AP_MATER")%></a></td>
                                </tr>
                                <tr>
                                  <td valign='top' style='display:none ;' id='<%=Rs2("CODRAMO") & "_" & Rs6("CodProf")& "_" & Rs2("PERIODO")&"_" & Rs2("codsecc")& "_" & Rs2("PERIODO")& "_" &Rs2("ANO")%>'>
                                    <%
TablaDocumentos=""
HeadDoctos=""
BodyDoctos=""
FooterDoctos=""  
 
 
if DOCTOSENACPA="SI" then 
	strD="select  IDCONTER,CODPROF,CODCARR,CODRAMO,TIPODOCUMENTO,REFERENCIA,DETALLEDOCUMENTO,IDARCHIVO1,IDARCHIVO2,IDARCHIVO3,CONVERT(VARCHAR(10),CONVERT(DATE,FechaUpLoad),105)FECHAUPLOAD,CONVERT(DATE,FECHAUPLOAD)x,ANOPUBLICACION,PERIODOPUBLICACION,FECHAACTUALIZACION,ACTIVO,CODPROFEFOLDER,"
	if MaterialSeccion="SI" then
		strD = strD & "CODSECC,"
	end if 
	strD = strD & "EVALUADO,FECHA_INI,FECHA_FIN,ACEPTA_MAT_FUERA_PLAZO from ca_doctos where codramo='" & Rs2("CODRAMO") & "' and CodProf='" & Rs6("CodProf") & "' and AnoPublicacion='"&trim(ano)&"' "
	strD=strD & "And PeriodoPublicacion='2' And Activo='1' and codcarr in (select codcarr COLLATE Modern_Spanish_CI_AS from mt_Carrer where sede='"& session("codsede") &"')"
	if MaterialSeccion="SI" then
		strD=strD & "and codsecc='"& Rs2("codsecc") &"'"
	end if
	
	if MaterialSeccion="SI" then
		Set Rs5 = session("conn").Execute("SELECT distinct IDCONTER,doc.CODPROF,doc.CODCARR,doc.CODRAMO,doc.TIPODOCUMENTO,REFERENCIA,DETALLEDOCUMENTO,IDARCHIVO1,IDARCHIVO2,IDARCHIVO3,CONVERT(VARCHAR(10),CONVERT(DATE,FechaUpLoad),105)FECHAUPLOAD,CONVERT(DATE,FECHAUPLOAD)x,ANOPUBLICACION,PERIODOPUBLICACION,FECHAACTUALIZACION,ACTIVO,CODPROFEFOLDER,sec.CODSECC,EVALUADO,FECHA_INI,FECHA_FIN,ACEPTA_MAT_FUERA_PLAZO FROM RA_NOTA NOTA, RA_RAMO RAMO,RA_SECCIO SEC , RA_PROFES PRO, RA_HORPROF HOR, MT_ALUMNO ALUM , CA_Doctos DOC WHERE NOTA.RAMOEQUIV = RAMO.CODRAMO AND NOTA.CODCLI = '" & Session ("CarreraEnCurso") & "' AND NOTA.CODCLI = ALUM.CODCLI AND NOTA.ANO = '" & Rs1("ANO") & "' AND HOR.CODPROF = '" & Rs4("CodProf") & "' AND NOTA.PERIODO = 2 AND NOTA.CODSECC <> 0 AND SEC.ANO = NOTA.ANO AND SEC.PERIODO = NOTA.PERIODO AND SEC.CODSECC = NOTA.CODSECC AND SEC.CODRAMO = NOTA.RAMOEQUIV AND SEC.CODSEDE = ALUM.CODSEDE AND SEC.CODRAMO = HOR.CODRAMO  AND SEC.ANO = HOR.ANO AND SEC.PERIODO = HOR.PERIODO AND SEC.CODSEDE = HOR.CODSEDE AND HOR.CODPROF = PRO.CODPROF and hor.CODPROF = doc.codprof collate Modern_Spanish_CI_AS and sec.CODRAMO = doc.codramo collate Modern_Spanish_CI_AS and sec.CODCARR = doc.codcarr collate Modern_Spanish_CI_AS  and convert (varchar(4),sec.ANO) =  doc.AnoPublicacion and convert (varchar(1),sec.PERIODO) = doc.PeriodoPublicacion  and doc.codramo='" & Rs2("CODRAMO") & "' union "& strD & " order by x desc ")
	else		
		Set Rs5 = session("conn").Execute("SELECT distinct IDCONTER,doc.CODPROF,doc.CODCARR,doc.CODRAMO,doc.TIPODOCUMENTO,REFERENCIA,DETALLEDOCUMENTO,IDARCHIVO1,IDARCHIVO2,IDARCHIVO3,CONVERT(VARCHAR(10),CONVERT(DATE,FechaUpLoad),105)FECHAUPLOAD,CONVERT(DATE,FECHAUPLOAD)x,ANOPUBLICACION,PERIODOPUBLICACION,FECHAACTUALIZACION,ACTIVO,CODPROFEFOLDER,EVALUADO,FECHA_INI,FECHA_FIN,ACEPTA_MAT_FUERA_PLAZO FROM RA_NOTA NOTA, RA_RAMO RAMO,RA_SECCIO SEC , RA_PROFES PRO, RA_HORPROF HOR, MT_ALUMNO ALUM , CA_Doctos DOC WHERE NOTA.RAMOEQUIV = RAMO.CODRAMO AND NOTA.CODCLI = '" & Session ("CarreraEnCurso") & "' AND NOTA.CODCLI = ALUM.CODCLI AND NOTA.ANO = '" & Rs1("ANO") & "' AND HOR.CODPROF = '" & Rs4("CodProf") & "' AND NOTA.PERIODO = 2 AND NOTA.CODSECC <> 0 AND SEC.ANO = NOTA.ANO AND SEC.PERIODO = NOTA.PERIODO AND SEC.CODSECC = NOTA.CODSECC AND SEC.CODRAMO = NOTA.RAMOEQUIV AND SEC.CODSEDE = ALUM.CODSEDE AND SEC.CODRAMO = HOR.CODRAMO  AND SEC.ANO = HOR.ANO AND SEC.PERIODO = HOR.PERIODO AND SEC.CODSEDE = HOR.CODSEDE AND HOR.CODPROF = PRO.CODPROF and hor.CODPROF = doc.codprof collate Modern_Spanish_CI_AS and sec.CODRAMO = doc.codramo collate Modern_Spanish_CI_AS and sec.CODCARR = doc.codcarr collate Modern_Spanish_CI_AS  and convert (varchar(4),sec.ANO) =  doc.AnoPublicacion and convert (varchar(1),sec.PERIODO) = doc.PeriodoPublicacion and doc.codramo='" & Rs2("CODRAMO") & "' union "& strD & " order by x desc ")
	end if

	
else
	strD="select  IDCONTER,CODPROF,CODCARR,CODRAMO,TIPODOCUMENTO,REFERENCIA,DETALLEDOCUMENTO,IDARCHIVO1,IDARCHIVO2,IDARCHIVO3,CONVERT(VARCHAR(10),CONVERT(DATE,FechaUpLoad),105)FECHAUPLOAD,CONVERT(DATE,FECHAUPLOAD)x,ANOPUBLICACION,PERIODOPUBLICACION,FECHAACTUALIZACION,ACTIVO,CODPROFEFOLDER,"
	if MaterialSeccion="SI" then
		strD = strD & "CODSECC,"
	end if 
	strD = strD & "EVALUADO,FECHA_INI,FECHA_FIN,ACEPTA_MAT_FUERA_PLAZO from ca_doctos where codramo='" & Rs2("CODRAMO") & "' and CodProf='" & Rs6("CodProf") & "' and AnoPublicacion='"&trim(ano)&"' "
	strD=strD & "And PeriodoPublicacion='2' And Activo='1' and codcarr in (select codcarr COLLATE Modern_Spanish_CI_AS from mt_Carrer where sede='"& session("codsede") &"')"
	if MaterialSeccion="SI" then
		strD=strD & "and codsecc='"& Rs2("codsecc") &"'"
	end if
	
	strD=strD & " order by x desc "
	Set Rs5=session("conn").Execute(strD)

end if
'response.Write(strD)


if Not Rs5.Bof or Not Rs5.Eof _
	then
		Rs5.MoveFirst%>
                                      <table width='100%' border='1' cellpadding='2' cellspacing='0'>
                                        <tr>
                                          <td width='7%' align='center' valign='middle'><span class="normalname">Fecha Publicaci&oacute;n</span></td>
                                          <td colspan='4'><span class="normalname">Detalle de Documentaci&oacute;n / Apuntes</span></td>
                                        </tr>
                                        <%
		While Not Rs5.Eof
		
											  
		%>
                                        <tr>                                          
						  <%if not isnull(Rs5("tipodocumento")) then%> 
                          		<%if DOCTOSMPWPA="SI" then%>
                         		    <% if Rs5("IdArchivo1")<>"" then%>
                                    <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\{"& Rs5("codprofefolder") &"}\"& Rs5("anopublicacion") &"\S2\"&Rs5("IdArchivo1")) = true then%>
                                    <td align='center' valign='middle'><span class="normalname"><%=Day(cDate(Rs5("FechaUpLoad"))) & "/" & Month(cDate(Rs5("FechaUpLoad"))) & "/" & year(cDate(Rs5("FechaUpLoad")))%></span></td>
                                    <td width='89%' align="left" valign="middle"><span class="normalname"><%=Rs5("REFERENCIA")%></span>&nbsp;</td>
                                    <td width='4%' align="center" valign="top" onClick="ActualizaVisto('<%=Rs5("CODRAMO")%>','<%=Rs5("CODPROF")%>','<%=Rs5("CODSECC")%>');">
                                    <a href='Doctos/<%=Rs5("codramo")%>/{<%=Rs5("codprofefolder")%>}/<%=Rs5("anopublicacion")%>/S2/<%=Rs5("IdArchivo1")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                    </td>
                                    <%end if%>
                                    <%end if%>
                                     
                                    <% if Rs5("IdArchivo2")<>"" then%>
                                    <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\{"& Rs5("codprofefolder") &"}\"& Rs5("anopublicacion") &"\S2\"&Rs5("IdArchivo2")) = true then%>                                  
                                    <td width='4%' align="center" valign="top" onClick="ActualizaVisto('<%=Rs5("CODRAMO")%>','<%=Rs5("CODPROF")%>','<%=Rs5("CODSECC")%>');">
                                    <a href='Doctos/<%=Rs5("codramo")%>/{<%=Rs5("codprofefolder")%>}/<%=Rs5("anopublicacion")%>/S2/<%=Rs5("IdArchivo2")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                    </td>
                                    <%end if%>
                                    <%end if%>                                  
                                      
                                    <% if Rs5("IdArchivo3")<>"" then%>
                                    <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\{"& Rs5("codprofefolder") &"}\"& Rs5("anopublicacion") &"\S2\"&Rs5("IdArchivo3")) = true then%>                                
                                    <td width='4%' align="center" valign="top" onClick="ActualizaVisto('<%=Rs5("CODRAMO")%>','<%=Rs5("CODPROF")%>','<%=Rs5("CODSECC")%>');">
                                    <a href='Doctos/<%=Rs5("codramo")%>/{<%=Rs5("codprofefolder")%>}/<%=Rs5("anopublicacion")%>/S2/<%=Rs5("IdArchivo3")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                    </td>
                                    <%end if%>
                                    <%end if%>
                         		<%elseif DOCTOSENACPA="SI" then%>
                                
                                    <td align='center' valign='middle'><span class="normalname"><%=Day(cDate(Rs5("FechaUpLoad"))) & "/" & Month(cDate(Rs5("FechaUpLoad"))) & "/" & year(cDate(Rs5("FechaUpLoad")))%></span></td>
                                      <td width='89%' align="left" valign="middle"><span class="normalname"><%=Rs5("REFERENCIA")%></span>&nbsp;</td>
                                      <td width='4%' align="center" valign="top">
                                      
									  <% if Rs5("IdArchivo1")<>"" then%>
                                      <% if ExisteArchivo(RutaDoctos & Rs5("CODCARR") &"\"& Rs5("CODPROF") &"\"& Rs5("CODRAMO") &"\"& Rs5("idconter")&"\"& Rs5("IdArchivo1")) = true then%>
                                      <a href='doctos/<%=Rs5("CODCARR")%>/<%=Rs5("CODPROF")%>/<%=Rs5("CODRAMO")%>/<%=Rs5("idconter")%>/<%=Rs5("IdArchivo1")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                      <%end if%>
                                      <%end if%>
                                      
									  <% if Rs5("IdArchivo2")<>"" then%>
                                      <% if ExisteArchivo(RutaDoctos & Rs5("CODCARR") &"\"& Rs5("CODPROF") &"\"& Rs5("CODRAMO") &"\"& Rs5("idconter")&"\"& Rs5("IdArchivo2")) = true then%>
                                      <a href='doctos/<%=Rs5("CODCARR")%>/<%=Rs5("CODPROF")%>/<%=Rs5("CODRAMO")%>/<%=Rs5("idconter")%>/<%=Rs5("IdArchivo2")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                      <%end if%>
                                      <%end if%>
                                      
                                      <% if Rs5("IdArchivo3")<>"" then%>
                                      <% if ExisteArchivo(RutaDoctos & Rs5("CODCARR") &"\"& Rs5("CODPROF") &"\"& Rs5("CODRAMO") &"\"& Rs5("idconter")&"\"& Rs5("IdArchivo3")) = true then%>
                                      <a href='doctos/<%=Rs5("CODCARR")%>/<%=Rs5("CODPROF")%>/<%=Rs5("CODRAMO")%>/<%=Rs5("idconter")%>/<%=Rs5("IdArchivo3")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                      <%end if%>
                                      <%end if%>		                                      								  											  
                                      </td>
                                      
                         		<%else%>
									  <% if Rs5("IdArchivo1")<>"" then%>
                                      <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\"& Rs5("codprofefolder") &"\"& Rs5("anopublicacion") &"\S2\"&Rs5("IdArchivo1")) = true then%>
                                      <td align='center' valign='middle'><span class="normalname"><%=day(cDate(Rs5("FechaUpLoad"))) & "/" & Month(cDate(Rs5("FechaUpLoad"))) & "/" & Year(cDate(Rs5("FechaUpLoad")))%></span></td>
                                      <td width='89%' align="left" valign="middle"><span class="normalname"><%=Rs5("REFERENCIA")%>&nbsp;</span></td>
                                      <td width='4%' align="center" valign="top">
                                      <a href='Doctos/<%=Rs5("codramo")%>/{<%=Rs5("codprofefolder")%>}/<%=Rs5("anopublicacion")%>/S2/<%=Rs5("IdArchivo1")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                      </td>
                                      <%end if%>
                                      <%end if%>
                                      
                                      <% if Rs5("IdArchivo2")<>"" then%>
                                      <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\"& Rs5("codprofefolder") &"\"& Rs5("anopublicacion") &"\S2\"&Rs5("IdArchivo2")) = true then%>                                  
                                      <td width='4%' align="center" valign="top">
                                      <a href='Doctos/<%=Rs5("codramo")%>/{<%=Rs5("codprofefolder")%>}/<%=Rs5("anopublicacion")%>/S2/<%=Rs5("IdArchivo2")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                      </td>
                                      <%end if%>
                                      <%end if%>
                                      
                                      <% if Rs5("IdArchivo3")<>"" then%>
                                      <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\"& Rs5("codprofefolder") &"\"& Rs5("anopublicacion") &"\S2\"&Rs5("IdArchivo3")) = true then%>                                  
                                      <td width='4%' align="center" valign="top">
                                      <a href='Doctos/<%=Rs5("codramo")%>/{<%=Rs5("codprofefolder")%>}/<%=Rs5("anopublicacion")%>/S2/<%=Rs5("IdArchivo3")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                      </td>
                                      <%end if%>
                                      <%end if%>
                                <%end if%>                                 
                         <%else%> 
                              <%if MaterialSeccion="SI" then%>                                            
                                  <% if Rs5("IdArchivo1")<>"" then%>
                                  <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\"& Rs5("CodProf") &"\"& Rs5("CodProfefolder") &"\"& Rs5("fechaupload") &"\"& Rs5("periodopublicacion") &"\"& trim(Rs5("codsecc")) &"\"& Rs5("IdArchivo1")) = true then%>
                                  <td align='center' valign='middle'><span class="normalname"><%=day(cDate(Rs5("FechaUpLoad"))) & "/" & Month(cDate(Rs5("FechaUpLoad"))) & "/" & Year(cDate(Rs5("FechaUpLoad")))%></span></td>
                                  <td width='89%' align="left" valign="middle"><span class="normalname"><%=Rs5("REFERENCIA")%>&nbsp;</span></td>
                                  <td width='4%' align="center" valign="top">
                                  <a href='Doctos/<%=Rs5("codramo")%>/<%=Rs5("CodProf")%>/<%=Rs5("CodProfefolder")%>/<%=Rs5("fechaupload")%>/<%=Rs5("periodopublicacion")%>/<%=trim(Rs5("codsecc"))%>/<%=Rs5("IdArchivo1")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                  </td>
                                  <%end if%>
                                  <%end if%>
                                  
                                  <% if Rs5("IdArchivo2")<>"" then%>
                                  <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\"& Rs5("CodProf") &"\"& Rs5("CodProfefolder") &"\"& Rs5("fechaupload") &"\"& Rs5("periodopublicacion") &"\"& trim(Rs5("codsecc")) &"\"& Rs5("IdArchivo2")) = true then%>                                  
                                  <td width='4%' align="center" valign="top">
                                  <a href='Doctos/<%=Rs5("codramo")%>/<%=Rs5("CodProf")%>/<%=Rs5("CodProfefolder")%>/<%=Rs5("fechaupload")%>/<%=Rs5("periodopublicacion")%>/<%=trim(Rs5("codsecc"))%>/<%=Rs5("IdArchivo2")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                  </td>
                                  <%end if%>
                                  <%end if%>
                                  
                                  <% if Rs5("IdArchivo3")<>"" then%>
                                  <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\"& Rs5("CodProf") &"\"& Rs5("CodProfefolder") &"\"& Rs5("fechaupload") &"\"& Rs5("periodopublicacion") &"\"& trim(Rs5("codsecc")) &"\"& Rs5("IdArchivo3")) = true then%>                                  
                                  <td width='4%' align="center" valign="top">
                                  <a href='Doctos/<%=Rs5("codramo")%>/<%=Rs5("CodProf")%>/<%=Rs5("CodProfefolder")%>/<%=Rs5("fechaupload")%>/<%=Rs5("periodopublicacion")%>/<%=trim(Rs5("codsecc"))%>/<%=Rs5("IdArchivo3")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                  </td>
                                  <%end if%>
                                  <%end if%>
                              
                              <%else%>                                           
                                  <% if Rs5("IdArchivo1")<>"" then%>
                                  <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\"& Rs5("CodProf") &"\"& Rs5("CodProfefolder") &"\"& Rs5("fechaupload") &"\"& Rs5("periodopublicacion") &"\"& Rs5("IdArchivo1")) = true then%>
                                  <td align='center' valign='middle'><span class="normalname"><%=day(cDate(Rs5("FechaUpLoad"))) & "/" & Month(cDate(Rs5("FechaUpLoad"))) & "/" & Year(cDate(Rs5("FechaUpLoad")))%></span></td>
                                  <td width='89%' align="left" valign="middle"><span class="normalname"><%=Rs5("REFERENCIA")%>&nbsp;</span></td>
                                  <td width='4%' align="center" valign="top">
                                  <a href='Doctos/<%=Rs5("codramo")%>/<%=Rs5("CodProf")%>/<%=Rs5("CodProfefolder")%>/<%=Rs5("fechaupload")%>/<%=Rs5("periodopublicacion")%>/<%=Rs5("IdArchivo1")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                  </td>
                                  <%end if%>
                                  <%end if%>
                                  
                                  <% if Rs5("IdArchivo2")<>"" then%>
                                  <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\"& Rs5("CodProf") &"\"& Rs5("CodProfefolder") &"\"& Rs5("fechaupload") &"\"& Rs5("periodopublicacion") &"\"& Rs5("IdArchivo2")) = true then%>                                  
                                  <td width='4%' align="center" valign="top">
                                  <a href='Doctos/<%=Rs5("codramo")%>/<%=Rs5("CodProf")%>/<%=Rs5("CodProfefolder")%>/<%=Rs5("fechaupload")%>/<%=Rs5("periodopublicacion")%>/<%=Rs5("IdArchivo2")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                  </td>
                                  <%end if%>
                                  <%end if%>
                                  
                                  <% if Rs5("IdArchivo3")<>"" then%>
                                  <% if ExisteArchivo(RutaDoctos & Rs5("codramo") &"\"& Rs5("CodProf") &"\"& Rs5("CodProfefolder") &"\"& Rs5("fechaupload") &"\"& Rs5("periodopublicacion") &"\"& Rs5("IdArchivo3")) = true then%>                                  
                                  <td width='4%' align="center" valign="top">
                                  <a href='Doctos/<%=Rs5("codramo")%>/<%=Rs5("CodProf")%>/<%=Rs5("CodProfefolder")%>/<%=Rs5("fechaupload")%>/<%=Rs5("periodopublicacion")%>/<%=Rs5("IdArchivo3")%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a>
                                  </td>
                                  <%end if%>
                                  <%end if%>
                                  
                              <%end if %>
                         <%end if %> 
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
		
<Script>
	expandit('<%=Rs1("ANO")%>');expandit('<%=Rs1("ANO") & "_I" %>');
</script>
   <%
		  Rs1.MoveNext
				wend
			end if%>

<table>
           <tr>
    <td width="83" align="left"><A onMouseOver="MM_swapImage('Image7','','Imagenes/botones/volver-on.gif',1)" onmouseout=MM_swapImgRestore() href="javascript: window.top.location.href ='menu_material.asp'"><IMG src="Imagenes/botones/volver-of.gif" border="0"  name="Image7"></A></td>
  </tr>
  </table>
			  </tr>  
              
			</table>
			</td>
		  </tr>
         
		</table>
	</td>
  </tr>
 </table>
</body>
<%
		


function ExisteArchivo(archivo)
	'response.Write(archivo)
	dim fs
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	if fs.FileExists(archivo) then
	  ExisteArchivo = true  'response.write("File c:\asp\introduction.asp exists!")
	else
	  ExisteArchivo = false 'response.write("File c:\asp\introduction.asp does not exist!")
	end if
	set fs=nothing 
	
end function

%>

</html>
<!--#INCLUDE file="include/desconexion.inc" -->  
