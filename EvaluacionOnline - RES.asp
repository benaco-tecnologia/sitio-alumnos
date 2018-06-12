<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#include file="addon/includes/libreria.asp" --> 
<!--#include file="vars.inc.asp" -->
<!--#INCLUDE file="analytics.asp" -->

<%
strParame="SELECT dbo.Fn_ValorParame('ANALYTICSESUCOMEXPA')Parame"
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
	ANALYTICSESUCOMEXPA=rstParame("Parame")
end if

if ANALYTICSESUCOMEXPA="SI" then
	analitycsEsucomex() 
end if 
%>

<%
Exe = cInt(Request.QueryString("Cmd"))
if Exe=0 then Exe = 1 end if

%>
<%
dim fs, file 
set fs = Server.CreateObject("Scripting.FileSystemObject") 
%>
<%
RutaDoctos = Server.MapPath("Doctos")& "\"
'response.Write(RutaDoctos) 
'Response.End()


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

Set Rs1 = session("conn").Execute("SELECT DISTINCT ano FROM ra_nota WHERE codcli='" & Session ("Codcli") & "'  and ano='"&trim(ano)&"'  and periodo='"&trim(periodo)&"' ORDER BY ano")

Dim today
today=Date

if  Rs1.bof or Rs1.eof  then 
	response.Redirect("MensajeSinDatos.asp")
	response.end()
end if 


if EstaHabilitadaNW (802)="S" then 
	if GetPermisoNW(802) ="N" then
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
				session("MensajeBloqueosVarios") ="Para acceder a esta opciï¿½n debe responder sus Encuestas Pendientes."
				response.Redirect("MensajeBloqueo.asp")
			end if 
		end if 
		
	end if
else
	response.Redirect("MensajeBloqueoHabilita.asp")
end if 
Audita 802,"Ingresa a Evauación Online" 

strParame="SELECT dbo.Fn_ValorParame('MATERIAL_SECCION')Parame"
if BCL_ADO(strParame, rstParame) then
	MaterialSeccion=trim(valnulo(rstParame("Parame"),STR_))
else
	MaterialSeccion=""
end if

strParame="SELECT dbo.Fn_ValorParame('RAMOS_POR_SECCION')Parame"
if BCL_ADO(strParame, rstParame) then
	RAMOS_POR_SECCION=trim(valnulo(rstParame("Parame"),STR_))
else
	RAMOS_POR_SECCION=""
end if

strParame="SELECT dbo.Fn_ValorParame('RutaDescargaEvaluacionesOnline')Parame"
if BCL_ADO(strParame, rstParame) then
	RUTA_DOC=trim(valnulo(rstParame("Parame"),STR_))
else
	RUTA_DOC=""
end if


Set Rs2 = session("conn").Execute("SELECT a.ANO, a.PERIODO, a.RAMOEQUIV as CODRAMO from ra_nota as a where a.codcli='" & Session ("Codcli") & "' and a.CodRamo<>'' and a.ano='" & Rs1("ANO") & "' and a.periodo='1' order by a.ano, a.periodo")

if Not Rs1.Bof or Not Rs1.Eof then
	Rs1.MoveFirst
	While Not Rs1.Eof
	
	
	%>
	
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
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
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<script LANGUAGE="JavaScript">

function confirmSubmit(form)
{
var agree=confirm("No podrá realizar modificaciones luego de enviar el archivo. ¿Desea continuar?");

if (agree)
form.action = "EvaluacionesOnline_upldFile.asp";
else
return false ;
}

</script>



<Body onLoad="expandit('<%=Rs1("ANO")%>');expandit('<%=Rs1("ANO") & "_I" %>');">

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
	
		<table width="770" border='1' cellpadding='5' cellspacing='0' onKeyPress="" align="center" > 
  <tr>
    <td class="caption" background="Imagenes/fdo-cabecera-cel22.jpg" height="30"><span class="normalname" style="font-size: 14"><a class="text-cabecera-celda" style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;'>Evaluaciones Online <%=Rs1("ANO")%></a></span></td>
  </tr>
  <tr>
    <td valign='top' style='display:none ;' id='<%=Rs1("ANO")%>'><table width="100%" border='1' cellpadding='0' cellspacing='0'>
        <tr>
    
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
Set Rs2 = session("conn").Execute("SELECT a.ANO, a.PERIODO, a.RAMOEQUIV as CODRAMO,a.codsecc from ra_nota as a where a.codcli='" & Session ("Codcli") & "' and a.CodRamo<>'' and a.ano='" & Rs1("ANO") & "' and a.periodo='1' order by a.ano, a.periodo")

if Not Rs2.Bof or Not Rs2.Eof _  
	then
		Rs2.MoveFirst
		While Not Rs2.Eof
			Set Rs3 = session("conn").Execute("SELECT NOMBRE from ra_ramo where codramo='" & Rs2("CODRAMO") & "'")
%>
                      <table width='100%' border='1' cellpadding='3' cellspacing='0' class='smalltext'>
                        <tr> 
                          <td align="left" valign="middle" bgcolor='#FFFFCC'><span class="normalname"><a style='text-align:left; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs2("ANO") & "_" & Rs2("PERIODO") & "_" & Rs2("CODRAMO")%>');">
                          <%if ConterProfesRamos(Rs2("CODRAMO"), 1,"")>0 then response.Write("<img src='addon/imagenes/btn_show-all_bg.gif' width='16' height='16'>") end if%>&nbsp;&nbsp;<%=Rs2("CODRAMO")%>: <%=Rs3("NOMBRE")%></a></span></td>
                        </tr>
                        <tr >
                          <td valign='top'  style='display:none ;' id='<%=Rs2("ANO") & "_" & Rs2("PERIODO") & "_" & Rs2("CODRAMO")%>'>
                            <%

TablaProfesores=""
strD0="select DISTINCT CodProf from ca_doctos where codramo='" & Rs2("CODRAMO") & "' and AnoPublicacion= '"&trim(ano)&"' And PeriodoPublicacion='1' And Activo='1' and codcarr in (select codcarr COLLATE Modern_Spanish_CI_AS from mt_Carrer where sede='"& session("codsede") &"')"

if MaterialSeccion="SI" then
	strD0=strD0 & "and codsecc='"& Rs2("codsecc") &"'"
end if

Set Rs4 = session("conn").Execute(strD0)
if Not Rs4.Bof or Not Rs4.Eof _
	then
		Rs4.MoveFirst
		While Not Rs4.Eof
			Set Rs6=session("conn").Execute("select * from ra_profes where codprof='" & Rs4("CodProf") & "'")
			if Not Rs6.Bof or Not Rs6.Eof _ 
				then%>
                              <table width='100%' border='1' cellpadding='0' cellspacing='0' class='smalltext'>
                                <tr>
                                  <td class='maintitle' ><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs2("CODRAMO") & "_" & Rs6("CODPROF")%>');"><%=Rs6("NOMBRES") & " " & Rs6("AP_PATER") & " " & Rs6("AP_MATER")%></a></td>
                                </tr>
                                <tr>
                                  <td valign='top' style='display:none ;' id='<%=Rs2("CODRAMO") & "_" & Rs6("CodProf")%>'>
                                    <%
TablaDocumentos=""
HeadDoctos=""
BodyDoctos=""
FooterDoctos=""

strD="select * from ca_doctos where codramo='" & Rs2("CODRAMO") & "' and CodProf='" & Rs6("CodProf") & "' and AnoPublicacion='"&trim(ano)&"' "
strD=strD & "And PeriodoPublicacion='1' And Activo='1' and EVALUADO = 'SI' and codcarr in (select codcarr COLLATE Modern_Spanish_CI_AS from mt_Carrer where sede='"& session("codsede") &"')"
if MaterialSeccion="SI" then
	strD=strD & "and codsecc='"& Rs2("codsecc") &"'"
end if
'response.Write(strD)

Set Rs5=session("conn").Execute(strD)




				
if Not Rs5.Bof or Not Rs5.Eof _
	then
		Rs5.MoveFirst%>
                                      <table width='100%' border='1' cellpadding='2' cellspacing='0'>
                                        <tr>
                                          <td><span class="normalname">Detalle de Documentaci&oacute;n / Apuntes </span></td>
                                           <td width='7%' align='center' valign='middle'><span class="normalname">Fecha de publicación</span></td>
                                          <td width='7%' align='center' valign='middle'><span class="normalname">Fecha de entrega inicial</span></td>
                                          <td width='7%' align='center' valign='middle'><span class="normalname">Fecha de entrega final</span></td>
                                          
                                         
                                        </tr>
                                        <%
	While Not Rs5.Eof%>
    
    <%   
   'RUTA DE ARCHIVOS                        
archivo1 = "Doctos" +"/"+ Rs5("CODRAMO") + "/" + Rs5("CODPROF") + "/" + Rs5("CODPROFEFOLDER") + "/" + Rs5("FECHAUPLOAD") + "/" + Rs5("PERIODOPUBLICACION") + "/" + Rs5("IdArchivo1")

archivo2 = "Doctos" +"/"+ Rs5("CODRAMO") + "/" + Rs5("CODPROF") + "/" + Rs5("CODPROFEFOLDER") + "/" + Rs5("FECHAUPLOAD") + "/" + Rs5("PERIODOPUBLICACION") + "/" + Rs5("IdArchivo2")

archivo3 = "Doctos" +"/"+ Rs5("CODRAMO") + "/" + Rs5("CODPROF") + "/" + Rs5("CODPROFEFOLDER") + "/" + Rs5("FECHAUPLOAD") + "/" + Rs5("PERIODOPUBLICACION") + "/" + Rs5("IdArchivo3")

if RAMOS_POR_SECCION = "SI" then
archivo1 = "Doctos" +"/"+ Rs5("CODRAMO") + "/" + Rs5("CODPROF") + "/" + Rs5("CODPROFEFOLDER") + "/" + Rs5("FECHAUPLOAD") + "/" + Rs5("PERIODOPUBLICACION") + "/" + Rs5("CODSECC") + "/" + Rs5("IdArchivo1")

archivo2 = "Doctos" +"/"+ Rs5("CODRAMO") + "/" + Rs5("CODPROF") + "/" + Rs5("CODPROFEFOLDER") + "/" + Rs5("FECHAUPLOAD") + "/" + Rs5("PERIODOPUBLICACION") + "/" + Rs5("CODSECC") + "/" + Rs5("IdArchivo2")

archivo3 = "Doctos" +"/"+ Rs5("CODRAMO") + "/" + Rs5("CODPROF") + "/" + Rs5("CODPROFEFOLDER") + "/" + Rs5("FECHAUPLOAD") + "/" + Rs5("PERIODOPUBLICACION") + "/" + Rs5("CODSECC") + "/" + Rs5("IdArchivo3")
end if

if not isnull(Rs5("tipodocumento")) then

archivo1 = "Doctos" +"/"+ Rs5("CODRAMO") + "/" + Rs5("CODPROFEFOLDER") + "/" + Rs5("ANOPUBLICACION") + "/S1" +  Rs5("IdArchivo1")

archivo2 = "Doctos" +"/"+ Rs5("CODRAMO") + "/" + Rs5("CODPROFEFOLDER") + "/" + Rs5("ANOPUBLICACION") + "/S1" +  Rs5("IdArchivo2")

archivo3 = "Doctos" +"/"+ Rs5("CODRAMO") + "/" + Rs5("CODPROFEFOLDER") + "/" + Rs5("ANOPUBLICACION") + "/S1" +  Rs5("IdArchivo3")

end if

%>
<%
 strMaterialSubido = "Select * from ca_evaluacionesOnline where ID_CONTER = " & Rs5("IDCONTER") &" AND CODCLI = '" & session("Codcli") & "'"
Set Rs7=session("conn").Execute(strMaterialSubido)
if Not Rs7.Bof or Not Rs7.Eof _
	then										
		subido="SI"
		else
		subido="NO"
end if	
%>
                              <tr>

                        
                                 
                                   <td width='52%' align="left" valign="middle">
                                   <span class="normalname">Tema: <%=Rs5("DetalleDocumento")%>
                                   </span>&nbsp;</td>
                                    <td width='11%' align='center' valign='middle'><span class="normalname"><%=Day(cDate(Rs5("FECHAUPLOAD"))) & "/" & Month(cDate(Rs5("FECHAUPLOAD"))) & "/" & year(cDate(Rs5("FECHAUPLOAD")))%></span></td>
                                  <td width='11%' align='center' valign='middle'><span class="normalname"><%=Day(cDate(Rs5("FECHA_INI"))) & "/" & Month(cDate(Rs5("FECHA_INI"))) & "/" & year(cDate(Rs5("FECHA_INI")))%></span></td>
                                  <td width='11%' align='center' valign='middle'><span class="normalname"><%=Day(cDate(Rs5("FECHA_FIN"))) & "/" & Month(cDate(Rs5("FECHA_FIN"))) & "/" & year(cDate(Rs5("FECHA_FIN")))%></span></td>
                              	  <% if Rs5("IdArchivo1")<>"" or Rs5("IdArchivo1")<>null then%>
                                  <td width='4%' align="center" valign="top"><div title="descargar material">
                                  <a href='<%=archivo1%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a></div>
                                  </td>
                                  <%end if%>
                                  
                                  <% if Rs5("IdArchivo2")<>"" or Rs5("IdArchivo2")<>null then%>  
                                  <td width='4%' align="center" valign="top"><div title="descargar material">
                                  <a href='<%=archivo2%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a></div>
                                  </td>
                                  <%end if%>                                  
                                  
                                  <% if Rs5("IdArchivo3")<>"" or Rs5("IdArchivo3")<>null then%>
                                  <td width='4%' align="center" valign="top"><div title="descargar material">
                                  <a href='<%=archivo3%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a></div>
                                  </td>
                                  <%end if%>
                                   
                                        </tr>
                                        <tr>
                                    
                                        <%'UPLOAD%>
                                        <td colspan="16">
                                  
                                       <%if subido = "NO" then%>
                                        
                                        <% if   today >= Rs5("FECHA_INI") then  %>
                                        <% if   today =< Rs5("FECHA_FIN") then  %>
                                        <form method="POST" name="envio" onsubmit= "confirmSubmit(this)" enctype="multipart/form-data">
                                        
                                          Seleccione el archivo que desea adjuntar <input name="Docto" type="file"  id="Docto" size="50" >
                                          <input name='goSave' type="submit" class='button' id='goSave' value='Enviar' >
                                        
                							<input type = "hidden" name = "CodCli" value = "<%=Session ("Codcli")%>" >
                							<input type = "hidden" name = "Conter" value = "<%=Rs5("IDCONTER")%>" >
                                            <input type = "hidden" name = "Atrasado" value = "NO" >
                                        </form>
                                           
                                          <%else%>
                                           <% if Rs5("ACEPTA_MAT_FUERA_PLAZO") = "SI" then 'FUERA DE PLAZO %>
                                            <form method="POST" name="envio" onsubmit= "confirmSubmit(this)" enctype="multipart/form-data">
                                          Seleccione el archivo que desea adjuntar <input name="Docto" type="file"  id="Docto" size="50" >
                                          <input name='goSave' type="submit" class='button' id='goSave' value='Enviar'>
                                        
                							<input type = "hidden" name = "CodCli" value = "<%=Session ("Codcli")%>">
                							<input type = "hidden" name = "Conter" value = "<%=Rs5("IDCONTER")%>">
                                            <input type = "hidden" name = "Atrasado" value = "SI">
                                        </form>
                                         
                                           <%else%>
                                          <% Response.Write("El plazo para subir material ha vencido") %>
                                           <%end if %> 
                                          <%end if %>
                                           <%else%>
                                          <% Response.Write("El plazo para subir material no ha comenzado") %>
                                           <%end if %>
                                            <%else%>
                                           <% Response.Write("El documento ya ha sido enviado.") %>  <a href="<%= RUTA_DOC & Rs5("IDCONTER") &"/"& Rs7("ARCHIVO")%>">
   ver documento 
</a>
                                             <%end if%>
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
Set Rs2 = session("conn").Execute("SELECT a.ANO, a.PERIODO, a.RAMOEQUIV as CODRAMO,a.codsecc from ra_nota as a where a.codcli='" & Session ("Codcli") & "' and a.ano='" & Rs1("ANO") & "' and a.periodo='2' order by a.ano, a.periodo")

if Not Rs2.Bof or Not Rs2.Eof _
	then
		Rs2.MoveFirst
		While Not Rs2.Eof
			Set Rs3 = session("conn").Execute("SELECT NOMBRE from ra_ramo where codramo='" & Rs2("CODRAMO") & "'")%>
                    </span>
                      <table width='100%' border='1' cellpadding='3' cellspacing='0' class='smalltext'>
                        <tr>
                          <td bgcolor='#FFFFCC'><span class="normalname"><a style='text-align:left; cursor:hand; cursor:pointer; text-decoration:none;' onClick="expandit('<%=Rs2("CODRAMO")%>');"><%if ConterProfesRamos(Rs2("CODRAMO"), 2,"")>0 then response.Write("<img src='addon/imagenes/btn_show-all_bg.gif' width='16' height='16'>") end if%>&nbsp;&nbsp;(<%=Rs2("CODRAMO")%>):<%=Rs3("NOMBRE")%></a>
                          </span></td>
                        </tr>
                        <tr>
                          <td valign='top' style='display:none ;' id='<%=Rs2("CODRAMO")%>'>
                            <%
Set Rs4 = session("conn").Execute("select DISTINCT codprof from ca_doctos where codramo='" & Rs2("CODRAMO") & "' and AnoPublicacion='"&trim(ano)&"' And PeriodoPublicacion='2' And Activo='1' and codcarr in (select codcarr COLLATE Modern_Spanish_CI_AS from mt_Carrer where sede='"& session("codsede") &"')")

if Not Rs4.Bof or Not Rs4.Eof _
	then
		Rs4.MoveFirst
		While Not Rs4.Eof
			Set Rs6=session("conn").Execute("select * from ra_profes where codprof='" & Rs4("codprof") & "'")
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
 
strD="select * from ca_doctos where codramo='" & Rs2("CODRAMO") & "' and CodProf='" & Rs6("CodProf") & "' and AnoPublicacion='"&trim(ano)&"' "
strD=strD & "And PeriodoPublicacion='2' And Activo='1' and codcarr in (select codcarr COLLATE Modern_Spanish_CI_AS from mt_Carrer where sede='"& session("codsede") &"')"
if MaterialSeccion="SI" then
	strD=strD & "and codsecc='"& Rs2("codsecc") &"'"
end if
Set Rs5=session("conn").Execute(strD)

if Not Rs5.Bof or Not Rs5.Eof _
	then
		Rs5.MoveFirst%>
                                      <table width='100%' border='1' cellpadding='2' cellspacing='0'>
                                        <tr>
                                         <td><span class="normalname">Detalle de Documentaci&oacute;n / Apuntes</span></td>
                                          <td width='7%' align='center' valign='middle'><span class="normalname">Fecha Inicio</span></td>
                                           <td width='7%' align='center' valign='middle'><span class="normalname">Fecha Entrega</span></td>
                                        </tr>
                                        <%
		While Not Rs5.Eof
		
											  
		%>
                                                
 <%   
   'RUTA DE ARCHIVOS                        
archivo1 = "Doctos" +"/"+ Rs5("CODRAMO") + "/" + Rs5("CODPROF") + "/" + Rs5("CODPROFEFOLDER") + "/" + Rs5("FECHAUPLOAD") + "/" + Rs5("PERIODOPUBLICACION") + "/" + Rs5("IdArchivo1")

archivo2 = "Doctos" +"/"+ Rs5("CODRAMO") + "/" + Rs5("CODPROF") + "/" + Rs5("CODPROFEFOLDER") + "/" + Rs5("FECHAUPLOAD") + "/" + Rs5("PERIODOPUBLICACION") + "/" + Rs5("IdArchivo2")

archivo3 = "Doctos" +"/"+ Rs5("CODRAMO") + "/" + Rs5("CODPROF") + "/" + Rs5("CODPROFEFOLDER") + "/" + Rs5("FECHAUPLOAD") + "/" + Rs5("PERIODOPUBLICACION") + "/" + Rs5("IdArchivo3")

if RAMOS_POR_SECCION = "SI" then
archivo1 = "Doctos" +"/"+ Rs5("CODRAMO") + "/" + Rs5("CODPROF") + "/" + Rs5("CODPROFEFOLDER") + "/" + Rs5("FECHAUPLOAD") + "/" + Rs5("PERIODOPUBLICACION") + "/" + Rs5("CODSECC") + "/" + Rs5("IdArchivo1")

archivo2 = "Doctos" +"/"+ Rs5("CODRAMO") + "/" + Rs5("CODPROF") + "/" + Rs5("CODPROFEFOLDER") + "/" + Rs5("FECHAUPLOAD") + "/" + Rs5("PERIODOPUBLICACION") + "/" + Rs5("CODSECC") + "/" + Rs5("IdArchivo2")

archivo3 = "Doctos" +"/"+ Rs5("CODRAMO") + "/" + Rs5("CODPROF") + "/" + Rs5("CODPROFEFOLDER") + "/" + Rs5("FECHAUPLOAD") + "/" + Rs5("PERIODOPUBLICACION") + "/" + Rs5("CODSECC") + "/" + Rs5("IdArchivo3")
end if

if not isnull(Rs5("tipodocumento")) then

archivo1 = "Doctos" +"/"+ Rs5("CODRAMO") + "/" + Rs5("CODPROFEFOLDER") + "/" + Rs5("ANOPUBLICACION") + "/S1" +  Rs5("IdArchivo1")

archivo2 = "Doctos" +"/"+ Rs5("CODRAMO") + "/" + Rs5("CODPROFEFOLDER") + "/" + Rs5("ANOPUBLICACION") + "/S1" +  Rs5("IdArchivo2")

archivo3 = "Doctos" +"/"+ Rs5("CODRAMO") + "/" + Rs5("CODPROFEFOLDER") + "/" + Rs5("ANOPUBLICACION") + "/S1" +  Rs5("IdArchivo3")

end if

%> 
                                        <tr>    
                                                                            
						
                                  <td align='center' valign='middle'><span class="normalname"><%=day(cDate(Rs5("FechaUpLoad"))) & "/" & Month(cDate(Rs5("FechaUpLoad"))) & "/" & Year(cDate(Rs5("FechaUpLoad")))%></span></td>
                                  <td width='89%' align="left" valign="middle">
                                  <span class="normalname">Tema: <%=Rs5("DetalleDocumento")%>
                                  &nbsp;</span></td>
                                  <td width='4%' align="center" valign="top">
                                    <% if Rs5("IdArchivo1")<>"" then%>  <div title="descargar material">
                                  <a href='<%=archivo1%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a></div>
                                  </td><br>
                                  <%end if%>
                                  
                                  <% if Rs5("IdArchivo2")<>"" then%>         
                                  <td width='4%' align="center" valign="top"><div title="descargar material">
                                  <a href='<%=archivo2%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a></div>
                                  </td><br>
                                  <%end if%>
                                  
                                  <% if Rs5("IdArchivo3")<>"" then%>          
                                  <td width='4%' align="center" valign="top"><div title="descargar material">
                                  <a href='<%=archivo3%>'><img src='addon/imagenes/IconoDescarga[1].gif' width='25' height='25' border='none'></a></div>
                                  </td><br>
                                  <%end if%>
                                                                 
                      
                                        </tr>
                                        
                                        <tr>
                                        <%'UPLOAD%>
                                        <td colspan="16">
                                         <%if subido = "NO" then%>
                                        
                                        <% if   today >= Rs5("FECHA_INI") then  %>
                                        <% if   today =< Rs5("FECHA_FIN") then  %>
                                        <form method="POST" name="envio" onsubmit= "confirmSubmit(this)" enctype="multipart/form-data">
                                        
                                          Seleccione el archivo que desea adjuntar <input name="Docto" type="file"  id="Docto" size="50" >
                                          <input name='goSave' type="submit" class='button' id='goSave' value='Enviar' >
                                        
                							<input type = "hidden" name = "CodCli" value = "<%=Session ("Codcli")%>" >
                							<input type = "hidden" name = "Conter" value = "<%=Rs5("IDCONTER")%>" >
                                            <input type = "hidden" name = "Atrasado" value = "NO" >
                                        </form>
                                           
                                          <%else%>
                                           <% if Rs5("ACEPTA_MAT_FUERA_PLAZO") = "SI" then 'FUERA DE PLAZO %>
                                            <form method="POST" name="envio" onsubmit= "confirmSubmit(this)" enctype="multipart/form-data">
                                         Seleccione el archivo que desea adjuntar <input name="Docto" type="file"   id="Docto" size="50" >
                                          <input name='goSave' type="submit" class='button' id='goSave' value='Enviar'>
                                        
                							<input type = "hidden" name = "CodCli" value = "<%=Session ("Codcli")%>">
                							<input type = "hidden" name = "Conter" value = "<%=Rs5("IDCONTER")%>">
                                            <input type = "hidden" name = "Atrasado" value = "SI">
                                        </form>
                                         
                                           <%else%>
                                          <% Response.Write("El plazo para subir material ha vencido") %>
                                           <%end if %> 
                                          <%end if %>
                                           <%else%>
                                          <% Response.Write("El plazo para subir material no ha comenzado") %>
                                           <%end if %>
                                            <%else%>
                                           <% Response.Write("El documento ya ha sido enviado.") %>  
                                           <%end if%>
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
		Rs1.MoveNext
	wend
end if


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
