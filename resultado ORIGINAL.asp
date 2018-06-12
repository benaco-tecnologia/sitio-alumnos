<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
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

<script language="JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<link rel="stylesheet" href="css/tex-normales.css" type="text/css">
<body  leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#FFFFFF">
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=RamosDebe name=RamosDebe>
</OBJECT> 
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=RamosSolici name=RamosSolici>
</OBJECT> 
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=RamosPuede name=RamosPuede>
</OBJECT> 

<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if
%>
<%
Dim strRamosDebe, strRamosPuede, Rut, CodCli
CodCli = Session("CodCli")
CodSede = Session("CodSede")
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if

ano =AnoAcad()
periodo=PeriodoAcad()

strANOP="select ANO,PERIODO from MT_CARRERAPERIODO WHERE CODCARR='"& session("codcarr") &"' AND codpestud='"& session("codpestud") &"' "

if bcl_ado(strANOP,rstANOP) then
	ano =valnulo(rstANOP("ANO"),NUM_)
	periodo=valnulo(rstANOP("PERIODO"),NUM_)
end if

if EstaHabilitadaNW (470)="S" then 
	if GetPermisoNW(470) ="N" then
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

Audita 470,"Ingresa a Resumen de inscripcion"

dim TOTALCREDITOS 
TOTALCREDITOS = 0

strParame="SELECT dbo.Fn_ValorParame('MUESTRA_OPORTUNIDAD')Parame"
if BCL_ADO(strParame, rstParame) then
	MUESTRA_OPORTUNIDAD=trim(valnulo(rstParame("Parame"),STR_))
end if

strParamePA="SELECT dbo.Fn_ValorParame('MUESTRA_CREDITOS_PA')Parame"
if BCL_ADO(strParamePA, rstParamePA) then
	MUESTRA_CREDITOS_PA=trim(valnulo(rstParamePA("Parame"),STR_))
end if

strPRXPA="SELECT dbo.Fn_ValorParame('USAVISTAXPROFERAMORESUMENPA')Parame"
if BCL_ADO(strPRXPA, rstPRXPA) then
	USAVISTAXPROFERAMORESUMENPA=trim(valnulo(rstPRXPA("Parame"),STR_))
end if



'Periodo = Session("Periodo")
'desde la tabla mt_parame obtener el año y periodo
strRamosDebe = "sp_listaramosalumnoportal '" & CodCli & "','" & ano & "','" & periodo & "'"

'response.write(strRamosDebe)
'response.End()			     
		  
				' Order By a.prioridad "
'agregar el año y periodo a esta query
RamosDebe.Open strRamosDebe, Session("Conn")

strRamosSolici = "SELECT b.codramo, b.nombre, a.estado, a.codsecc, a.Ramoequiv, c.nombre as NombreReal, a.CodSeccOri " & _
                " FROM ra_solici a, ra_ramo b, ra_ramo c " & _
                " WHERE a.ramoequiv = b.codramo and " & _
                " a.ramoequiv *= c.codramo and " & _
                " a.codcli ='" & CodCli & "' and " & _
                " a.ano = '" & ano & "' and " & _
                " a.PerInscrip = " & SESSION("PER_ID") & " and " & _
                " a.periodo = '" & periodo & "' "

'agregar el año y periodo a esta query
'RESPONSE.WRITE(strRamosSolici)
'RESPONSE.End

RamosSolici.Open strRamosSolici, Session("Conn")
dim paso 
paso=0
do while not RamosSolici.eof and paso=0
    if RamosSolici("codsecc")<>RamosSolici("CodSeccori") then
	 paso=1	  
    end if
   RamosSolici.movenext
loop
RamosSolici.close

strRamosSolici = "SELECT b.codramo, b.nombre, a.estado, a.codsecc, a.Ramoequiv, c.nombre as NombreReal, CodSeccOri " & _
                " FROM ra_solici a, ra_ramo b, ra_ramo c " & _
                " WHERE a.ramoequiv = b.codramo and " & _
                " a.ramoequiv *= c.codramo and " & _
                " a.codcli ='" & CodCli & "' and " & _
                " a.ano = '" & ano & "' and " & _
                " a.periodo = '" & periodo & "'  "
'agregar el año y periodo a esta query
'RESPONSE.WRITE(strRamosSolici)
'RESPONSE.End 
RamosSolici.Open strRamosSolici, Session("Conn")

fechaCalendarioinicio=""
fechaCalendariofin=""

strCalendar="MT_CARRERAPERIODO_BUSCAPERIODO2 '"& session("codcarr") &"'"

if bcl_ado(strCalendar,RstCalendar) then
   strFechaCalendar ="SP_FechaClases "& RstCalendar("num_max_periodo") &", "& RstCalendar("ano") &", "& RstCalendar("periodo") &", '"& session("codcarr") &"'"  

   if bcl_ado(strFechaCalendar,RstFechaCalendar) then
		if isnull(RstFechaCalendar("fecha")) then
			fechaCalendarioinicio=""
		else
			fechaCalendarioinicio=RstFechaCalendar("fecha")
		end if 
		if isnull(RstFechaCalendar("fecha1")) then
			fechaCalendariofin=""
		else
			fechaCalendariofin=RstFechaCalendar("fecha1")
		end if 
   end if 
end if 

	
%>
<Script Language="JavaScript">
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
</Script><html>
<head> 
<title>Portal de Alumnos en L&iacute;nea</title>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1"> 
</head>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript"  SRC="include/tooltip.js"></SCRIPT>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
<%' if not InscritoNormal() then %>
  <script>
  //   alert("No ha inscrito aún vía solicitud Normal ...");
  </script>   
<%'  end if%>
<table width="1049" border="0"  align="left" cellpadding="0" cellspacing="0">
  <tr>
	<td>
		<table width="862" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right" ><% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td width="795" valign="top">
			<% CargarTop1()%><% SubMenu()%>	
				<!--<div style="width:825PX; height:480px; overflow:scroll;" align="center">-->
                
				  <form name="form1" method="post" action="asignatura-seccion.asp">
					<input type = "hidden" name = "Ramos" value = "">
					<table border="0" cellpadding="0" cellspacing="15" width="962">
					  <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
					  <tr valign="top" > 
						<td ><img name="fder_r2_c1" src="imagenes/titulos/T-resumen_inscripcion.gif" width="471" height="38" border="0"></tr>
						<tr valign="top" > 
                        <% 
						strParameMODIPA="SELECT dbo.Fn_ValorParame('PERMITEMODIFICARCARGAPA')Parame"
						if BCL_ADO(strParameMODIPA, rstParameMODIPA) then
							PERMITEMODIFICARCARGAPA=trim(valnulo(rstParameMODIPA("Parame"),STR_))
						end if
						
						strValidaNuevo ="select nuevo from mt_alumno where codcli='"& session("codcli") &"'"
						if BCL_ADO(strValidaNuevo, rstValidaNuevo) then
							VALIDANUEVO=trim(valnulo(rstValidaNuevo("nuevo"),STR_))
						end if
						
						if ucase(PERMITEMODIFICARCARGAPA) = "SI" and SESSION("PER_ID") <> "-1" and VALIDANUEVO <> "S" then  %>                        
                        <td width="83" align="left"><A onMouseOver="MM_swapImage('Image123','','Imagenes/botones/modificar-on.jpg',1)" onmouseout=MM_swapImgRestore() href="#" onClick="ConfirmaModi();"><IMG src="Imagenes/botones/modificar-of.jpg" border="0"  name="Image123"></A></td>
                        <%end if%>
						<td  align="left"><a href="javascript:imprime()"><img src="imag/f-der/impresion.gif" width="36" height="36" border="0" align="right"></a></tr>
					  <tr valign="top">
					  	<td height="1" colspan=3><span class="Tit-celdas"><b class="Tit-celdas"></b></span><span style="font-size: 13px; font-weight: bold; font-family: Verdana, Arial, Helvetica, sans-serif;"><b class="Tit-celdas"><font face="Verdana, Arial, Helvetica, sans-serif">Resultado 
						  de Inscripci&oacute;n de : &nbsp;&nbsp;&nbsp;</font></b></span><span class="Tit-celdas"><b class="Tit-celdas"></b><font class="Tit-celdas"><%=GetNombrealumno(Codcli)%> </font></span></td>
				      </tr>
					  <tr valign="bottom"> 
						<td width="727" class="text-normal-celdas"><font class="text-normal-celdas">Fecha:&nbsp;&nbsp;<%=date()%></font></span></b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="tex-totales-celda"><font class="text-normal-celdas">Hora:&nbsp;&nbsp;<%=time()%></font></span></b></span></td>
						
					  <tr valign="top"> 
						<td colspan="3" height="80"> 
						  <table width="960" cellspacing="1" cellpadding="0" height="59" border="0" bordercolor="#FFFFFF">
							<td height="20" colspan="11"> <div align="left"> 
								<p class="text-normal-celdas"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">Ud. 
								  Tiene inscritas las siguientes asignaturas:</font></p>
							  </div></td>
							</tr>
							<tr background="imagenes/fdo-cabecera-cel22.jpg"> 
							  <td width="63" height="30"  background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center" ><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">C&oacute;digo</font></b></font></div></td>
							  <td width="77" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Asignatura Malla</font></b></font></div></td>
							  <td width="50" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Secci&oacute;n</font></b></font></div></td>
                              <%if MUESTRA_OPORTUNIDAD ="SI" then%>
                              <td width="62" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Oportunidad</font></b></font></div></td>
                              <%end if%>
							  <td width="104" height="30" background="imagenes/fdo-cabecera-cel22.jpg" NOWRAP> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Horario (M&oacute;dulo)</font></b></font></div></td>
                              <td width="80" height="30" background="imagenes/fdo-cabecera-cel22.jpg" NOWRAP> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Sala</font></b></font></div></td>
							  <td width="67" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Asignatura Dictada</font></b></font></div></td>
							  <td width="50" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Tipo</font></b></font></div></td>
							  <td width="148" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Nombre Profesor </font></strong></div></td>
                              <td width="68" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Fecha Inicio</font></b></font></div></td>
							  <td width="37" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Fecha Fin</font></b></font></div></td>
                              <%if MUESTRA_CREDITOS_PA ="SI" then%>
                              <td width="50" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Cr&eacute;ditos</font></b></font></div></td>
                              <%end if%>
							</tr>
							<% 
						  If RamosDebe.Eof Then
						  %>
							<%
						  Else
						   While Not RamosDebe.Eof

							 if Ucase(RamosDebe("inscrito")) = "S" then
							   CodSecc = RamosDebe("CodSecc")
				
      
							   if USAVISTAXPROFERAMORESUMENPA ="SI" then
							   StrSqlH = "Select dbo.HORARIO_RAMO_SECCION_PROFE_ano_periodo('" &RamosDebe("RamoEquiv") & "','" &RamosDebe("CodSecc") & "','" &RamosDebe("codprof") & "','" &Ano & "','" &Periodo & "') as Horario"
							   else
							   StrSqlH = "Select dbo.HORARIO_RAMO_SECCION_SEDE('" &RamosDebe("RamoEquiv") & "','" &RamosDebe("CodSecc") & "','" &Ano & "','" &Periodo & "','" &RamosDebe("codsede") & "') as Horario" 
							   end if
							   	'response.Write(StrSqlH)

							   if bcl_ado(StrSqlH, SqlH) then 
							   	Horario = SqlH("Horario")
							   end if
							 else
							   CodSecc = ""
							   Horario = ""
							 end if
							 
							 Sala = ""
							 if USAVISTAXPROFERAMORESUMENPA ="SI" then
							 	StrSala = "Select dbo.fn_ra_obtiene_sala_sede('" &RamosDebe("RamoEquiv") & "','" &RamosDebe("CodSecc") & "','" &Ano & "','" &Periodo & "','" &RamosDebe("codprof") & "','" & RamosDebe("codsede") & "') as sala" 
							 else
							 	StrSala = "Select dbo.fn_ra_obtiene_sala_sede('" &RamosDebe("RamoEquiv") & "','" &RamosDebe("CodSecc") & "','" &Ano & "','" &Periodo & "',null,'" & RamosDebe("codsede") & "') as sala" 
							 end if 
							 							 
							 if bcl_ado(StrSala, rstSala) then 
							   	Sala = rstSala("sala")
							 end if
							 
							if isnull(RamosDebe("fechainicial")) then
								fechainicio=fechaCalendarioinicio
							else
								fechainicio=RamosDebe("fechainicial")
							end if 
							if isnull(RamosDebe("fechafinal")) then
								fechafin=fechaCalendariofin
							else
								fechafin=RamosDebe("fechafinal")
							end if 
							 
						   %>
							<tr bgcolor="#DBECF2" height="25"> 
							  <td width="63" height="30" align="center" ><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("CodRamo")%></font></td>
							  <td width="77" height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("Nombre")%></font></td>
							  <td width="50" height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=CodSecc%></font></td>
                   <%if MUESTRA_OPORTUNIDAD ="SI" then%>             
               <td width="62" height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("intentos")%></font></td>
                  <%end if%>
							  <td width="104" height="30" align="center" nowrap="nowrap"><font size="1" face="Arial, Helvetica, sans-serif"><%=Horario%></font></td>
                              <td width="80" height="30" align="center" nowrap="nowrap"><font size="1" face="Arial, Helvetica, sans-serif"><%=Sala%></font></td>                              
							  <td width="67" height="30" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=RamosDebe("NombreReal")%></font></td>
							  <td width="50" height="30" align="center"><font size="1" face="Arial, Helvetica, sans-serif"> 
								<%if RamosDebe("Tipocurso")="T"  then
																																 response.write("Te&oacute;rico")
																															  end if	 
																															  if  RamosDebe("Tipocurso")="L" then
																																  response.write("Laboratorio")
																															  end if
																															   if RamosDebe("Tipocurso")="P" then
																																  response.write("Practico")
																															   end if
																															   
																															   
													  %>
							  </font></td>
							  <td width="148" height="30" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=RamosDebe("prof")%></font></td>
                              <td width="68" height="30" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=fechainicio%></font></td>
                              <td width="37" height="30" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=fechafin%></font></td>
							 <%if MUESTRA_CREDITOS_PA ="SI" then%>             
               <td width="50" height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("credito")%></font></td>
               					               		
                  <%TOTALCREDITOS = TOTALCREDITOS + RamosDebe("credito")   
				  end if%>
							</tr>
							<%
							 RamosDebe.MoveNext
						   Wend
						 End If %>
                         

                         <%if MUESTRA_CREDITOS_PA="SI" then%> 
                         <tr bgcolor="#DBECF2" height="25"> 
							  <td colspan="<%if MUESTRA_OPORTUNIDAD ="SI" then%>
							  				10
                                            <%else%>
                                            9
                                            <%end if %>"width="63" height="30" align="right" ><font face="Arial, Helvetica, sans-serif" size="1"> TOTAL CR&Eacute;DITOS INSCRITOS : </font></td>
                              <td width="77" height="30" align="center" ><font face="Arial, Helvetica, sans-serif" size="1"><%=TOTALCREDITOS%></font></td>
                         <%end if %>
						  </table>
						  <br>
						  <table width="960" cellspacing="1" cellpadding="0" height="59" border="0" bordercolor="#FFFFFF">
							<td height="17" colspan="6"> 
							  <div align="left"> 
								<p class="text-normal-celdas"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">Ud. 
								  solicit&oacute; las siguientes asignaturas:</font></p>
							  </div>							</td>
							</tr>
							<tr background="Imagenes/fdo-cabecera-cel22.jpg"> 
							  <td width="40" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
								<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">C&oacute;digo</font></b></font></div>							  </td>
							  <td width="166" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
								<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Asignatura Malla</font></b></font></div>							  </td>
							  <td width="49" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
								<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Secci&oacute;n</font></b></font></div>							  </td>
										  <td width="106" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
								<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Horario (M&oacute;dulo)</font></b></font></div>							  </td>
							  <td width="46" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
								<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Estado</font></b></font></div>							  </td>
							  <td width="110" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
								<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Asignatura Dictada</font></b></font></div>							  </td>
							  
							  <%if paso=1 then%>
							   <td width="49" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
								<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Secci&oacute;n Solicitada</font></b></font></div>							  </td>
							  <%end if%>
							  
							  <%if paso=1 then
							  
								  if Ucase(RamosSolici("Estado")) = "A" then %> 
								  <td width="217" height="30"> 
									<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Observaci&oacute;n</font></b></font></div>							  </td>
							  <%end if    
							  end if%>
							</tr>
							<% 
						  If RamosSolici.Eof Then
						  %>
							<%
						  Else
						   While Not RamosSolici.Eof
							 'response.write ("seccion :" & RamosSolici("CodSecc"))
							 if Ucase(RamosSolici("Estado")) = "A" then
							   CodSecc = RamosSolici("CodSecc")
							   'Horario = GetHor2(RamosSolici("RamoEquiv"), CodSede, RamosSolici("CodSecc"), Ano, Periodo)
							     StrSqlH2 = "Select dbo.HORARIO_RAMO_SECCION('" &RamosSolici("RamoEquiv") & "','" &RamosSolici("CodSecc") & "','" &Ano & "','" &Periodo & "') as Horario" 
							   if bcl_ado(StrSqlH2, SqlH2) then 
							   Horario = SqlH2("Horario")
							   end if
							 else
							   CodSecc = ""
							   'CodSecc = RamosSolici("CodSecc")
							   Horario = ""
							 end if
							CodSecc =RamosSolici("CodSecc")
						   %>
						   
							<tr bgcolor="#DBECF2" height="25"> 
							  <td width="40" height="30" align="center"><font face="Verdana" size="1"><%=RamosSolici("CodRamo")%></font></td>
							  <td width="166" height="30" align="center"><font face="Verdana" size="1"><%=RamosSolici("Nombre")%></font></td>
							  <td width="49" height="30" align="center"><font face="Verdana" size="1"><%=CodSecc%></font></td>
							  
							  <td width="106" height="30" align="center"><font face="Verdana" size="1"><%=Horario%></font></td>
							  <td width="46" height="30" align="center"><font face="Verdana" size="1">
							  <%If RamosSolici("Estado")="A" then
																											 response.write("Aprobado")
																										 elseif RamosSolici("Estado")="P" then
																											   response.write("Pendiente")
																										 else
																											  response.write("Rechazado")
																										 end if	   
																									   '=RamosSolici("Estado")%></font></td>																		                                        
																																						 
							  <td width="110" height="30" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=RamosSolici("NombreReal")%></font></td>
							   <%  if paso=1 then %> 
									 <%'if RamosSolici("codsecc")<>RamosSolici("CodSeccori") then %>
									<td width="49" height="30" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=RamosSolici("CodSeccori")%></font></td>
									<td width="217" height="30" align="center"><font size="1" face="Arial, Helvetica, sans-serif">
							  <%if Ucase(RamosSolici("Estado")) = "A" then
																																		if trim(RamosSolici("codsecc"))<>trim(RamosSolici("CodSeccori")) then 
																																			   response.write("Aprobada en otra Sección")
																																		 end if 
																																	 end if
																																		 %></font></td>
							   <%'END IF%>
							   <%END IF%>      
						   </tr>
							<%
							 RamosSolici.MoveNext
						   Wend
						 End If %>
						  </table>
						  <%
X=REQUEST("X")
IF X=1 THEN
	 
	 %>
	
	 <script>
	 alert("Esta sección no corresponde a la asignatura teórica seleccionada")
	 </script>
	
	 <%
END IF


'Dim strRamosDebe, strRamosPuede, Rut, CodCli
Dim RstHorario

CodCli = Session("CodCli")
IF CodCli = "" then
   'Response.Redirect "pagina de login"
   'Response.Write( "Me voy : " + CodCli + " " + session("CodCli") )
end if
'Ano = Session("Ano")

'sql = ""
'sql = sql & " Select distinct top 1 periodo from ra_carga "
'sql = sql & " where codcli ='" & Session("CodCli") & "' "
'sql = sql & " and ano = '" & Ano & "' "
'sql = sql & " order by periodo desc "
'response.Write(sql)

'if bcl_ado(sql, rstsql2) then 
'	Periodo = rstsql2("Periodo")
'else 
'	Periodo = Session("Periodo")
'end if 
'Periodo = Session("Periodo")

CodSede = Session("CodSede")
P = REQUEST("P")
M = Request("M")

Dim MtrHorario(60, 28)
Dim MtrNombreRamo(60, 28) 
Dim MtrHorSt(60, 28)
Dim MtrHorRamo(60, 28)
Dim MtrHorarioRAMO(60, 28)
Dim MtrHorarioCODSECC(60, 28)
dim MtrHorarioDia(60, 28)

If P = "S" then
  ConPreInscrito = true
else
  ConPreInscrito = false
end if
strHorario =""
T = ucase(REQUEST("T"))
If T = "S" then
  strHorario = " SELECT h.codramo, h.codsecc, h.codsal, h.dia, h.codmod, c.codramo as ramomalla " & _
               " FROM ra_horari h, tmpTest c  " & _
               " WHERE c.codcli ='" & CodCli & "' and " & _
               " c.codramo = h.codramo and " & _
               " c.codsecc = h.codsecc and  " & _
               " c.ano = h.ano and  " & _
               " c.periodo = h.periodo and  " & _
               " h.ano = '" & ano & "' and " & _
               " h.periodo = '" & periodo & "' and " & _
               " h.CodSede = '" & Codsede & "' " & _
               " order by h.dia, h.codmod"

  set RstHorario = Session("Conn").execute(strHorario)

  do while not RstHorario.eof
    j = GetDia(RstHorario("DIA"))
    i = RstHorario("CodMod")
    MtrHorario(i, j) = RstHorario("CODRAMO") & "/" & RstHorario("CodSecc") & " " & RstHorario("CodSal")
    'MtrHorario(i, j) =  "/" & RstHorario("CodSecc") & " " & RstHorario("CodSal")
    MtrHorSt(i,j) = "T"
    MtrHorRamo(i, j) = RstHorario("RamoMalla")
    'Response.Write(MtrHorario(i, j))
    RstHorario.movenext
  loop

end if

S = ucase(REQUEST("S"))
If S = "S" then
  strHorario = " SELECT h.codramo, h.codsecc, h.codsal, h.dia, h.codmod, c.codramo as ramomalla " & _
               " FROM ra_horari h, tmpSolici c  " & _
               " WHERE c.codcli ='" & CodCli & "' and " & _
               " c.ramoequiv = h.codramo and " & _
               " c.codsecc = h.codsecc and  " & _
               " c.ano = h.ano and  " & _
               " c.periodo = h.periodo and  " & _
               " h.ano = '" & ano & "' and " & _
               " h.periodo = '" & periodo & "' and " & _
               " h.CodSede = '" & Codsede & "' " & _
               " order by h.dia, h.codmod"

  'Response.Write(StrHorario)
  set RstHorario = Session("Conn").execute(strHorario)

  do while not RstHorario.eof
    j = GetDia(RstHorario("dia"))
    i = RstHorario("codmod")
    'Response.Write("Dia = " & RstHorario("DIA"))
    'Response.Write("Hora = " & i)
    'Response.Write("Ramo = " & RstHorario("CODRAMO") & "/" & RstHorario("CodSecc") & " " & RstHorario("CodSal"))
    'Response.end
    MtrHorario(i, j) = RstHorario("codramo") & "/" & RstHorario("codsecc") & " " & RstHorario("codsal")
    MtrHorRamo(i, j) = RstHorario("ramomalla")
    if Trim(MtrHorSt(i,j)) = "" then
       MtrHorSt(i,j) = "S"
    else
       MtrHorSt(i,j) = "E"
    end if
    'Response.Write(MtrHorario(i, j))
    RstHorario.movenext
  loop

end if

strParame="SELECT dbo.Fn_ValorParame('MUESTRAHORARIOSEMANALPORTAL')Parame"
if BCL_ADO(strParame, rstParame) then
	if isnull(rstParame("Parame")) then
		HORARIOSEMANAL=""
	else
		HORARIOSEMANAL=trim(rstParame("Parame"))
	end if 
else
	HORARIOSEMANAL=""
end if
' Analizar según el procedimiento TopeHorario
if HORARIOSEMANAL="SI" then
	strHorario = " SELECT h.codramo, h.codsecc, h.codsal, h.dia, h.codmod, 'S' Inscrito, 'S' Preinscrito, c.codramo as ramomalla, ramo.nombre nombreramo "
	strHorario = strHorario & " FROM ra_reservasala h, ra_nota c, ra_ramo ramo  "
	strHorario = strHorario & " WHERE c.codcli ='" & CodCli & "' and ramo.codramo =h.codramo and "
	if ConPreInscrito then
	  strHorario = strHorario & " c.ramoequiv_i = h.codramo and " 
	  strHorario = strHorario & " c.codsecc_i = h.codsecc and  " 
	else
	  strHorario = strHorario & " c.ramoequiv = h.codramo and "
	  strHorario = strHorario & " c.codsecc = h.codsecc and  "
	end if   
	strHorario = strHorario & " coalesce(c.convalidado,'')<>'S' and "
	strHorario = strHorario & " c.ano = h.ano and  "
	strHorario = strHorario & " c.periodo = h.periodo and  "
	strHorario = strHorario & " h.ano = '" & ano & "' and "  
	strHorario = strHorario & " h.periodo = '" & periodo & "' and "
	'strHorario = strHorario & " h.CodSede = '" & Codsede & "' and "
	strHorario = strHorario & " h.CodSede = c.codsede and "	
	strHorario = strHorario & " h.FECHA BETWEEN (SELECT CONVERT(date,DATEADD(wk, DATEDIFF(wk,0,getdate()), 0))) and "
	strHorario = strHorario & " (SELECT CONVERT(date,DATEADD(wk, DATEDIFF(wk,0,getdate()), 0)+6))"					   
	strHorario = strHorario & " UNION SELECT h.codramo, h.codsecc, h.codsal, h.dia, h.codmod, 'S' Inscrito, 'S' Preinscrito, c.codramo as ramomalla, ramo.nombre nombreramo "
	strHorario = strHorario & " FROM ra_reservasala h, ra_notaactividad c, ra_ramo ramo  "  
	strHorario = strHorario & " WHERE c.codcli ='" & CodCli & "' and ramo.codramo =h.codramo and "
	if ConPreInscrito then
	  strHorario = strHorario & " c.ramoequiv_i = h.codramo and "
	  strHorario = strHorario & " c.codsecc_i = h.codsecc and  "
	else
	  strHorario = strHorario & " c.ramoequiv = h.codramo and "
	  strHorario = strHorario & " c.codsecc = h.codsecc and  "
	end if 
	strHorario = strHorario & " coalesce(c.convalidado,'')<>'S' and "
	strHorario = strHorario & " c.ano = h.ano and  "
	strHorario = strHorario & " c.periodo = h.periodo and  "
	strHorario = strHorario & " h.ano = '" & ano & "' and "  
	strHorario = strHorario & " h.periodo = '" & periodo & "' and "
	'strHorario = strHorario & " h.CodSede = '" & Codsede & "' and "
	strHorario = strHorario & " h.CodSede = c.codsede and "	
	strHorario = strHorario & " h.FECHA BETWEEN (SELECT CONVERT(date,DATEADD(wk, DATEDIFF(wk,0,getdate()), 0))) and "
	strHorario = strHorario & " (SELECT CONVERT(date,DATEADD(wk, DATEDIFF(wk,0,getdate()), 0)+6))"					 
	strHorario = strHorario & " order by h.dia, h.codmod"
else
	strHorario = " SELECT h.codramo, h.codsecc, h.codsal, h.dia, h.codmod, 'S' Inscrito, 'S' Preinscrito, c.codramo as ramomalla, ramo.nombre nombreramo "
	strHorario = strHorario & " FROM ra_reservasala h, ra_nota c,ra_ramo ramo  "
	strHorario = strHorario & " WHERE c.codcli ='" & CodCli & "' and ramo.codramo =h.codramo and "
	if ConPreInscrito then
	  strHorario = strHorario & " c.ramoequiv_i = h.codramo and " 
	  strHorario = strHorario & " c.codsecc_i = h.codsecc and  " 
	else
	  strHorario = strHorario & " c.ramoequiv = h.codramo and "
	  strHorario = strHorario & " c.codsecc = h.codsecc and  "
	end if   
	strHorario = strHorario & " coalesce(c.convalidado,'')<>'S' and "
	strHorario = strHorario & " c.ano = h.ano and  "
	strHorario = strHorario & " c.periodo = h.periodo and  "
	strHorario = strHorario & " h.ano = '" & ano & "' and "  
	strHorario = strHorario & " h.periodo = '" & periodo & "' and "
	'strHorario = strHorario & " h.CodSede = '" & Codsede & "' "		
	strHorario = strHorario & " h.CodSede = c.codsede "		
	strHorario = strHorario & " group by h.codramo, h.codsecc, h.codsal, h.dia, h.codmod, c.codramo , ramo.nombre "	   
	strHorario = strHorario & " UNION SELECT h.codramo, h.codsecc, h.codsal, h.dia, h.codmod, 'S' Inscrito, 'S' Preinscrito, c.codramo as ramomalla , ramo.nombre nombreramo " 
	strHorario = strHorario & " FROM ra_reservasala h, ra_notaactividad c,ra_ramo ramo  "
	strHorario = strHorario & " WHERE c.codcli ='" & CodCli & "' and ramo.codramo =h.codramo and "
	if ConPreInscrito then
	  strHorario = strHorario & " c.ramoequiv_i = h.codramo and "
	  strHorario = strHorario & " c.codsecc_i = h.codsecc and  "
	else
	  strHorario = strHorario & " c.ramoequiv = h.codramo and "
	  strHorario = strHorario & " c.codsecc = h.codsecc and  "
	end if 
	strHorario = strHorario & " coalesce(c.convalidado,'')<>'S' and "
	strHorario = strHorario & " c.ano = h.ano and  "
	strHorario = strHorario & " c.periodo = h.periodo and  "
	strHorario = strHorario & " h.ano = '" & ano & "' and "  
	strHorario = strHorario & " h.periodo = '" & periodo & "' and "
	'strHorario = strHorario & " h.CodSede = '" & Codsede & "' "
	strHorario = strHorario & " h.CodSede = c.codsede "	
	strHorario = strHorario & " group by h.codramo, h.codsecc, h.codsal, h.dia, h.codmod, c.codramo , ramo.nombre "	
	strHorario = strHorario & " order by h.dia, h.codmod"
end if 
'RstHorario.Open strHorario, Session("Conn")
'Response.Write(StrHorario)
'Response.End

set RstHorario = Session("Conn").execute(strHorario)

do while not RstHorario.eof
  j = GetDia(RstHorario("dia"))
  i = RstHorario("codmod")
  if ConPreinscrito then
    if RstHorario("Preinscrito") = "S" then
       'Response.Write("Ramo = " & RstHorario("CODRAMO") & " " & rstHorario("RamoMalla") & "<br>" ) 
       MtrHorario(i, j) = RstHorario("CODRAMO") & "/" & RstHorario("CodSecc") & " " & RstHorario("CodSal")
	   MtrNombreRamo(i, j) = RstHorario("nombreramo")
	   MtrHorarioRAMO(i, j) = RstHorario("CODRAMO")
	   MtrHorarioCODSECC(i, j) = RstHorario("CodSecc")	   
	   MtrHorarioDia(i, j) = RstHorario("dia")
       if Trim(MtrHorSt(i,j)) = "" or Trim(MtrHorSt(i,j)) = "T" then
         MtrHorSt(i,j) = "P"
       else
         MtrHorSt(i,j) = "E"
       end if  
       MtrHorRamo(i, j) = RstHorario("RamoMalla")
         
       'MtrHorario(i, j, 2) = "P"
    else
      'if RstHorario("Inscrito") = "S" then
      '   MtrHorario(i, j) = RstHorario("CODRAMO") & "/" & RstHorario("CodSecc") & " " & RstHorario("CodSal")
      '   if Trim(MtrHorSt(i,j)) <> "" then
      '     MtrHorSt(i,j) = "I"
      '   else
      '     MtrHorSt(i,j) = "E"
      '   end if  
         'MtrHorario(i, j, 2) = "I"
      'end if
    end if
  else
    if trim(RstHorario("Inscrito")) = "S" then
       MtrHorario(i, j ) = RstHorario("CODRAMO") & "/" & RstHorario("CodSecc") & " " & RstHorario("CodSal")
	   MtrNombreRamo(i, j) = RstHorario("nombreramo")
   	   MtrHorarioRAMO(i, j) = RstHorario("CODRAMO")
	   MtrHorarioCODSECC(i, j) = RstHorario("CodSecc")	   
	   MtrHorarioDia(i, j) = RstHorario("dia")
       if Trim(MtrHorSt(i,j)) = "" or Trim(MtrHorSt(i,j)) = "T" then
         MtrHorSt(i,j) = "I"
       else
         MtrHorSt(i,j) = "E"
       end if  
       MtrHorRamo(i, j) = RstHorario("RamoMalla")
       'MtrHorario(i, j, 2) = "I"
    end if
  end if
  
  'Response.Write(MtrHorario(i, j))
  RstHorario.movenext
loop

'RstHorario.close()



%>
						  <p></p></td>
					  </tr><tr valign="top"> 
        <td height="1" colspan="2" class="Tit-celdas"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Horario 
          de asignaturas:&nbsp;&nbsp;&nbsp;</font></b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=GetNombrealumno(Codcli)%></font></td> 
        </tr>
      <tr valign="top" > 
        <td colspan="2" height="2"> 
        <table width="960" cellspacing="1" cellpadding="0" height="72" border="0" bordercolor="#FFFFFF" >
          <tr background="Imagenes/fdo-cabecera-cel22.jpg"> 
              <td width="24" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Mod</font></b></font></div>            </td>
              <td width="98" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Hora</font></b></font></div>            </td>
              <td width="94" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Lunes</font></b></font></div>            </td>
              <td width="94" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Martes</font></b></font></div>            </td>
              <td width="94" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Mi&eacute;rcoles</font></b></font></div>            </td>
              <td width="94" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Jueves</font></b></font></div>            </td>
              <td width="89" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Viernes</font></b></font></div>            </td>
              <td width="110" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">S&aacute;bado</font></b></font></div>            </td>
              <td width="110" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Domingo</font></b></font></div>            </td>  
          </tr>
			<%
			dim DiaMaxModulo
			Set RstModulos=Session("Conn").execute("select 'LUNES'X ") 
		if not RstModulos.eof then
			DiaMaxModulo=RstModulos(0)
'			response.Write(DiaMaxModulo)
			RstModulos.close 	
			%>
        <% 'strModulos = "Select Codmod, Hor_Ini, Hor_Fin from ra_modulo where CodSede = '" & CodSede & "' and dia = '"& trim(DiaMaxModulo) &"' Order By convert(numeric,Codmod)"
		strModulos = "select distinct convert(numeric,Codmod)Codmod, Hor_Ini, Hor_Fin from VW_MODULO WHERE CODCARR='"& session("codcarr") &"' and dia = '"& trim(DiaMaxModulo) &"'  and codsede ='"& session("codsede") &"'Order By convert(numeric,Codmod),Hor_Ini, Hor_Fin" 
		'response.Write(strModulos)
           set RstModulos = Session("Conn").execute(strModulos)
           do while not RstModulos.eof
             i = RstModulos("CodMod")
        %>
		  <div id="tooltip" align="center" style="position:absolute;visibility:hidden;border:1px solid black;font-size:12px;layer-background-color:lightyellow;background-color:lightyellow;padding:1px" ></div>
          <tr bgcolor="#DBECF2"> 
             
            <td width="24" height="17" align="center"><font face="Arial" size="1"><%=RstModulos("CodMod")%></font></td>
             <td width="98" height="17" align="center"><font face="Arial" size="1"><%=RstModulos("Hor_Ini") & " - " & RstModulos("Hor_Fin")%></font></td>
             
          <% for j = 1 to 7 %>
             
             <% 
			 if trim(MtrHorSt(i, j)) = "" then 
                  color = "#DBECF2"
             %>
              <td width="94" height="17"><font face="Verdana" size="1">  <br> &nbsp 
             <% 
                else
				
					if MtrHorSt(i,j) ="E" then
				   	   'Analiza si tiene tope de acuerdo a fechas
					   RamoNoTope = MtrHorarioRAMO(i, j)
					   CodsecNoTope = MtrHorarioCODSECC(i, j)
					   ModuloNoTope = RstModulos("CodMod")
					   DiaNoTope = MtrHorarioDia(i, j) 
					   
						strNoTope = "SELECT DISTINCT TOP 1 h.codramo+'/'+CONVERT(VARCHAR,h.codsecc)+'-'+h.codsal Ramo,h.FECHA FROM RA_RESERVASALA h,ra_nota c"
						strNoTope = strNoTope &" WHERE   c.codcli = '"& session("codcli") &"' "
						strNoTope = strNoTope &" AND c.ramoequiv = h.codramo "
						strNoTope = strNoTope &" AND c.codsecc = h.codsecc"
						strNoTope = strNoTope &" AND c.ano = h.ano"
						strNoTope = strNoTope &" AND c.periodo = h.periodo"
						strNoTope = strNoTope &" AND h.ano = '"& session("ano") &"'"
						strNoTope = strNoTope &" AND h.periodo = '"& session("periodo") &"'"
						strNoTope = strNoTope &" AND h.CodSede = '"& session("codsede") &"'"
						strNoTope = strNoTope &" and h.dia ='"& DiaNoTope &"'"
						strNoTope = strNoTope &" AND h.codmod= "& ModuloNoTope &""
						strNoTope = strNoTope &" order by h.FECHA asc"
						
						set RstNoTope = Session("Conn").execute(strNoTope)
						RamoFinal =""
						do while not RstNoTope.eof
							 RamoFinal = RamoFinal & RstNoTope("Ramo") & "<br>"
							 RstNoTope.movenext
						loop
						
					   MtrHorario(i, j) = RamoFinal
					   MtrHorSt(i,j)="I"
				   end if
				   
                   Select case  MtrHorSt(i,j) 
                     Case "I": Color = "#00FF00"
                     Case "E": Color = "#FF0000"
                     Case "P": Color = "#FFCC66"
                     Case "T": Color = "#DBECF2"
                     Case "S": Color = "#FFCC66"
                   End Select 
              %>                
                 <td title="<%=MtrNombreRamo(i, j)%>" bgcolor="<%=color%>" width="94" height="17" ><font face="Verdana" size="1"><%=MtrHorario(i, j)%> 
             <% end if%>
            </font></td>
            
          <% next %>
          </tr>
        <%
           RstModulos.Movenext 
           loop 
	end if 
         %>    
          </table>        </td>
      </tr>
      
      <tr valign="middle"> 
        <td colspan="2" height="10"> 
          <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#666666">(Este 
            documento NO constituye certificado)</font></div>        </td>
      </tr>
      <tr>
      <td align="center">
      
        <img src="Imagenes/Leyenda.bmp"></td>
      </tr>
	</table>
    <table>
           <tr>
    <td width="83" align="left"><A onMouseOver="MM_swapImage('Image7','','Imagenes/botones/volver-on.gif',1)" onmouseout=MM_swapImgRestore() href="javascript: window.top.location.href ='menu_tomaderamos.asp'"><IMG src="Imagenes/botones/volver-of.gif" border="0"  name="Image7"></A></td>
  </tr>
  </table>
	</form></table>
	</td>
  </tr>
</table>


</body>
<script languaje = "javascript">

function ConfirmaModi()
{
	if (confirm("¿Est\u00e1 seguro de modificar su inscripci\u00f3n? "))
	   { 
		   window.location = "modificainscripcion.asp";
	   }
}

function ConfirmaInscrip()
{
  document.form1.action = "confirmacion.asp"
 // alert("Hola");
  document.form1.submit();
} 
function visualiza() {
  var x=window.open ("tex-resumenInscrip.htm","ResumenInscripción","width=500,height=400,resizable=yes,toolbar=yes");
 }

function imprime()
{
  parent.focus();
  parent.print(); 
  //window.print();
  //parent.Horario.focus();
  //parent.Horario.print();
}

</script>
</html>
<%
RamosDebe.Close()
RamosSolici.Close()
%>
<!--#INCLUDE file="include/desconexion.inc" -->
