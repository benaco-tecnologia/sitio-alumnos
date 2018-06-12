<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE file="analytics.asp" -->


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
  response.End()
end if

Dim strRamosDebe, strRamosPuede, Rut, CodCli
CodCli = Session("CodCli")
CodSede = Session("CodSede")

strParame="SELECT coalesce(dbo.Fn_ValorParame('USAANALYTICS'),'NO')Parame"
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
	USAANALYTICS=rstParame("Parame")
end if

if USAANALYTICS="SI" then
	analitycs() 
end if 


if EstaHabilitadaNW (466)="S" then 
	if GetPermisoNW(466) ="N" then
		response.Redirect("MensajesBloqueos.asp")	
	else
		if TomadeRamoED(session("Codcarr"),session("Codcli"),session("ano_ed"),session("periodo_ed"))=0 then
			session("MensajeBloqueosVarios") ="Para acceder a esta opcion debe responder la Encuesta Docente."
			response.Redirect("MensajeBloqueo.asp")
		end if 
		
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
				session("MensajeBloqueosVarios") ="Para acceder a esta opcion debe responder sus Encuestas Pendientes."
				response.Redirect("MensajeBloqueo.asp")
			end if 
		end if
	end if
else
	response.Redirect("MensajeBloqueoHabilita.asp")
end if


strParamePOP="SELECT dbo.Fn_ValorParame('ACTIVATRAUTO')ACTIVATRAUTO"
set rstParamePOP= Session("Conn").Execute(strParamePOP)

if not rstParamePOP.eof then
	ACTIVATRAUTO = rstParamePOP("ACTIVATRAUTO")
else
	ACTIVATRAUTO =""
end if 

Ano = Session("Ano")
Periodo = Session("Periodo")

' desde la tabla mt_parame obtener el año y periodo
strRamosDebe = "SP_LISTADO_RAMOS_PREINSCRITOS_PA '"& codcli &"',"& Ano &","& Periodo & ""
RamosDebe.Open strRamosDebe, Session("Conn")

strAcepto="select acepto FROM SIS_REG_INSCRIPCION WHERE codcli='" & codcli & "' and ano ='" & ano &"' and periodo='" & periodo & "'"
set rstAcepto = session("conn").Execute(strAcepto)
if not rstAcepto.eof then
	AceptoCarga = valnulo(rstAcepto("acepto"),STR_)
end if



strFechaTR="select fecha,fechaAcepto,tipo FROM SIS_REG_INSCRIPCION WHERE codcli='" & codcli & "' and ano ='" & ano &"' and periodo='" & periodo & "'"

set rstFechaTR = session("conn").Execute(strFechaTR)
if not rstFechaTR.eof then
	FechaTR = valnulo(rstFechaTR("fecha"),STR_)
	FechaAcept = valnulo(rstFechaTR("fechaAcepto"),STR_)
	TipoAcept=valnulo(rstFechaTR("tipo"),STR_)
	FueTRA="SI"	
else
	FueTRA="NO"	
end if

strAlumNuevo="select nuevo from mt_alumno WHERE codcli='" & codcli & "'"

set rstAlumNuevo = session("conn").Execute(strAlumNuevo)
if not rstAlumNuevo.eof then
	if ucase(valnulo(rstAlumNuevo("nuevo"),STR_)) ="S" then
		AlumNuevo ="SI"
	else
		AlumNuevo ="NO"
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
<meta http-equiv="Content-Type" content="text/html;">
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
<%' if not InscritoNormal() then %>
  <script>
  //   alert("No ha inscrito aún vía solicitud Normal ...");
  </script>   
<%'  end if%>
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
			<td width="825" valign="top">
			<% CargarTop1()%><% SubMenu()%>	
				<!--<div style="width:825PX; height:480px; overflow:scroll;" align="center">-->
                
				  <form name="form1" method="post" action="asignatura-seccion.asp">
					<input type = "hidden" name = "Ramos" value = "">
					<table border="0" cellpadding="0" cellspacing="15" width="817">
					  <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
					  <tr valign="top" >
						<td ><span style="font-size: 14px"><p style="font-size: 25px" class="text-menu">Resumen de Inscripción</p></span></tr>
						<tr valign="top" > 
						<td  align="right"><a href="javascript:imprime()"><img src="imag/f-der/impresion.gif" width="36" height="36" border="0" align="right"></a></tr>
                       <tr valign="top">
					  	<td height="1" colspan=3><p style="font-size: 13px; font-weight: bold;"><b style="font-weight: bold; color: #0099CC; font-size: 13px;"><font face="Verdana, Arial, Helvetica, sans-serif">Cuadro Informativo:</font></b></p>
				  	     <p  align="justify"style="font-size: 13px; font-weight: bold;"><font face="Verdana, Arial, Helvetica, sans-serif"><span style="font-size: 13px">La prioridad académica es un índice semestral cuyo uso permite ordenar a los alumnos en el proceso de inscripción de asignaturas.
Su cálculo considera el promedio acumulado, las eliminaciones académicas (si las hay), la condición de alumno trabajador y las observaciones disciplinarias.
</span></font></p>
				  	     <p style="font-size: 13px; font-weight: bold;">&nbsp;</p></td>
				      </tr> 	
					  <tr valign="top">
					  	<td height="1" colspan=3><span class="Tit-celdas"><b class="Tit-celdas"></b></span><span style="font-size: 13px; font-weight: bold; font-family: Verdana, Arial, Helvetica, sans-serif;"><b class="Tit-celdas"><font face="Verdana, Arial, Helvetica, sans-serif">Resultado 
						  de Inscripción de : &nbsp;&nbsp;&nbsp;</font></b></span><span class="Tit-celdas"><b class="Tit-celdas"></b><font class="Tit-celdas"><%=GetNombrealumno(Codcli)%> </font></span></td>
				      </tr>
                   <%if FueTRA = "SI" then%>
					  	<tr valign="bottom"> 
						<td width="371" class="text-normal-celdas"><font class="text-normal-celdas">Fecha:&nbsp;&nbsp;<%=left(FechaTR,10)%></font></span></b></td>
                        </tr>
                        					  <tr valign="bottom"> 
						<td width="371" class="text-normal-celdas"><span class="tex-totales-celda"><font class="text-normal-celdas">Hora&nbsp;:&nbsp;&nbsp;&nbsp;<%=Right(FechaTR,8)%></font></span></b></span></td></tr>   

                        <%if fechaAcept <> "" then%>
                        <tr valign="bottom"> 
                            <%if tipoacept="PORTAL" then%>
                            <td width="371" class="text-normal-celdas"><span class="tex-totales-celda"><font class="text-normal-celdas">El Alumno confirm&oacute; la toma de ramos el d&iacute;a &nbsp;<%=left(fechaAcept,10)%> a las <%=Right(fechaAcept,8)%> hrs.</font></span></b></span></td>
							<%else%>
                            <td width="375" class="text-normal-celdas"><span class="tex-totales-celda"><font class="text-normal-celdas">Se confirm&oacute; por Sistema la toma de ramos el d&iacute;a &nbsp;<%=left(fechaAcept,10)%> a las <%=Right(fechaAcept,8)%> hrs.</font></span></b></span></td>
                            <%end if%>
                        </tr>
                         <%
							 GlosaRamos ="Ud. Tiene inscritas las siguientes asignaturas:"
					     else
						 	 GlosaRamos ="Ud. Tiene preinscritas las siguientes asignaturas:"
						 end if %>             
                  <%end if %>
					  <tr valign="top"> 
						<td colspan="3" height="80"> 
						  <table width="776" cellspacing="1" cellpadding="0" height="59" border="0" bordercolor="#FFFFFF">
                      <td height="20" colspan="11"> <div align="left"> 
                          <p class="text-normal-celdas"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000"><%=GlosaRamos%></font></p>
                        </div></td>
                      </tr>
                      <tr background="imagenes/fdo-cabecera-cel22.jpg"> 
                        <td width="61" height="30"  background="imagenes/fdo-cabecera-cel22.jpg"> 
                          <div align="center" ><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">C&oacute;digo</font></b></font></div></td>
                        <td width="166" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                          <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Asignatura</font></b></font></div></td>
                        <td width="42" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                          <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">secci&oacute;n</font></b></font></div></td>
                        <td width="45" height="30" background="imagenes/fdo-cabecera-cel22.jpg"s> 
                          <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Horario</font></b></font></div></td>
                        <td width="80" height="30" background="imagenes/fdo-cabecera-cel22.jpg" NOWRAP> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Sala</font></b></font></div></td>
                        <td width="88" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                          <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Ramo</font></b></font></div></td>
                        <td width="33" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                          <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Tipo</font></b></font></div></td>
                       
                      </tr>
                      <% 
						  If RamosDebe.Eof Then
						  %>
                      <%
						  Else
						  
						   While Not RamosDebe.Eof
						   						    
							 if Ucase(RamosDebe("inscrito")) = "S" or  Ucase(RamosDebe("preinscrito")) = "S"then
							   CodSecc = RamosDebe("CodSecc")
							   'Horario = GetHor2(RamosDebe("RamoEquiv"), CodSede, RamosDebe("CodSecc"), Ano, Periodo)
							   StrSqlH = "Select dbo.HORARIO_RAMO_SECCION_SEDE('" &RamosDebe("RamoEquiv") & "','" &RamosDebe("CodSecc") & "','" &Ano & "','" &Periodo & "','"& session("codsede") &"') as Horario" 
							   if bcl_ado(StrSqlH, SqlH) then 
							   	Horario = SqlH("Horario")
							   end if
							   
							   	Sala = ""
								StrSala = "Select dbo.fn_ra_obtiene_sala('" &RamosDebe("RamoEquiv") & "','" &RamosDebe("CodSecc") & "','" &Ano & "','" &Periodo & "',null) as sala" 
								if bcl_ado(StrSala, rstSala) then 
									Sala = rstSala("sala")
								end if
							 
							 else
							   CodSecc = ""
							   Horario = ""
							 end if
							 RamosEnvio= RamosEnvio & RamosDebe("CodRamo")&";"
						   %>
                      <tr bgcolor="#DBECF2" height="25"> 
                        <td width="61" height="30" align="center" ><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("CodRamo")%></font></td>
                        <td width="166" height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("Nombre")%></font></td>
                        <td width="42" height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=CodSecc%></font></td>
                        <td width="45" height="30" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=Horario%></font></td>
                        <td width="80" height="30" align="center" nowrap="nowrap"><font size="1" face="Arial, Helvetica, sans-serif"><%=Sala%></font></td>
                        <td width="88" height="30" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=RamosDebe("NombreReal")%></font></td>
                        <td width="33" height="30" align="center"><font size="1" face="Arial, Helvetica, sans-serif"> 
                          <%if RamosDebe("Tipocurso")="T"  then
																																 response.write("Teórico")
																															  end if	 
																															  if  RamosDebe("Tipocurso")="L" then
																																  response.write("Laboratorio")
																															  end if
																															   if RamosDebe("Tipocurso")="P" then
																																  response.write("Práctico")
																															   end if
																															   
																															   
													  %>
                          </font></td>
                        
                      </tr>
                      <%
							 RamosDebe.MoveNext
						   Wend
						 End If %>
                    </table>
						  <br>
						<%if 0=1 then%>	
						  <table width="776" cellspacing="1" cellpadding="0" height="59" border="0" bordercolor="#FFFFFF">
							<td height="17" colspan="6"> 
							  <div align="left"> 
								<p class="text-normal-celdas"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">Ud. 
								  solicit&oacute; las siguientes asignaturas:</font></p>
							  </div>							</td>
							</tr>
							<tr background="Imagenes/fdo-cabecera-cel22.jpg"> 
							  <td width="42" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
								<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">C&oacute;digo</font></b></font></div>							  </td>
							  <td width="177" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
								<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Asignatura</font></b></font></div>							  </td>
							  <td width="52" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
								<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">secci&oacute;n</font></b></font></div>							  </td>
										  <td width="66" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
								<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Horario</font></b></font></div>							  </td>
							  <td width="49" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
								<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Situación</font></b></font></div>							  </td>
							  <td width="116" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
								<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Ramo</font></b></font></div>							  </td>
							  
							  <%if paso=1 then%>
							   <td width="52" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
								<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">secci&oacute;n Solicitada</font></b></font></div>							  </td>
							  <%end if%>
							  
							  <%if paso=1 then
							  
								  if Ucase(RamosSolici("Estado")) = "A" then %> 
								  <td width="188" height="30"> 
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
							   Horario = GetHor2(RamosSolici("RamoEquiv"), CodSede, RamosSolici("CodSecc"), Ano, Periodo)
							 else
							   CodSecc = ""
							   'CodSecc = RamosSolici("CodSecc")
							   Horario = ""
							 end if
							CodSecc =RamosSolici("CodSecc")
						   %>
						   
							<tr bgcolor="#DBECF2" height="25"> 
							  <td width="42" height="30" align="center"><font face="Verdana" size="1"><%=RamosSolici("CodRamo")%></font></td>
							  <td width="177" height="30" align="center"><font face="Verdana" size="1"><%=RamosSolici("Nombre")%></font></td>
							  <td width="52" height="30" align="center"><font face="Verdana" size="1"><%=CodSecc%></font></td>
							  
							  <td width="66" height="30" align="center"><font face="Verdana" size="1"><%=Horario%></font></td>
							  <td width="49" height="30" align="center"><font face="Verdana" size="1">
							  <%If RamosSolici("Estado")="A" then
																											 response.write("Aprobado")
																										 elseif RamosSolici("Estado")="P" then
																											   response.write("Pendiente")
																										 else
																											  response.write("Rechazado")
																										 end if	   
																									   '=RamosSolici("Estado")%></font></td>																		                                        
																																						 
							  <td width="116" height="30" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=RamosSolici("NombreReal")%></font></td>
							   <%  if paso=1 then %> 
									 <%'if RamosSolici("codsecc")<>RamosSolici("CodSeccori") then %>
									<td width="52" height="30" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=RamosSolici("CodSeccori")%></font></td>
									<td width="188" height="30" align="center"><font size="1" face="Arial, Helvetica, sans-serif">
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
	<%end if %>
	 <%
END IF


'Dim strRamosDebe, strRamosPuede, Rut, CodCli
Dim RstHorario

CodCli = Session("CodCli")
IF CodCli = "" then
   'Response.Redirect "pagina de login"
   'Response.Write( "Me voy : " + CodCli + " " + session("CodCli") )
end if
Ano = Session("Ano")
Periodo = Session("Periodo")
CodSede = Session("CodSede")
P = REQUEST("P")
M = Request("M")

Dim MtrHorario(25, 7)
Dim MtrHorSt(25, 7)
Dim MtrHorRamo(25, 7)

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



' Analizar según el procedimiento TopeHorario
if ConPreInscrito then
  strHorario = " SELECT h.codramo, h.codsecc, h.codsal, h.dia, h.codmod, c.Inscrito, c.Preinscrito, c.codramo as ramomalla "
  strHorario = strHorario & " FROM ra_horari h, ra_carga c  "
  strHorario = strHorario & " WHERE c.codcli ='" & CodCli & "' and " 
  strHorario = strHorario & " c.ramoequiv_i = h.codramo and " 
  strHorario = strHorario & " c.codsecc_i = h.codsecc and  " 
  strHorario = strHorario & " c.ano = h.ano and  " 
  strHorario = strHorario & " c.periodo = h.periodo and  "
  strHorario = strHorario & " h.ano = '" & ano & "' and "
  strHorario = strHorario & " h.periodo = '" & periodo & "' and "
  strHorario = strHorario & " h.CodSede = '" & Codsede & "' "
			   
  strHorario = strHorario & " UNION SELECT h.codramo, h.codsecc, h.codsal, h.dia, h.codmod, c.Inscrito, c.Preinscrito, c.codramo as ramomalla "
  strHorario = strHorario & " FROM ra_horari h, ra_cargaactividad c  "
  strHorario = strHorario & " WHERE c.codcli ='" & CodCli & "' and "
  strHorario = strHorario & " c.ramoequiv_i = h.codramo and "
  strHorario = strHorario & " c.codsecc_i = h.codsecc and  "
  strHorario = strHorario & " c.ano = h.ano and  "
  strHorario = strHorario & " c.periodo = h.periodo and  "
  strHorario = strHorario & " h.ano = '" & ano & "' and "
  strHorario = strHorario & " h.periodo = '" & periodo & "' and "
  strHorario = strHorario & " h.CodSede = '" & Codsede & "' "
  strHorario = strHorario & " order by h.dia, h.codmod"
  
else
  strHorario = " SELECT h.codramo, h.codsecc, h.codsal, h.dia, h.codmod, c.Inscrito, c.Preinscrito, c.codramo as ramomalla "
  strHorario = strHorario & " FROM ra_horari h, ra_carga c  "
  strHorario = strHorario & " WHERE c.codcli ='" & CodCli & "' and "
  strHorario = strHorario & " c.ramoequiv = h.codramo and "
  strHorario = strHorario & " c.codsecc = h.codsecc and  "
  strHorario = strHorario & " c.ano = h.ano and  "
  strHorario = strHorario & " c.periodo = h.periodo and  "
  strHorario = strHorario & " h.ano = '" & ano & "' and "
  strHorario = strHorario & " c.preinscrito = 'S' and "
  strHorario = strHorario & " h.periodo = '" & periodo & "' and "
  strHorario = strHorario & " h.CodSede = '" & Codsede & "' "
			   
  strHorario = strHorario & " UNION SELECT h.codramo, h.codsecc, h.codsal, h.dia, h.codmod, c.Inscrito, c.Preinscrito, c.codramo as ramomalla "
  strHorario = strHorario & " FROM ra_horari h, ra_cargaactividad c  "
  strHorario = strHorario & " WHERE c.codcli ='" & CodCli & "' and "
  strHorario = strHorario & " c.ramoequiv = h.codramo and "
  strHorario = strHorario & " c.codsecc = h.codsecc and  "
  strHorario = strHorario & " c.ano = h.ano and  "
  strHorario = strHorario & " c.periodo = h.periodo and  "
  strHorario = strHorario & " h.ano = '" & ano & "' and "
  strHorario = strHorario & " c.preinscrito = 'S' and "
  strHorario = strHorario & " h.periodo = '" & periodo & "' and "
  strHorario = strHorario & " h.CodSede = '" & Codsede & "' "
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
       if Trim(MtrHorSt(i,j)) = "" or Trim(MtrHorSt(i,j)) = "T" then
         MtrHorSt(i,j) = "P"
       else
         MtrHorSt(i,j) = "E"
       end if  
       MtrHorRamo(i, j) = RstHorario("RamoMalla")
         
    end if
  else
    if trim(RstHorario("Preinscrito")) = "S" then
       MtrHorario(i, j ) = RstHorario("CODRAMO") & "/" & RstHorario("CodSecc") & " " & RstHorario("CodSal")
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
        <table width="777" cellspacing="1" cellpadding="0" height="72" border="0" bordercolor="#FFFFFF" >
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
          </tr>
			<%
			dim DiaMaxModulo
			Set RstModulos=Session("Conn").execute("select top 1 dia from ra_modulo where dia='LUNES' and codmod in (select max(convert(int,codmod)) as codmod from ra_modulo )")
			DiaMaxModulo=RstModulos(0)
'			response.Write(DiaMaxModulo)
			RstModulos.close 	
			%>
        <% strModulos = "Select Codmod, Hor_Ini, Hor_Fin from ra_modulo where CodSede = '" & CodSede & "' and dia = '"& trim(DiaMaxModulo) &"' Order By convert(numeric,Codmod)" 
           set RstModulos = Session("Conn").execute(strModulos)
           do while not RstModulos.eof
             i = RstModulos("CodMod")
        %>
		  <div id="tooltip" align="center" style="position:absolute;visibility:hidden;border:1px solid black;font-size:12px;layer-background-color:lightyellow;background-color:lightyellow;padding:1px" ></div>
          <tr bgcolor="#DBECF2"> 
             
            <td width="24" height="17" align="center"><font face="Arial" size="1"><%=RstModulos("CodMod")%></font></td>
             <td width="98" height="17" align="center"><font face="Arial" size="1"><%=RstModulos("Hor_Ini") & " - " & RstModulos("Hor_Fin")%></font></td>
          <% for j = 1 to 6 %>
             
             <% if trim(MtrHorSt(i, j)) = "" then 
                  color = "#DBECF2"
             %>
              <td width="94" height="17"><font face="Verdana" size="1">  <br> &nbsp 
             <% 
                else
                   Select case  MtrHorSt(i,j) 
                     Case "I": Color = "#00FF00"
                     Case "E": Color = "#FF0000"
                     Case "P": Color = "#FFCC66"
                     Case "T": Color = "#DBECF2"
                     Case "S": Color = "#FFCC66"
                   End Select
              %>                
                 <td bgcolor="<%=color%>" width="94" height="17" onMouseOver="javascript:showtip(this, event, '<%=MtrHorRamo(i, j)%>')" onMouseout="javascript:hidetip()"><font face="Verdana" size="1"><%=MtrHorario(i, j) %> 
             <% end if%>
            </font></td>
          <% next %>
          </tr>
        <%
           RstModulos.Movenext 
           loop 
         %>    
          </table>        </td>
          	<%if  ACTIVATRAUTO="SI" then%>
		  <%if AceptoCarga<>"SI"  and AlumNuevo <>"SI" then%>    
          </tr>
          <tr valign="middle"> 
           <td align="center">
           <a target="_top"  href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('aceptacarga','','Imagenes/aceptacargaon.jpg',1)" onClick="javascript:confirma();" ><img src="Imagenes/aceptacarga.jpg" name="aceptacarga"  id="aceptacarga"></a>
           <a target="_top"  href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('RevisaSecciones','','Imagenes/RevisaSeccioneson.jpg',1)" onClick="javascript:confirma2();" ><img src="Imagenes/RevisaSecciones.jpg" name="RevisaSecciones"  id="RevisaSecciones" onClick="" ></a>
           </td>
    
          </tr>
          <%end if %>    
                    <%end if %>    
      <tr valign="middle"> 
        <td colspan="2" height="10"> 
        
          <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#666666">(Este 
            documento NO constituye certificado)</font></div>        </td>
      </tr>
	</table>
	</form></table>
	</td>
  </tr>
</table>

</body>
<script languaje = "javascript">
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
function confirma()
{
	if (confirm("Una vez confirmada la carga de propuesta, su horario será definitivo."))
	{
		location.href="insertaCargaConfirmada.asp"
	}			
}
function confirma2()
{
	alert("Usted aún no ha confirmado la propuesta de asignaturas para este período. Si no confirma, corre riesgo de no tener carga académica para el semestre.");	
	location.href="asignatura-seccionTRA.asp?Ramos=<%=RamosEnvio%>"
}

</script>
</html>
<%
RamosDebe.Close()
%>
<!--#INCLUDE file="include/desconexion.inc" -->
