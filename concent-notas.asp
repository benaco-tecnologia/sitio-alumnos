<!-- saved from url=(0022)http://internet.e-mail -->
<!--<OBJECT RUNAT=server PROGID=ADODB.Recordset id=malla name=malla></OBJECT> -->
<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
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
<!--<OBJECT RUNAT=server PROGID=ADODB.Recordset id=Detalle name=Detalle > </OBJECT>-->
<%

dim strsql,strsql2,strMax,strn,DN
codcli=session("codcli")
codcarr = session("codcarr")
'codcli=200011001
'response.Write(codcli)
'response.end

 
strMax = "SELECT TOP 1 COUNT(codramo) AS maximo FROM ra_Acteval_Nota_det WHERE CodCLi ='" & (codcli) & "' and linea not in (99,999)"
strMax = strMax &" GROUP BY codramo ORDER BY maximo desc"

set Rsm = Session("Conn").Execute(strMax)


'strsql="select distinct case when coalesce(s.tipocurso,'')='T' then 'T&eacute;orico' when coalesce(s.tipocurso,'')='P' then 'Pr&aacute;ctico' when coalesce(s.tipocurso,'')='L' then 'Laboratorio' else 'T&eacute;orico' end tipocurso,a.codsecc,a.codramo AS CODRAMO,a.ramoequiv,c.nombre,a.periodo,a.ano,a.np,CASE WHEN ((a.nf = 0 OR a.nf IS null) AND a.ESTADO IN ( 'E', 'I' )) THEN a.CONCEPTO ELSE CONVERT(VARCHAR,a.nf) end nf,a.escala ,a.nej,a.ncert1,a.ncert2,a.ncert3,a.ncert4,a.ncert5,a.ncert6,a.ncert7,a.ncert8,a.ncert9,a.ncert10,a.asistencia,a.ne,a.ner,a.nfcontrol"
'strsql=strsql & " from ra_nota a, ra_ramo c, ra_Seccio s "
'strsql=strsql & " where a.codcli='" & (codcli) & "' and "
'strsql=strsql & " a.codramo=c.codramo " 
'strsql=strsql & " AND a.codramo = c.codramo"        
'strsql=strsql & " AND s.codramo =* a.ramoequiv"
'strsql=strsql & " AND s.codsecc = a.codsecc"
'strsql=strsql & " AND s.ano = a.ano"
'strsql=strsql & " AND s.periodo = a.periodo"
'strsql=strsql & " AND s.codsecc = a.codsecc"
'strsql=strsql & " AND s.CODSEDE = a.CODSEDE"
'strsql=strsql & " group by s.tipocurso,a.codsecc,a.codramo,a.ramoequiv,c.nombre,a.periodo,a.ano,a.np,a.nf,a.escala,a.nej,a.ncert1,a.ncert2,a.ncert3,a.ncert4,a.ncert5,a.ncert6,a.ncert7,a.ncert8,a.ncert9,a.ncert10,a.asistencia,a.ne,ner,a.nfcontrol,a.CONCEPTO,a.ESTADO"
'strsql= strsql & " union select distinct case when coalesce(s.tipocurso,'')='T' then 'T&eacute;orico' when coalesce(s.tipocurso,'')='P' then 'Pr&aacute;ctico' when coalesce(s.tipocurso,'')='L' then 'Laboratorio' else 'T&eacute;orico' end tipocurso,a.codsecc,a.codramo AS CODRAMO,a.ramoequiv,c.nombre,a.periodo,a.ano,a.np,CASE WHEN ((a.nf = 0 OR a.nf IS null) AND a.ESTADO IN ( 'E', 'I' )) THEN a.CONCEPTO ELSE CONVERT(VARCHAR,a.nf) end nf,'' ,a.nej,a.ncert1,a.ncert2,a.ncert3,a.ncert4,a.ncert5,a.ncert6,a.ncert7,a.ncert8,a.ncert9,a.ncert10,a.asistencia,a.ne,a.ner,a.nfcontrol"
'strsql=strsql & " from ra_notaactividad a, ra_ramo c, ra_Seccio s "
'strsql=strsql & " where a.codcli='" & (codcli) & "' and "
'strsql=strsql & " a.codramo=c.codramo " 
'strsql=strsql & " AND a.codramo = c.codramo"        
'strsql=strsql & " AND s.codramo =* a.ramoequiv"
'strsql=strsql & " AND s.codsecc = a.codsecc"
'strsql=strsql & " AND s.ano = a.ano"
'strsql=strsql & " AND s.periodo = a.periodo"
'strsql=strsql & " AND s.codsecc = a.codsecc"
'strsql=strsql & " AND s.CODSEDE = a.CODSEDE"
'strsql=strsql & " group by s.tipocurso,a.codsecc,a.codramo,a.ramoequiv,c.nombre,a.periodo,a.ano,a.np,a.nf,a.nej,a.ncert1,a.ncert2,a.ncert3,a.ncert4,a.ncert5,a.ncert6,a.ncert7,a.ncert8,a.ncert9,a.ncert10,a.asistencia,a.ne,ner,a.nfcontrol,a.CONCEPTO,a.ESTADO"
'strsql= strsql & " order by a.ano, a.periodo  "

strsql = "SP_LISTA_RAMOS_CONCENT_PA '"& codcli &"'"
'response.Write(StrSql)

Set Rst = Session("Conn").Execute(StrSql)

%>

<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if

if EstaHabilitadaNW (477)="S" then 
	if GetPermisoNW(477) ="N" then
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
 
'MEJORA BLOQUEA ENCUESTASEK
strPar="SELECT dbo.Fn_ValorParame('VALIDAENCUESTASEK')Parame" 
set rstPar= Session("Conn").Execute(strPar)		
if not rstPar.eof then
		VALIDAENCUESTASEK=rstPar("Parame")
	else
		VALIDAENCUESTASEK="" 
end if

if VALIDAENCUESTASEK = "SI" then
	strAl="SP_VALIDAENCUESTA_PA_SEK '"& session("codcli") &"' "
	set rstAl= Session("Conn").Execute(strAl)		
	if not rstAl.eof then
		BloqueaTRRep=rstAl(0) 
	end if
	
	if BloqueaTRRep="NO" then  
		session("MensajeBloqueosVarios") ="Para proceder a ver sus calificaciones de sus asignaturas debe realizar todas las encuestas dispuestas en su portal. </br> Si presentas algún inconveniente en la realización de la encuesta, por favor comunícate con el departamento de Informática de la Universidad, o envíanos un mail a <a href='mailto:soporte.uisek@usek.cl'>soporte.uisek@usek.cl</a>"	
		
		response.Redirect("MensajeBloqueo.asp")
	end if
end if 
'FIN MEJORA

strOculta="SELECT coalesce(dbo.Fn_ValorParame('OCULTACOLUMNASCONTROLCONCENTPA'),'NO')Parame" 
set rstOculta= Session("Conn").Execute(strOculta)		
if not rstOculta.eof then
		OCULTACOLUMNASCONTROLCONCENTPA=rstOculta("Parame")
	else
		OCULTACOLUMNASCONTROLCONCENTPA="NO" 
end if



Audita 477,"Ingresa a Concentracion de notas"

strParame="SELECT dbo.Fn_ValorParame('ACTIVACONNOTASDETALLEPORTAL')Parame,dbo.Fn_ValorParame('USACONCENTNOTASPA_CENFOTUR')USACONCENTNOTASPA_CENFOTUR"
if BCL_ADO(strParame, rstParame) then
	ConcentracionDetalle=rstParame("Parame")
	USACONCENTNOTASPA_CENFOTUR = rstParame("USACONCENTNOTASPA_CENFOTUR")
else
	ConcentracionDetalle=""
	USACONCENTNOTASPA_CENFOTUR=""
end if 


%>

<html>
<head>
<title><%=Session("NombrePestana")%></title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=charset=iso-8859-1"> 
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
<link rel="stylesheet" href="file://///Gotzilla/uddweb/admalumnos/css/tex-normales.css" type="text/css">
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<link href="css/tex-normales.css" rel="stylesheet" type="text/css">
<body >
<table border="0" cellpadding="0" cellspacing="0"  align="left">
  <tr>
	<td>
		<table  border="0" cellpadding="0" cellspacing="0" align="center">
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
			<div>
			  <table border="0" cellpadding="15" cellspacing="0" height="187" align="center" width="1000">
				<!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
				<tr> 
                <%if USACONCENTNOTASPA_CENFOTUR = "SI" then%>                
               	 <td colspan="2"><p><span style="font-size: 14px"><p id ="lblTituloInscripcion" style="font-size: 25px" class="text-menu">
                    Hist&oacute;rico de notas</p></span></p></td>
                <%else%>
              	  <td colspan="2"><img src="Imagenes/titulos/T-concentracion_notas.gif" width="300" height="38"></td>
                <%end if%>
				</tr>
				<tr valign="bottom"> 
				  <td colspan="2" height="20" class="tex-normales"><font id="lblSubtitAsig" color="#333333" size="2" face="Arial, Helvetica, sans-serif" class="text-normal-celdas">Es 
					el listado de tus asignaturas cursadas a la fecha.</font></td>
				</tr> 
                <%strParame="SELECT dbo.Fn_ValorParame('ACTIVACONNOTASDETALLEPORTAL')Parame"
					if BCL_ADO(strParame, rstParame) then
						ConcentracionDetalle=rstParame("Parame")
					else
						ConcentracionDetalle=""
					end if 
					  
					if ConcentracionDetalle="NO" then%>
				<tr valign="top"> 
				  <td colspan="2" height="70"> <table  cellspacing="1" cellpadding="0" border="0" bordercolor="#FFFFFF">
					  <tr background="imagenes/fdo-cabecera-cel22.jpg"> 
						<td width="81" height="30"background="imagenes/fdo-cabecera-cel22.jpg" > <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font id="lblCodRamo" size="1">C&oacute;digo 
						Ramo </font></b></font></div></td>
						<td width="163" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font id="lblNombreRamo" size="1">Nombre 
						del Ramo</font></b></font></div></td>
                        <td width="94" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Secci&oacute;n</font></b></font></div></td>
						<td width="40" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font id="lblPeriodo" size="1">Per&iacute;odo</font></b></font></div></td>
						<td width="48" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">A&ntilde;o</font></b></font></div></td>
						<td width="61" height="30" background="imagenes/fdo-cabecera-cel22.jpg">               <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Asistencia</font></b></font>
						</div>
                        <%if OCULTACOLUMNASCONTROLCONCENTPA = "NO" then%>
						<td width="39" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>N.Ctrl.</b></font></div></td>
						<td width="38" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">N.Ej</font></b></font></div></td>
                        <%end if%>
                        
						<% dim i
						if Rsm.eof and Rsm.bof then 
						
						else
						 For i = 1 To Rsm("maximo") Step 1 %>
                         <td width="43" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">P<%=i%></font></b></font></div></td>
                         <% Next  
						 end if
						 %>
                        
                     <!--   <td width="41" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">P1</font></b></font></div></td>
						<td width="41" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">P2</font></b></font></div></td>
						<td width="41" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">P3</font></b></font></div></td>
						<td width="41" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">P4</font></b></font></div></td>
						<td width="41" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">P5</font></b></font></div></td>
						<td width="41" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">P6</font></b></font></div></td>
						<td width="41" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">P7</font></b></font></div></td>
						<td width="41" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">P8</font></b></font></div></td>
						<td width="41" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">P9</font></b></font></div></td>
						<td width="41" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">P10</font></b></font></div></td> -->
						<td width="48" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Nota Pres.</font></b></font></div></td>
						<td width="52" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Nota 
						Ex.</font></b></font></div></td>
						<td width="56" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Nota 
						Ex.Rep.</font></b></font></div></td>
						<td width="75" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Nota Final</font></b></font></div></td>
                        <%if OCULTACOLUMNASCONTROLCONCENTPA = "NO" then%>
						<td width="54" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Escala</font></b></font></div></td>
                        <%end if%>
					  </tr>
					  <%do while not rst.eof %>
					  <%
						if not rst.eof then
						codigoramo=rst("ramoequiv")

						tipocurso = rst("tipocurso") 						
						Nombre=rst("Nombre")						
						seccion =rst("codsecc")
						periodo=rst("periodo")
						ano=rst("ano")
						notacont=rst("nfcontrol")
						notaejer=rst("nej")
						ncert1=rst("ncert1")
						ncert2=rst("ncert2")
						ncert3=rst("ncert3")
						ncert4=rst("ncert4")  
						ncert5=rst("ncert5")  			
						Escala=rst("escala")  
						ncert6=rst("ncert6")  
						ncert7=rst("ncert7")
						ncert8=rst("ncert8")
						ncert9=rst("ncert9")
						ncert10=rst("ncert10")
						notaPres=  replace(valnulo(rst("np"),NUM_), ",", ".")
						notaEx= replace(valnulo(rst("ne"),NUM_), ",", ".") 
						notaExRep=  replace(valnulo(rst("ner"),NUM_), ",", ".")
						notafinal=  rst("nf") 
						Asistencia=rst("asistencia")  
			  
						end if
						
						strn = "SELECT Nota FROM ra_Acteval_Nota_det WHERE codcli = '"& codcli &"' AND codramo = '"&codigoramo&"' and linea not in (99,999) and ano="& ano &" and periodo="& periodo &""	
						'response.Write(strn)
						set RsNotas =  Session("Conn").Execute(strn)
						%>
					  <tr bgcolor="#DBECF2" class="text-normal-celdas"> 
						<td width="81" height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=codigoramo%></font></div></td>
						<td width="163" height="32" bgcolor="#DBECF2"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=Nombre%> 
					    </font></div></td>
                        <td width="94" height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=seccion & " - " & tipocurso%></font></div></td>
						<td width="40" height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=periodo%></font></div></td>
						<td width="48" height="32" bgcolor="#DBECF2"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=ano%></font></div></td>
						<td width="61" height="32"><div align="center">
						  <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=Asistencia%></font></div>
						</div></td>
                        <%if OCULTACOLUMNASCONTROLCONCENTPA = "NO" then%>
						<td width="39" height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=nfcontrol%></font></div></td>
						<td width="38" height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=nej%></font></div></td>
                        <%end if%>
                        <%
						if  Rsm.eof and Rsm.bof then
						
						else
						
						 for i = 1 To Rsm("maximo") Step 1 
								 if not RsNotas.EOF then
									if NOT IsNull(RsNotas("Nota")) then
									DN= RsNotas("Nota")
									else
									DN = ""
									end if
									RsNotas. movenext
								 else
								 DN = ""
								 end if 
								 DN=replace(DN, ",", ".")
						%>   
                                             
                        <td width="43" height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=DN%></font></div></td>
                         <% 	
						 					
						 Next 
						 
						 end if 
						 %>
						<!--
                        <td width="41" height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1">%=ncert1%</font></div></td>                        
						<td width="41"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1">%=ncert2% </font></div></td>
						<td width="41"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1">%=ncert3% </font></div></td>
						<td width="41"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1">%=ncert4% </font></div></td>
						<td width="41"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1">%=ncert5% </font></div></td>
						<td width="41"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1">%=ncert6% </font></div></td>
						<td width="41"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1">%=ncert7% </font></div></td>
						<td width="41"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1">%=ncert8% </font></div></td>
						<td width="41"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1">%=ncert9% </font></div></td>
						<td width="41"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1">%=ncert10% </font></div></td> -->
                        
						<td width="48"><div align="center"><font face="Arial, Helvetica, sans-serif" color="#2E31B8" size="1"><b><%=notaPres%></b></font></div></td>
						<td width="52" height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" color="#2E31B8" size="1"><b><%=notaEx%></b></font></div></td>
						<td width="56" height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" color="#2E31B8" size="1"><b><%=notaExRep%></b></font></div></td>
						<td width="75"><div align="center"><font face="Arial, Helvetica, sans-serif" color="#2E31B8" size="1"><b><%=notafinal%></b></font></div></td>
                        <%if OCULTACOLUMNASCONTROLCONCENTPA = "NO" then%>
						<td width="54" height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" color="#2E31B8" size="1"><b><%=Escala%></b></font></div></td>
                        <%end if%>
                        
						<%  rst.movenext %>
					  </tr>
					  <% loop%>
				  </table></td>
				</tr>
                <%else 'nueva concentracion%>
                   <input type="hidden" id="lblCodRamo" value="">
                   <input type="hidden" id="lblNombreRamo" value="">
                   <input type="hidden" id="lblPeriodo" value="">

                <tr > 
				  <td colspan="2"> <table  cellspacing="1" cellpadding="0" border="0" bordercolor="#FFFFFF" width="1045" align="center">
                
					<%
					if Rst.eof and Rst.bof then 
					
					else			
	                    Rst.movefirst			
					end if 	
					
                    while not Rst.eof
                    
                    strProc ="PA08_INGRESOCALIFICACIONES_SEL '"&valnulo(Rst("Ano"),STR_)&"','"&valnulo(Rst("Periodo"),STR_)&"','" &session("Codsede")&"','','"& valnulo(Rst("ramoequiv"),STR_) &"','"&valnulo(Rst("Codsecc"),STR_)&"','"& session("Codcli") &"','SI'"
                    
					Set RstProc = Session("Conn").Execute(strProc)   
					'response.Write(strProc)                  
					x=RstProc.Fields.Count
                    %> 
                    <tr  >  
				  		<td colspan="2"> <table cellspacing="1" cellpadding="0" border="0" bordercolor="#FFFFFF" width="1045">                    <tr background="imagenes/fdo-cabecera-cel22.jpg" > 
                    <td  width="100" height="32" background="imagenes/fdo-cabecera-cel22.jpg" align="center"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font id="lblCRamo" name="lblCRamo"  size="1">C&oacute;digo del Ramo</font></b></font></div><div align="center"></div></td>
                    <td width="300" height="32" background="imagenes/fdo-cabecera-cel22.jpg" align="center"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font id="lblNRamo" name="lblNRamo" size="1">&nbsp;&nbsp;Nombre del Ramo</font></b></font></div></td>
                    <td  width="40" height="32" background="imagenes/fdo-cabecera-cel22.jpg" align="center"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Secci&oacute;n</font></b></font></div></td>
                     <td  width="40" height="32" background="imagenes/fdo-cabecera-cel22.jpg" align="center"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Periodo</font></b></font></div></td>
                      <td width="55"  height="32" background="imagenes/fdo-cabecera-cel22.jpg" align="center"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">A&ntilde;o</font></b></font></div></td>
                    	<%for K=2 to x-1 						
							columna=RstProc.fields(k).name
							if trim(columna) <> "N Pr R 0%" and trim(columna)<>"N Ex R 0%" and trim(columna)<>"N Ex 0%" and trim(columna)<>"N Pr 0%"then							
								if (USACONCENTNOTASPA_CENFOTUR = "SI" and left(trim(columna),4) <>"N Pr"and left(trim(columna),4) <>"N Ex"and trim(columna) <>"Concepto Ex"and left(trim(columna),6) <>"N Pr R"and left(trim(columna),6) <>"N Ex R"and trim(columna) <>"Concepto ExRep") OR (USACONCENTNOTASPA_CENFOTUR <> "SI") then
									%>  
									<td height="32" background="imagenes/fdo-cabecera-cel22.jpg" align="center"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1"><%
									if instr(columna, "999") then
										response.Write(replace(columna,"999","Prom"))
									elseif instr(columna, "99") then
										response.Write(replace(columna,"99","Rec"))
									else
										response.Write(columna)
									end if 
									
									%></font></b></font></div></td> 
									<%
								end if
							end if 
						next%>
                    
	                    <%tipocurso = rst("tipocurso") 					
						seccion = rst("codsecc")%>
                    </tr>
                    <tr bgcolor="#DBECF2" class="text-normal-celdas"> 
						<td  width="100" height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=valnulo(Rst("Codramo"),STR_) %></font></div></td>             
                        <td  width="300" height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=valnulo(Rst("Nombre"),STR_) %></font></div></td>   
                        <td  width="40" height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=seccion &" - "& tipocurso %></font></div></td>                   
                        <td  width="40" height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=valnulo(Rst("periodo"),STR_)%></font></div></td>             
                        <td  width="55" height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=valnulo(Rst("Ano"),STR_) %></font></div></td>             
						<%for K=2 to x-1 
							columna=RstProc.fields(k).name 
							if trim(columna) <> "N Pr R 0%" and trim(columna)<>"N Ex R 0%" and trim(columna)<>"N Ex 0%" and trim(columna)<>"N Pr 0%"then
								if (USACONCENTNOTASPA_CENFOTUR = "SI" and left(trim(columna),4) <>"N Pr"and left(trim(columna),4) <>"N Ex"and trim(columna) <>"Concepto Ex"and left(trim(columna),6) <>"N Pr R"and left(trim(columna),6) <>"N Ex R"and trim(columna) <>"Concepto ExRep") OR (USACONCENTNOTASPA_CENFOTUR <> "SI") then
									%>
									<td  height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%
									if not RstProc.eof then						
										if isnull(RstProc(k)) then	
											response.Write("")
										else
											response.Write(RstProc(k))
										end if 
									else
										strNF="select coalesce(NF,0)NF,coalesce(estado,'')estado,coalesce(concepto,'') + ' - ' +(select descripcion from ra_concepto where codigo=ra_nota.concepto)concepto  from ra_nota where ramoequiv='"& valnulo(Rst("ramoequiv"),STR_) &"' and codsecc='"& valnulo(Rst("Codsecc"),STR_) &"' and ano='"& valnulo(Rst("Ano"),STR_) &"' and periodo='"& valnulo(Rst("Periodo"),STR_) &"' and codcli='"& codcli &"'"
				
										Set rstNF = Session("Conn").Execute(strNF)  
										if not rstNF.eof then
											if k =9 then
												if rstNF("NF") = "0" then
													response.Write("")
												else																		
													response.Write(rstNF("NF"))
												end if 
											end if
											if k =10 then
												response.Write(rstNF("estado"))
											end if
											if k =11 then								
												response.Write(rstNF("concepto"))
											end if
										end if 
									end if 
									
									%></font></div></td>            
									<%
								end if
							end if 
						next%>
                        </tr>
                        <tr height="10" bgcolor="#FFFFFF"></tr> 
                        </table>
                    </td> 
                    </tr>
                        
					<%				
                    Rst.movenext	
                    wend
                    %>
                    </table>
                    </td> 
                    </tr>
               
                <%end if %>
                
                <tr valign="bottom"> 
				  <td colspan="3" height="20" class="tex-normales">
                  <font color="#333333" size="2" face="Arial, Helvetica, sans-serif" class="text-normal-celdas">
                  
                  <%strConcepto="select COALESCE(codigo,'') +' : '+ COALESCE(descripcion,'')concepto from ra_concepto WHERE Estado IN ('E','I')"
				  Set rstConcepto = Session("Conn").Execute(strConcepto)%>
                  
                  
                  <BR><BR>LEYENDA <BR>
                  <%while not rstConcepto.eof %>
				  
                  <%=rstConcepto("concepto")%><BR>
                  
                  <%rstConcepto.movenext
				  	wend %>
                  </font>
                  </td>
				</tr> 
				<tr> 
				  <td  colspan="3"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#666666">(Este 
					documento NO constituye certificado)</font> </td>
				</tr>
			  </table>
              <table>
           <tr>
    <td width="83" align="left"><A onMouseOver="MM_swapImage('Image7','','Imagenes/botones/volver-on.gif',1)" onmouseout=MM_swapImgRestore() href="javascript: window.top.location.href ='menu_consultas.asp'"><IMG src="Imagenes/botones/volver-of.gif" border="0"  name="Image7"></A></td>
  </tr>
  </table>
			  </div>
			</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>

<%
Function OcultaOpcion(grupo,nota) 

if grupo = 1 then 
	if nota > cantgrupo1 then
		OcultaOpcion = "style='display:none'"
	end if 
end if 

if grupo = 2 then 
	if nota > cantgrupo2 then
		OcultaOpcion = "style='display:none'"
	end if 
end if 

End Function 
%>

<%ObjetosLocalizacion("concent-notas.asp")%>
</body>

</html>
<!--#INCLUDE file="include/desconexion.inc" -->
