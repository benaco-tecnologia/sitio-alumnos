

<html>
<head>
<title><%=Session("NombrePestana")%></title>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<link rel="stylesheet" href="css/tex-normales.css" type="text/css">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1">

<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=Malla name=Malla></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=Periodo name=Periodo></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=LC name=LC></OBJECT>

<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/requisitos.inc" -->
<!--#INCLUDE FILE="include/estadoramo.inc" -->
<!--#INCLUDE FILE="include/nombrecarrera.inc" -->
<!--#INCLUDE FILE="include/nombreramo.inc" -->
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
dim codpaln,i
dim strsql
dim valor
dim valorAux
dim ArregloMalla
'dim strmalla
Codcarr=session("Codcarr")
codplan =session("codpestud")
'Response.Write(Codplan)
codcli=session("codcli")
'response.Write(codcarr)
'response.End
%>

<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if

if EstaHabilitadaNW (467)="S" then 'SEGUIR ACA !!
	if GetPermisoNW(467) ="N" then
		'response.Redirect("MensajesBloqueos.asp")	
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
				'response.Redirect("MensajeBloqueo.asp")
			end if 
		end if 
		
	end if 
else
	response.Redirect("MensajeBloqueoHabilita.asp")
end if 

Audita 467,"Ingresa a Malla curricular"


strParameR="SELECT coalesce(dbo.Fn_ValorParame('OCULTAPREREQUISITOPA'),'NO')Parame"
set rstParameR= Session("Conn").Execute(strParameR)		
if not rstParameR.eof then
	OCULTAPREREQUISITOPA=rstParameR("Parame")
end if


strL="SP_CONSTANTE_LOCALIZACION 'LOCALIZACION'"
set rstL= Session("Conn").Execute(strL)		
if not rstL.eof then
	LOCALIZACION=rstL("ct_constantevalor")
end if
%>

<table border="0" cellpadding="0" cellspacing="0"  align="left">
  <tr>
	<td>
		<table border="0" cellpadding="0" cellspacing="0" >
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
			<div>
			   <table border="0" cellpadding="0" cellspacing="15" height="111" align="left" width="800">
				<!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
				<tr> 
				  <td><img name="fder_r2_c1" src="imagenes/titulos/T-malla_curri.gif" width="271" height="38" border="0"></td>
			     </tr>
				<tr valign="bottom"> 
				  <td height="20" valign="top" class="tex-normales"> 
					<table width="800" border="0">
					  <tr valign="bottom">
						<td align="left"><img src="Imagenes/cuadros-color.jpg"></td>
					  </tr>
					  
					  <tr valign="bottom"> 
					  	<td height="1" colspan="2" class="Tit-celdas"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">CARRERA:</font><font size="2" class="Tit-celdas">&nbsp;&nbsp;&nbsp;&nbsp;<font size="2" class="Tit-celdas"><%=GetNombreCarrera(CodCarr)%></font></font></td>
				      </tr>
					</table>	
                    <!-- Mejora Cuadro Creditos -->
                    <%
					strV="SELECT dbo.Fn_ValorParame('visualizarResumenElectivos')Parame"
					if BCL_ADO(strV, rstV) then
						visualizarResumenElectivos=trim(valnulo(rstV("Parame"),STR_))
					end if
						
					if visualizarResumenElectivos = "SI" then%>
                    <tr valign="top">
				      <td colspan="2" height="70"><table width="604" cellspacing="2" cellpadding="3" height="72" border="0" bordercolor="#FFFFFF">
				        <tr class="text-cabecera-celda" background="imagenes/fdo-cabecera-cel22.jpg" >
				          <td width="208" height="30" height="30" background="imagenes/fdo-cabecera-cel22.jpg" class="textos"><div align="center" class="Estilo2"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1">Tipo de Curso</font></font></div></td>
				          <td width="104" height="30"height="30" background="imagenes/fdo-cabecera-cel22.jpg" class="textos"><div align="center" class="Estilo2"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1">Exigencia</font></font></div></td>
				          <td width="104" height="30" height="30" background="imagenes/fdo-cabecera-cel22.jpg" class="textos"><div align="center" class="Estilo2"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1">Avance</font></font></div></td>		         
				          </tr>
				        <%
						strsql="SP_CUADRO_CREDITOS_ALUMNO_PE '"& session("codcli") &"' " 
						Set Rst = Session("Conn").Execute(StrSql) 
						do while not rst.eof 							
                         %>
				        <tr bgcolor="ffc172">
				          <td width="208" height="32" bgcolor="#DBECF2" class="textos"><div align="left" class="Estilo4"><font face="Arial, Helvetica, sans-serif" size="1"><%=rst(0)%></font></div></td>
                          <td width="104" height="32" bgcolor="#DBECF2" class="textos"><div align="right" class="Estilo4"><font face="Arial, Helvetica, sans-serif" size="1"><%=rst(1)%></font></div></td>
                          <td width="104" height="32" bgcolor="#DBECF2" class="textos"><div align="right" class="Estilo4"><font face="Arial, Helvetica, sans-serif" size="1"><%=rst(2)%></font></div></td>
				          <%  rst.movenext %>
				          </tr>
				        <% loop%>
			          </td>
			        </tr>
                    <tr>
                        <td>&nbsp; </td>
                    </tr>
                    <%end if%>
                    <!-- Mejora Cuadro Creditos -->		
                    	  
					<p>
					  <%
			dim strsqlperiodo,numlineas,numcolumnas,StrLC
			strSql ="select b.nombre, b.codramo,convert(numeric, a.nivel) as nivel, a.orden, b.credito, COALESCE(convert(numeric,b.duracion),0) as duracion " &_
					"From ra_curric a WITH (NOLOCK), ra_ramo b WITH (NOLOCK) " &_ 
					"where a.codpestud='" & codplan & "' and b.codramo=a.codramo and visibleenmallaweb = 'S' " &_ 
					"Order by convert(numeric,a.nivel), a.orden "
			'Set Malla = Conn.execute(strSql)

			Malla.Open strSql, Session("Conn")
			'response.Write(strSql)
			'response.end
			
			'StrSqlperiodo ="select nummaxper From mt_alumno WITH (NOLOCK) where codcli='" & codcli & "'" 
			StrSqlperiodo ="SELECT CASE WHEN c.TIPO_REGIMEN = 0 then CONVERT(INTEGER,c.NUM_MAX_PERIODO) " &_
			"ELSE (SELECT CONVERT(INTEGER,num_max_periodo) FROM RA_PESTUD WHERE CODPESTUD = a.CODPESTUD) END nummaxper " &_
			"FROM mt_carrer c, mt_alumno a WHERE a.codcli = '" & codcli & "' AND a.CODCARPR = c.CODCARR"
			
			'Set Periodo = Conn.execute(strsqlperiodo)
			Periodo.Open strsqlperiodo, Session("Conn")
			'response.Write(Periodo("nummaxper")&"periodo"&StrSqlperiodo)
			'response.End()
			
			StrLC ="SELECT MAX(orden) lineas,MAX(nivel) columnas From ra_curric a WITH (NOLOCK), ra_ramo b WITH (NOLOCK) " &_
			       "where a.codpestud='" & codplan & "' and b.codramo=a.codramo and visibleenmallaweb = 'S'"
			LC.Open StrLC, Session("Conn")
			
			'response.Write(StrLC)
			'response.End()
			'response.Write(LC("lineas") &" asd" & Periodo("nummaxper"))
			'response.End()
		    
			
			
			strDuracion="select coalesce(duracion,0)duracion  from RA_PESTUD WHERE CODPESTUD='" & codplan & "'"
			set rstDuracion= Session("Conn").Execute(strDuracion)		
			if not rstDuracion.eof then
					cols=Cdbl(rstDuracion("duracion")) 'Nuevo metodo sacado de la ficha curricular Net
				else
					cols = Cdbl(valnulo(LC("columnas"),NUM_)) 'metodo antiguo
			end if
			
			
			numcolumnas = cols / Cdbl(Periodo("nummaxper"))
			numlineas = Cdbl(valnulo(LC("lineas"),NUM_))
			numcolumnas = Ceiling(Cdbl(numcolumnas))
			
			valor= cols'session("NumeroMaximoPrueba") ' maximo de periodos del plan de estudio del Alumno
			if valor > 1 then
				valorAux = session("NumeroMaximoPrueba")
			else
				valorAux = 2
			end if
			
			if not Malla.eof then
			  ArregloMalla = Malla.GetRows
			else 
			  Response.Write("No tiene asignaturas configuradas con visible la malla web...")
			  Response.end  
			end if
			
			if valor > 2 then
			   valor = round(12 * clng(valor) )
			elseif valor = 1 then
				valor=round(12 * 2 )
			else
				valor=round(12 * clng(valor))
			end if 
			
			
			redim mimalla (15,clng(valor))
			redim mimalla2(15,clng(valor))
			redim mimalla3(15,clng(valor))
			
			'ArregloMalla, obtiene la cantidad de ramos en la malla
			 'response.Write(ubound(ArregloMalla,2))
			 'response.End()
			i=0
			do while i <= ubound(ArregloMalla,2)
			 '  response.Write(ArregloMalla(4,i)) & "</br>"			 
			   mimalla(ArregloMALLA(3, i), ArregloMALLA(2, i)) = LCase(ArregloMalla(1, i) & " - " & ArregloMalla(0, i)) & " (" &  ArregloMalla(4,i) & ")"
			 '  response.Write(mimalla(ArregloMALLA(3, i), ArregloMALLA(2, i)) & ArregloMalla(0, i)) & " (" &  ArregloMalla(4,i) & ")")
			 '  response.Write(LCase(ArregloMalla(0, i)) & " (" &  ArregloMalla(4,i) & ")")
			   if clng(Arreglomalla(5, i)) = clng(valor) then
				   mimalla(ArregloMALLA(3, i), clng(ArregloMALLA(2, i)) + 1) = LCase(ArregloMalla(1, i) & " - " & ArregloMalla(0, i)) & " (" &  ArregloMalla(4,i) & ")"
			   end if
			   
			   if trim(Arreglomalla(1, i) ) <> "" then
				   mimalla2(ArregloMALLA(3, i), ArregloMALLA(2, i))=trim(EstadoRamo(codcli,Arreglomalla(1, i) ))
				   'Response.Write(LCase(ArregloMalla(0, i)) & ":" & estadoramo(codcli,Arreglomalla(1, i) ) & ":<br>")
				   'response.end
				   mimalla3(ArregloMALLA(3, i), ArregloMALLA(2, i))=PreRequisito(Arreglomalla(1, i), codplan) 
			   end if
			   i=i+1
			   
			loop
			
			%>
				    </p>
					<p>
					  <%
			periodo.close()
			
			valor=ValorAux ' maximo de periodos de esta carrera
			%>
					</p>
					<table border="0" cellpadding="0" cellspacing="1" height="146" align="left" width="825">
                      <tr background="Imagenes/fdo-cabecera-cel22.jpg">
                        <%for i=1 to Cint(numcolumnas)%>
                        <td height="30"  colspan=<%=valor%> background="Imagenes/fdo-cabecera-cel2.jpg" align="left"><div align="center"><b><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> A&ntilde;o <%=i%> </font></b> </div>
                            <div align="center"></div></td>
                        <%next%>
                      </tr>
                      <tr bgcolor="8394d1">
                        <%for i=1 to Cint(numcolumnas) %>
                        <% O=0'variable=0
					  'response.Write(clng(valor))
					  'response.End
					  for j=1 to clng(valor)
					  
						  if LOCALIZACION ="PERU" then
							PeriodoColumna = PeriodoColumna + 1
						  else
							PeriodoColumna = j
						  end if 
					  %>
                        <td width="520" height="15" background="Imagenes/fdo-cabecera-cel22.jpg"><div align="center" class="text-cabecera-celda"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><font size="1">
                          <font name="lblPeriodo">Periodo</font> <%=PeriodoColumna%></font></font></div></td>
                        <%next%>
                        
                        <%next%>
                      </tr>
                      <%
					  O=0
					  for i = 1 to numlineas'  ubound(mimalla, 1)  'aqui defino el numero de lineas 
					  Response.write("<tr > ")
						for j = 1 to cols' ubound(mimalla, 2) 'aqui columnas
						'response.Write(ucase(Mid(mimalla(i,j),1,Len(Mid(mimalla(i,j), 1, InStr(1,mimalla(i,j), " "))))))
						'Response.Write(mimalla2(i,j))
						
						select case trim(mimalla2(i,j))
						case ":C"
						   'color="#DBECF2" ' fondo
						   fondo = "background='Imagenes/fondo_gris.gif'"
						case ":R"
						   'color="#FFFF55" 'amarillo
						   fondo = "background='Imagenes/fondo_amarillo.gif'"
						case ""
						   'color="#DBECF2" ' fondo
						   fondo = "background='Imagenes/fondo_gris.gif'"
						case ":"
						   'color="#55FFFF"  ' celeste
						   fondo = "background='Imagenes/fondo_celeste.gif'"
						case ":A"
						   'color="#80FF80"  ' verde
						   fondo = "background='Imagenes/fondo_verde.gif'"
						case ":E"
						   'color="#80FF80"  ' verde
						   fondo = "background='Imagenes/fondo_verde.gif'"
						case ":I"
						   'color="#80FF80"  ' verde
						   fondo = "background='Imagenes/fondo_verde.gif'"
						case ":P"
						   'color="#DBECF2" ' fondo
						   fondo = "background='Imagenes/fondo_gris.gif'"
						end select
						
						if j + 1 <= ubound(mimalla,2) then	
						'obtengo el codigo de ramo
						RAMO=ucase(Mid(mimalla(i,j),1,Len(Mid(mimalla(i,j), 1, InStr(1,mimalla(i,j), " ")))))
						'pregunto su duracion						
						strMalla="select duracion from ra_ramo where codramo='"& trim(ramo) &"'"				
						if bcl_ado(strMalla,rstMalla) then
							Durac=ValNulo(rstMalla("duracion"),NUM_)
						else
							Durac=0	
						end if
						
						   if Durac=2 then 
						   	  if OCULTAPREREQUISITOPA <> "SI" then
								 Response.Write("<td align='center' width='520' height='50' colspan=2 " & fondo & " color='#FFFFFF'> <font size='1' class='tex-totales-celda'  > <a href=" & chr(34) & "javascript:prueba('" & mimalla3(i,j) & "')" & chr(34) & "> " & ucase(Mimalla(i,j)))	
							  else
							  	 Response.Write("<td align='center' width='520' height='50' colspan=2 " & fondo & " color='#FFFFFF'> <font size='1' class='tex-totales-celda'  >" & ucase(Mimalla(i,j)))	
							  end if 
							  	 
							 j = j + 1
						   else
	    					 Response.Write("<td align='center' width='520' height='50' color='#FFFFFF' " & fondo & "><font size='1' class='tex-totales-celda'  > " )
							 
							 if OCULTAPREREQUISITOPA <> "SI" then
							 	Response.Write("<a href=" & chr(34) & "javascript:prueba('" & mimalla3(i,j) & "')" & chr(34) & ">" & ucase(Mimalla(i,j)) & "</a> ")
							 else
							 	Response.Write("<FONT COLOR='blue'>" & ucase(Mimalla(i,j)) & "</FONT>")							
							 end if 
							 
						   end if						    
						end if
						Response.Write("</font>")
						Response.Write("  </td>")
						
					   next 
					  Response.Write("</tr>")					  
					  next
					  %>
                      <!--<tr bgcolor="ffc172"> 
						<td width="40" height="16"> 
						 <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div> 
						</td>
						<td width="41" height="16"> 
						<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div> 
						</td>
						<td width="41" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="42" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="43" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="43" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="41" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="42" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="41" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="42" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="43" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="44" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
					  </tr>
					  <tr bgcolor="ffc172"> 
						<td width="40" height="16"> 
						 <!-- <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="41" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="41" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="42" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="43" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="43" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="41" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="42" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="41" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="42" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="43" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
						<td width="44" height="16"> 
						  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"></font></div>
						</td>
					  </tr>-->
                      <!-- fin rutina -->
                      <tr valign="top">
                        <td colspan="2" height="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#666666">(Este 
                          documento NO constituye certificado)</font> </td>
                      </tr>
                      <tr>
                          <td align="left">
                          <table>
                       <tr>
                <td width="83" align="left"><A onMouseOver="MM_swapImage('Image7','','Imagenes/botones/volver-on.gif',1)" onmouseout=MM_swapImgRestore() 				                 href="javascript: window.top.location.href ='menu_tomaderamos.asp'"><IMG src="Imagenes/botones/volver-of.gif" border="0"  name="Image7">                 </A></td>
              </tr>
              </table>
                          </td>
                      </tr>
                    </table>
                    
					<p>&nbsp;</p></td>
			    </tr>
		      </table>
              
			<p>&nbsp;</p>
			<p>&nbsp;</p>
			<p>&nbsp;</p>
			<P>&nbsp;</P>
			<P>&nbsp;</P>
			<P>&nbsp;</P>
			<P>&nbsp;</P>
			</table>
            
			</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>

			</body>
			
			<script>
			function prueba(ramo)
			{
			  alert(ramo) 
			}
			</script>
			<%ObjetosLocalizacion("malla-curri.asp")%>
			</html>
			<!--#INCLUDE file="include/desconexion.inc" -->
