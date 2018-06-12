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


strsql = "SP_LISTA_RAMOS_CONCENT_PA '"& codcli &"'"
'response.Write(StrSql)
Set Rst = Session("Conn").Execute(StrSql)

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
			  <table border="0" cellpadding="15" cellspacing="0" height="187" align="center" width="1200">
				<!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
				<tr> 
                    <td colspan="2"><p><span style="font-size: 14px"><p id ="lblTituloInscripcion" style="font-size: 25px" class="text-menu">
                    Notas del periodo actual</p></span></p></td>
				</tr>
				<tr valign="bottom"> 
				  <td colspan="2" height="20" class="tex-normales"><font color="#333333" size="2" face="Arial, Helvetica, sans-serif" class="text-normal-celdas">(*) Los promedios de las notas están de acuerdo a lo establecido en el sílabo de la asignatura.</font></td>
				</tr> 
                <%  
					cantgrupo1 = 0
					cantgrupo2 = 0
					strg="SP_LISTA_GRUPOS_CONCENT_PA_CENF '"& session("codcli") &"'"
					if BCL_ADO(strg, rstg) then
					
						while not rstg.eof
							if rstg("grupo")= "1" then
								cantgrupo1 = rstg("cantidad")
							end if 
							
							if rstg("grupo")= "2" then
								cantgrupo2 = rstg("cantidad")
							end if
							rstg.movenext
						wend
					end if					
					%>
                    <tr valign="top"> 
                        <td colspan="2" height="70"> <table  cellspacing="1" cellpadding="0" border="0" bordercolor="#FFFFFF">
                        	<tr> 
                        		<td width="150" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Asignatura</font></div></td>
                        		<td width="63" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Ciclo</font></div></td>
                                <td width="63" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Sección</font></div></td>
                                <td width="63" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Asistencia</font></div></td>
                                <td <%=OcultaOpcion(1,1)%> width="30" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">1P1</font></div></td>
                                <td <%=OcultaOpcion(1,2)%>width="30" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">1P2</font></div></td>
                                <td <%=OcultaOpcion(1,3)%>width="30" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">1P3</font></div></td>
                                <td <%=OcultaOpcion(1,4)%>width="30" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">1P4</font></div></td>
                                <td <%=OcultaOpcion(1,5)%>width="30" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">1P5</font></div></td>
                                <td width="35" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>1P Prom</font></b></font></div></td>
                                <td width="35" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>Parcial</font></b></font></div></td>
                                <td <%=OcultaOpcion(2,1)%>width="30" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">2P1</font></div></td>
                                <td <%=OcultaOpcion(2,2)%>width="30" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">2P2</font></div></td>
                                <td <%=OcultaOpcion(2,3)%>width="30" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">2P3</font></div></td>
                                <td <%=OcultaOpcion(2,4)%>width="30" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">2P4</font></div></td>
                                <td <%=OcultaOpcion(2,5)%>width="30" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">2P5</font></div></td>
                                <td width="35" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>2P Prom</b></font></div></td>
                                <td width="35" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>Final</b></font></div></td>
                                <td width="63" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>Sustitutorio</b></font></div></td>
                                <td width="63" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>Promedio</b></font></div></td>
                                <td width="63" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Estado</font></div></td>
                                <td width="115" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">CONCEPTO</font></div></td>
                        	</tr>
                            <%
							
							strsqlC = "SP_LISTA_RAMOS_CONCENT_PA_CENF '"& codcli &"'"
							Set RstC = Session("Conn").Execute(strsqlC)

							if RstC.eof and RstC.bof then 
							else			
								RstC.movefirst			
							end if 	
																							
							while not RstC.eof%>
                            <tr bgcolor="#DBECF2" class="text-normal-celdas"> 
                                <td  height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RstC("NOMBRE")%></font></div></td>             
                                <td  height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RstC("PERIODO")%></font></div></td>   
                                <td  height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RstC("CODSECC")%></font></div></td>                                <td  height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RstC("ASISTENCIA")%></font></div></td>
                                <td <%=OcultaOpcion(1,1)%> height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RstC("N1P1")%></font></div></td>
                                <td <%=OcultaOpcion(1,2)%> height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RstC("N1P2")%></font></div></td>
                                <td <%=OcultaOpcion(1,3)%> height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RstC("N1P3")%></font></div></td>
                                <td <%=OcultaOpcion(1,4)%> height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RstC("N1P4")%></font></div></td>
                                <td <%=OcultaOpcion(1,5)%> height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RstC("N1P5")%></font></div></td>
                                <td  height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><b><%=RstC("N1PPROM")%></b></font></div></td>
                                <td  height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><b><%=RstC("PROMEDIO")%></b></font></div></td>
                                <td <%=OcultaOpcion(2,1)%> height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RstC("N2P1")%></font></div></td>
                                <td <%=OcultaOpcion(2,2)%> height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RstC("N2P2")%></font></div></td>
                                <td <%=OcultaOpcion(2,3)%> height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RstC("N2P3")%></font></div></td>
                                <td <%=OcultaOpcion(2,4)%> height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RstC("N2P4")%></font></div></td>
                                <td <%=OcultaOpcion(2,5)%> height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RstC("N2P5")%></font></div></td>
                                <td  height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><b><%=RstC("N2PPROM")%></b></font></div></td>
                                <td  height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><b><%=RstC("FINAL")%></b></font></div></td>
                                <td  height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><b><%=RstC("SUSTITORIO")%></b></font></div></td>
                                <td  height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><b><%=RstC("NOTAFINAL")%></b></font></div></td>
                                <td  height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RstC("ESTADO")%></font></div></td>
                                <td  height="32"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RstC("CONCEPTO")%></font></div></td>
                            </tr>
                            <%
							RstC.movenext
							wend%>
                                
                        </td>
                    </tr>
					

                <tr valign="bottom"> 
                    <td colspan="12" height="20" class="tex-normales">&nbsp;
                         
                    </td>
                </tr> 
                
                <tr valign="bottom"> 
				  <td colspan="3" height="20" class="tex-normales">&nbsp;
                  
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


</body>
<%ObjetosLocalizacion("concent-notas.asp")%>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
