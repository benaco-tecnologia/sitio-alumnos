<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->

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

Ano = Session("Ano")
Periodo = Session("Periodo")
' desde la tabla mt_parame obtener el año y periodo
strRamosDebe = " SELECT b.codramo, b.nombre, a.inscrito, a.codsecc, a.Ramoequiv, c.nombre as NombreReal, tipo.tipocurso,"
strRamosDebe = strRamosDebe & " tipo.fechacert1, tipo.fechacert2, tipo.fechacert3, tipo.fechaexamen, tipo.fechaexamener"
strRamosDebe = strRamosDebe & " FROM ra_carga a, ra_ramo b, ra_ramo c , ra_seccio tipo" 
strRamosDebe = strRamosDebe & " WHERE a.ramoequiv = b.codramo and " 
strRamosDebe = strRamosDebe & " a.ramoequiv *= c.codramo and " 
strRamosDebe = strRamosDebe & " a.inscrito = 'S' and " 
strRamosDebe = strRamosDebe & " a.codcli ='" & CodCli & "' and " 
strRamosDebe = strRamosDebe & " a.ano = '" & ano & "' and " 
strRamosDebe = strRamosDebe & " a.periodo = '" & periodo & "'" 
strRamosDebe = strRamosDebe & " and b.codramo=tipo.codramo and a.codsecc=tipo.codsecc "
strRamosDebe = strRamosDebe & " and tipo.ano='"& ano &"'"
strRamosDebe = strRamosDebe & " and tipo.periodo='" & periodo & "'"

strRamosDebe = strRamosDebe & " Union "

strRamosDebe = strRamosDebe & " SELECT b.codramo, b.nombre, a.inscrito, a.codsecc, a.Ramoequiv, c.nombre as NombreReal,tipo.tipocurso, " 
strRamosDebe = strRamosDebe & " tipo.fechacert1, tipo.fechacert2, tipo.fechacert3, tipo.fechaexamen, tipo.fechaexamener"
strRamosDebe = strRamosDebe & " FROM ra_cargaactividad a, ra_ramo b, ra_ramo c ,ra_seccio tipo " 
strRamosDebe = strRamosDebe & " WHERE a.ramoequiv = b.codramo and " 
strRamosDebe = strRamosDebe & " a.ramoequiv *= c.codramo and " 
strRamosDebe = strRamosDebe & " a.inscrito = 'S' and " 
strRamosDebe = strRamosDebe & " a.codcli ='" & CodCli & "' and " 
strRamosDebe = strRamosDebe & " a.ano = '" & ano & "' and " 
strRamosDebe = strRamosDebe & " a.periodo = '" & periodo & "'" 
strRamosDebe = strRamosDebe & " and b.codramo=tipo.codramo and a.codsecc=tipo.codsecc "
strRamosDebe = strRamosDebe & " and tipo.ano='"& ano &"'"
strRamosDebe = strRamosDebe & " and tipo.periodo='" & periodo & "'"

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
                " a.periodo = '" & periodo & "' Order By a.prioridad "

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
                " a.PerInscrip = " & SESSION("PER_ID") & " and " & _
                " a.periodo = '" & periodo & "' Order By a.prioridad "
'agregar el año y periodo a esta query
'RESPONSE.WRITE(strRamosSolici)
'RESPONSE.End
RamosSolici.Open strRamosSolici, Session("Conn")

%>

<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title>Portal de Alumnos en L&iacute;nea</title>
<table border="0" cellpadding="0" cellspacing="0"  align="center">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" ><% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td valign="top" align="center">
			<% CargarTop1()%><% SubMenu()%>
			<table width="100%" border="0" cellspacing="0" cellpadding="20" height="550" bgcolor="#FFFFFF">
			  <tr> 
  <form name="form1" method="post" action="asignatura-seccion.asp">
    <input type = "hidden" name = "Ramos" value = "">
    <table border="0" cellpadding="0" cellspacing="0" align="left" width="844">
      <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
      <tr valign="top"> 
	   <td colspan="2"><img name="fder_r1_c1" src="imag/f-der/f-der_r1_c1.gif" width="336" height="19" border="0"></td>
        <td width="489"><img name="fder_r1_c2" src="imag/f-der/f-der_r1_c2.gif" width="248" height="19" border="0"></td>
      </tr>
      <tr> 
        <td colspan="2"><img name="fder_r2_c1" src="imag/f-der/resumen.gif" width="336" height="28" border="0"></td>
        <td width="489"><img src="imag/f-der/no-ayuda.gif" width="248" height="28"></td>
      </tr>
      <tr valign="bottom"> 
        <td colspan="2"><a href="javascript:;" onMouseUp="javascript:visualiza();"><img src="imag/botones/instrucciones.gif" width="122" height="23" border="0"></a></td>
        <td width="489"> 
          <div align="center"><a href="javascript:imprime()"><img src="imag/f-der/impresion.gif" width="36" height="36" border="0"></a></div>
        </td>
      </tr>
      <tr valign="bottom"> 
        <td colspan="3">&nbsp;</td>
      </tr>
      <tr valign="top"> 
        <td height="1" colspan=2><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="4a5da1">Resultado 
          de Inscripción</font></b></td>
        <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="4a5da1"><%=GetNombrealumno(Codcli)%></font></td>
      </tr>
      <tr valign="bottom"> 
        <td width="203"><b><font face="Geneva, Arial, Helvetica, san-serif" size="1">Fecha:<%=date()%></font></b></td>
        <td width="152"><b><font face="Geneva, Arial, Helvetica, san-serif" size="1">Hora 
          :<%=time()%> </font></b></td>
        <td width="489"> 
          <div align="right"></div>
        </td>
      </tr>
      <tr valign="top"> 
        <td colspan="3" height="80"> 
          <table width="706" cellspacing="0" cellpadding="0" height="59" border="1" bordercolor="#FFFFFF">
            <td height="17" colspan="11"> <div align="left"> 
                <p><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#000000">Ud. 
                  Tiene inscritas las siguientes asignaturas:</font></p>
              </div></td>
            </tr>
            <tr bgcolor="4a5da1"> 
              <td height="16" width="66" bgcolor="4a5da1"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">C&oacute;digo</font></b></font></div></td>
              <td height="16" width="193"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Asignatura</font></b></font></div></td>
              <td height="16" width="43"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Secci&oacute;n</font></b></font></div></td>
              <td height="16" width="46"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Horario</font></b></font></div></td>
              <td height="16" width="94"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Ramo</font></b></font></div></td>
              <td height="16" width="37"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Tipo</font></b></font></div></td>
              <td width="66"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Parcial 
                1</font></strong></td>
              <td width="69"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Parcial 
                2</font></strong></td>
              <td width="69"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Parcial 
                3</font></strong></td>
              <td width="67"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ex. 
                Final</font></strong></td>
              <td width="70"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ex. 
                Rep</font></strong></td>
            </tr>
            <% 
		  If RamosDebe.Eof Then
		  %>
            <%
		  Else
		   While Not RamosDebe.Eof
		     if Ucase(RamosDebe("inscrito")) = "S" then
		       CodSecc = RamosDebe("CodSecc")
		       Horario = GetHor2(RamosDebe("RamoEquiv"), CodSede, RamosDebe("CodSecc"), Ano, Periodo)
		     else
		       CodSecc = ""
		       Horario = ""
		     end if
		   %>
            <tr bgcolor="ffc172" height="25"> 
              <td width="66" height="8" align="center"><font face="Verdana" size="1"><%=RamosDebe("CodRamo")%></font></td>
              <td width="193" height="0" align="center"><font face="Verdana" size="1"><%=RamosDebe("Nombre")%></font></td>
              <td width="43" height="0" align="center"><font face="Verdana" size="1"><%=CodSecc%></font></td>
              <td width="46" height="8" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=Horario%></font></td>
              <td width="94" height="16" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=RamosDebe("NombreReal")%></font></td>
              <td width="37" height="16" align="center"><font size="1" face="Arial, Helvetica, sans-serif"> 
                <%if RamosDebe("Tipocurso")="T"  then
			                                                                                                     response.write("Teórico")
																											  end if	 
																											  if  RamosDebe("Tipocurso")="L" then
																											      response.write("Laboratorio")
																											  end if
																											   if RamosDebe("Tipocurso")="P" then
																											      response.write("Practico")
																											   end if
																											   
																											   
			                          %>
                </font></td>
              <td width="66" height="16" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=RamosDebe("FechaCert1")%></font></td>
              <td width="69" height="16" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=RamosDebe("FechaCert2")%></font></td>
              <td width="69" height="16" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=RamosDebe("FechaCert3")%></font></td>
              <td width="67" height="16" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=RamosDebe("FechaExamen")%></font></td>
              <td width="70" height="16" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=RamosDebe("FechaExamenER")%></font></td>
            </tr>
            <%
		     RamosDebe.MoveNext
		   Wend
		 End If %>
          </table>
          <br>
          <table width="707" cellspacing="0" cellpadding="0" height="59" border="1" bordercolor="#FFFFFF">
            <td height="17" colspan="6"> 
              <div align="left"> 
                <p><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#000000">Ud. 
                  solicit&oacute; las siguientes asignaturas:</font></p>
              </div>
            </td>
            </tr>
            <tr bgcolor="4a5da1"> 
              <td height="16" width="40" bgcolor="4a5da1"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">C&oacute;digo</font></b></font></div>
              </td>
              <td height="16" width="165"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Asignatura</font></b></font></div>
              </td>
              <td height="16" width="49"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Secci&oacute;n</font></b></font></div>
              </td>
              			  <td height="16" width="62"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Horario</font></b></font></div>
              </td>
              <td height="16" width="46"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Situación</font></b></font></div>
              </td>
              <td height="16" width="109"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Ramo</font></b></font></div>
              </td>
              
			  <%if paso=1 then%>
			   <td height="16" width="49"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Secci&oacute;n Solicitada</font></b></font></div>
              </td>
			  <%end if%>
              
			  <%if paso=1 then
			  
				  if Ucase(RamosSolici("Estado")) = "A" then %> 
				  <td height="16" width="163"> 
					<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Observaci&oacute;n</font></b></font></div>
				  </td>
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
		   
            <tr bgcolor="ffc172" height="25"> 
              <td width="40" height="8" align="center"><font face="Verdana" size="1"><%=RamosSolici("CodRamo")%></font></td>
              <td width="165" height="0" align="center"><font face="Verdana" size="1"><%=RamosSolici("Nombre")%></font></td>
              <td width="49" height="0" align="center"><font face="Verdana" size="1"><%=CodSecc%></font></td>
			  
              <td width="62" height="0" align="center"><font face="Verdana" size="1"><%=Horario%></font></td>
              <td width="46" height="16" align="center"><font face="Verdana" size="1"><%If RamosSolici("Estado")="A" then
			                                                                                 response.write("Aprobado")
																						 elseif RamosSolici("Estado")="P" then
																						       response.write("Pendiente")
																					     else
																						      response.write("Reprobado")
																						 end if	   
			                                                                           '=RamosSolici("Estado")%></font></td>																		                                        
																																		 
              <td width="109" height="16" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=RamosSolici("NombreReal")%></font></td>
			   <%  if paso=1 then %> 
			         <%'if RamosSolici("codsecc")<>RamosSolici("CodSeccori") then %>
			        <td width="49" height="16" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=RamosSolici("CodSeccori")%></font></td>
				    <td width="163" height="16" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%if Ucase(RamosSolici("Estado")) = "A" then
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
        </td>
      </tr>
    </table>
  </form>
			  </tr>
			</table>
			</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>
</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->