<%Response.Expires = -1 %>

<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%
X=REQUEST("X")
IF X=1 THEN
	 
	 %>
	
	 <script>
	 alert("Esta sección no corresponde a la asignatura teórica seleccionada")
	 </script>
	
	 <%
END IF

strcombo="SELECT coalesce(dbo.Fn_ValorParame('USACOMBOSTR'),'NO')USACOMBOSTR"
set rstcombo= Session("Conn").Execute(strcombo)		
if not rstcombo.eof then
	USACOMBOSTR=rstcombo("USACOMBOSTR")
end if
	
Dim strRamosDebe, strRamosPuede, Rut, CodCli
Dim RstHorario,Leyenda

CodCli = Session("CodCli")
IF CodCli = "" then
   'Response.Redirect "pagina de login"
   'Response.Write( "Me voy : " + CodCli + " " + session("CodCli") )
end if
Ano = AnoAcad()
Periodo =PeriodoAcad()
CodSede = Session("CodSede")
P =  REQUEST("P")
M =  Request("M")
NC =  Request("NC")
RC =  Request("RC")

Dim MtrHorario(60, 28)
Dim MtrHorarioRAMO(60, 28)
Dim MtrHorarioCODSECC(60, 28)
Dim MtrHorSt(60, 28)
Dim MtrHorRamo(60, 28)
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

  'response.Write(strHorario)
  'response.End()
  
  'Response.Write("horario.asp Consulta con tabla tmpsolici, tabla que no existe, ¿cuando se llena? , dejo standby mientras..... (preguntar a mauro)  ")
  'response.End()
' response.write("S->"+ucase(REQUEST("S"))+"T->"+ucase(REQUEST("T")))
' response.End()
  
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

  'Response.Write("horario.asp Consulta con tabla tmpsolici, tabla que no existe cuando se llena? , dejo standby mientras..... ")
  'response.End()
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
  strHorario = strHorario & " c.inscrito = 'S' and "
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
  strHorario = strHorario & " c.inscrito = 'S' and "
  strHorario = strHorario & " h.periodo = '" & periodo & "' and "
  strHorario = strHorario & " h.CodSede = '" & Codsede & "' "
  strHorario = strHorario & " order by h.dia, h.codmod"
end if
'RstHorario.Open strHorario, Conn
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
'response.Write(strHorario)


%>

<script Language="JavaScript">
function MM_openBrWindow(theURL,winName,features) 
{ 
  window.open(theURL,winName,features);
}
</script>
<html>
<head>
<title>Resultado de Inscripcion</title>

<link href="css/tablas.css" rel="stylesheet" type="text/css">
</head>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript"  SRC="include/tooltip.js"></SCRIPT>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div id="tooltip" align="center" style="position:absolute;visibility:hidden;border:1px solid black;font-size:50px;layer-background-color:lightyellow;background-color:lightyellow;padding:1px"></div>
<%
'response.Write("5")

If Trim(M) <> "" then
Dim Rst
GetTopes CodCli, Request("C"), Request("R"), Request("S"), CodSede, Ano, Periodo, Rst
Mensaje = "Error : El Ramo " & Request("R") & "/" & Request("S") & "\n"
Mensaje = Mensaje & "Posee tope con los siguientes ramos : \n\n  "
do while not Rst.eof
     Mensaje = Mensaje & Rst("Dia") & ":" & Rst("CodMod") & " : Ramo = " & Rst("CodRamo") & "\n  " 
     MtrHorSt( Rst("CodMod"), GetDia(Rst("DIA"))) = "E"

     Rst.movenext
loop
Rst.close()

If Trim(NC) <> "" then
	if NC ="SIN COMBO" then
		Mensaje = "La asignatura seleccionada no se encuentra definida en el combo, no la puede seleccionar."
	else
		Mensaje = "La asignatura seleccionada no es parte del combo " & NC & ", favor seleccionar otra secci\u00f3n."
	end if 
end if 
'Response.Write("mensaje = " & Mensaje)
'Response.End   

	if USACOMBOSTR = "SI" then%>
		<script languaje="javascript">
            parent.Seleccion.location="asignatura-seccion-combos.asp?Ramos=";
        </script>
	<%else%>
		<script languaje="javascript">
            parent.Seleccion.location="asignatura-seccion.asp?Ramos=";
        </script>
	<%end if

end if
%>
<div align="left">
      <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
	  <table cellspacing="15">

	  <%
	  if USACOMBOSTR = "SI" then%>
            <tr align="center"> 
            	<td> <a href="javascript:imprime()"><img id="imgPrint" src="imag/f-der/impresion.gif" width="36" height="36" border="0" align="center"> </a></td>
            </tr>
            <tr valign="top" align="left"> 
				<%if session("ComboTR") <> "" then
					strcc="SP_TRAE_DATOS_COMBO_TR "& session("ComboTR") &""
					set RSTcc= Session("Conn").Execute(strcc)		
					if not RSTcc.eof then
						NOMBRE = RSTcc("NOMBRE")
						CARRERA = RSTcc("CARRERA")
						JORNADA = RSTcc("JORNADA")
						NIVEL = RSTcc("NIVEL")
					end if
					
					valorTit ="Combo " & NOMBRE & " - Nivel " & NIVEL & " - Horario de Asignaturas"
				else 
					valorTit = "Horario de asignaturas"
				end if%>
            	<td width="667" height="1" align="center" class="Tit-celdas"><b><font size="1" class="Tit-celdas" color="4a5da1"><%=valorTit%></font></b></td> 
            </tr>
	  <%else%>      
            <tr valign="top" align="left"> 
                <td width="667" height="1" align="left" class="Tit-celdas"><b><font size="1"  id="lblHorario"class="Tit-celdas" color="4a5da1">Horario 
                de asignaturas:</font><font class="text-normal-celdas"> <%=GetNombrealumno(Codcli)%></font></b></td> 
            </tr>
      <%end if%>
      
      <tr valign="top"> 
        <td colspan="2" height="2">
		<table id ="tablaHorario" width="750" cellspacing="1" cellpadding="0" height="72" border="0" bordercolor="#FFFFFF" align="left">
          <tr background="Imagenes/fdo-cabecera-cel22.jpg"> 
              <td width="20" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Mod</font></b></font></div>
            </td>
              <td width="81" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Hora</font></b></font></div>
            </td>
              <td width="78" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Lunes</font></b></font></div>
            </td>
              <td width="78" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Martes</font></b></font></div>
            </td>
              <td width="78" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Mi&eacute;rcoles</font></b></font></div>
            </td>
              <td width="78" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Jueves</font></b></font></div>
            </td>
              <td width="78" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Viernes</font></b></font></div>
            </td>
            <td width="84" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">S&aacute;bado</font></b></font></div>
            </td>
            <td width="84" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Domingo</font></b></font></div>
            </td>
          </tr>
			<% 
			dim DiaMaxModulo
			Set RstModulos=Session("Conn").execute("select top 1 dia from ra_modulo where dia='LUNES' and codmod in (select max(convert(int,codmod)) as codmod from ra_modulo )")
			'response.Write("select top 1 dia from ra_modulo where dia='LUNES' and codmod in (select max(convert(int,codmod)) as codmod from ra_modulo )")
			'response.End()
			if RstModulos.eof then
				DiaMaxModulo="LUNES" 
			else
				DiaMaxModulo=RstModulos(0)
				RstModulos.close 	
			end if 
			
			%>
        <% strModulos = "Select Codmod, Hor_Ini, Hor_Fin from ra_modulo where CodSede = '" & CodSede & "' and dia = '"& trim(DiaMaxModulo) &"' Order By convert(numeric,Codmod)" 
		'response.Write(strModulos)
           set RstModulos = Session("Conn").execute(strModulos)
           do while not RstModulos.eof
             i = RstModulos("CodMod")
        %>
		  <div id="tooltip" align="center" style="position:absolute;visibility:hidden;border:1px solid black;font-size:12px;layer-background-color:lightyellow;background-color:lightyellow;padding:1px" ></div>
          <tr bgcolor="#DBECF2"> 
             
            <td width="20" height="17" align="center"><font face="Arial" size="1"><%=RstModulos("CodMod")%></font></td>
             <td width="81" height="17" align="center"><font face="Arial" size="1"><%=RstModulos("Hor_Ini") & " - " & RstModulos("Hor_Fin")%></font></td>
          <% for j = 1 to 7 %>
             
             <% if trim(MtrHorSt(i, j)) = "" then 
                  color = "#DBECF2"
             %>
              <td width="78" height="17"><font face="Verdana" size="1">  <br> &nbsp 
             <% 
                else
				   
				   if MtrHorSt(i,j) ="E" then
				   	   'Analiza si tiene tope de acuerdo a fechas
					   RamoNoTope = MtrHorarioRAMO(i, j)
					   CodsecNoTope = MtrHorarioCODSECC(i, j)
					   ModuloNoTope = RstModulos("CodMod")
					   DiaNoTope = MtrHorarioDia(i, j)
					   
						strNoTope = "SELECT h.codramo+'/'+CONVERT(VARCHAR,h.codsecc)+'-'+h.codsal Ramo FROM ra_horari h,ra_carga c"
						strNoTope = strNoTope &" WHERE   c.codcli = '"& session("codcli") &"' "
						strNoTope = strNoTope &" AND c.ramoequiv_i = h.codramo "
						strNoTope = strNoTope &" AND c.codsecc_i = h.codsecc"
						strNoTope = strNoTope &" AND c.ano = h.ano"
						strNoTope = strNoTope &" AND c.periodo = h.periodo"
						strNoTope = strNoTope &" AND h.ano = '"& session("ano") &"'"
						strNoTope = strNoTope &" AND h.periodo = '"& session("periodo") &"'"
						strNoTope = strNoTope &" AND h.CodSede = '"& session("codsede") &"'"
						strNoTope = strNoTope &" and h.dia ='"& DiaNoTope &"'"
						strNoTope = strNoTope &" AND h.codmod= "& ModuloNoTope &""
						
						set RstNoTope = Session("Conn").execute(strNoTope)
						RamoFinal =""
						do while not RstNoTope.eof
							 RamoFinal = RamoFinal & RstNoTope("Ramo") & "<br>"
							 RstNoTope.movenext
						loop
						
					   MtrHorario(i, j) = RamoFinal
					   MtrHorSt(i,j)="P"
				   end if 
				   
                   Select case  MtrHorSt(i,j) 
                     Case "I": Color = "#00FF00" 
					 Leyenda = "Inscrita"
                     Case "E": Color = "#FF0000" 
					 Leyenda = "Tope"
                     Case "P": Color = "#FFCC66" 
					 Leyenda = "Preinscrita"
                     Case "T": Color = "#DBECF2" 
					 Leyenda = "Sin Horario"
                     Case "S": Color = "#FFCC66" 
					 Leyenda = "Preinscrita" 
                   End Select
              %>                
                 <td bgcolor="<%=color%>" width="78" height="17" onMouseOver="javascript:showtip(this, event, '<%=MtrHorRamo(i, j) + "</br>("+ Leyenda+")" %>')" onMouseout="javascript:hidetip()"><font face="Verdana" size="1"><%=MtrHorario(i, j) %> 
             <% end if%>
            </font></td>
          <% next %>
          </tr>
        <%
           RstModulos.Movenext 
           loop 
         %>    
          </table>
        </td>
      </tr>
      <tr > 
        <td colspan="2" height="10" align="left"> 
          <div align="center">          
            <p><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#666666">(Este 
              documento NO constituye certificado)</font></p>
          </div>
        </td>
      </tr>
      <tr>
      <td align="center" id="leyenda">
        <img src="Imagenes/Leyenda.bmp"></td>
      </tr>
    </table>
	</table>
</div>
<% 
If Trim(RC) = "S" then
%>
<script languaje="javascript">
  parent.Seleccion.location ="inscrip-asigna-combos-det.asp";
</script>
<%
end if
%>
</body>



<% If Trim(M) <> "" then
%>
<script languaje="javascript">
  alert('<%=Mensaje%>');
  window.location = "Horario.asp?P=S&T=S";
  //parent.Seleccion.location="asignatura-seccion.asp?Ramos=";
</script>
<%
end if
%>

<%ObjetosLocalizacion("horario.asp")%>
<script languaje="javascript">  
function imprime()
{	
  document.getElementById('imgPrint').style.display = 'none';	
  document.getElementById('leyenda').style.display = 'none';
  document.getElementById('tablaHorario').style.border = "1px solid #000";
  document.getElementById('tablaHorario').border = "1";
  parent.focus();
  parent.print(); 
  document.getElementById('imgPrint').style.display = 'block';
  document.getElementById('leyenda').style.display = 'block';
  document.getElementById('tablaHorario').style.border = "0px";
  document.getElementById('tablaHorario').border = "0";
}
</script>
</html>
<%
RstHorario.Close()
RstModulos.Close()
%>
<!--#INCLUDE file="include/desconexion.inc" -->
