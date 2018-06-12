<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
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
<body  leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#FFFFFF" onLoad="MM_preloadImages('../../../../Desktop%20Folder/admalumnosx/imag/botones/todas-las-asig-on.gif','../../../../Desktop%20Folder/admalumnosx/imag/botones/seleccionar-on.gif','../../../../Desktop%20Folder/admalumnosx/imag/botones/terminar-inscrip-on.gif')">
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
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"> 

<%' if not InscritoNormal() then %>
  <script>
  //   alert("No ha inscrito aún vía solicitud Normal ...");
  </script>   
<%'  end if%>
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
			<td width="825" align="center" valign="top">
			<% CargarTop1()%><% SubMenu()%>	
<div style="width:825PX; height:600px; overflow:scroll;" align="center">
  <form name="form1" method="post" action="asignatura-seccion.asp">
    <input type = "hidden" name = "Ramos" value = ""
    <table border="0" cellpadding="0" cellspacing="0" align="center" width="825">
      <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
      <tr valign="top"> 
	   <td colspan="2">&nbsp;</td>
        <td width="489">&nbsp;</td>
      </tr>
      <tr> 
        <td colspan="2"><img name="fder_r2_c1" src="Imagenes/T-resumen_inscripcion.gif" width="471" height="38" border="0"></td>
        <td width="489" align="center" ><a target="_top" href="menu_tomaderamos.asp">Volver a Pantalla Completa</a> </td>
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
          <table width="706" cellspacing="1" cellpadding="0" height="59" border="0" bordercolor="#FFFFFF">
            <td height="17" colspan="11"> <div align="left"> 
                <p><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#000000">Ud. 
                  Tiene inscritas las siguientes asignaturas:</font></p>
              </div></td>
            </tr>
            <tr background="Imagenes/fdo-cabecera-cel22.jpg"> 
              <td width="66" height="30"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">C&oacute;digo</font></b></font></div></td>
              <td width="193" height="30"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Asignatura</font></b></font></div></td>
              <td width="43" height="30"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">secci&oacute;n</font></b></font></div></td>
              <td width="46" height="30"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Horario</font></b></font></div></td>
              <td width="94" height="30"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Ramo</font></b></font></div></td>
              <td width="37" height="30"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Tipo</font></b></font></div></td>
              <td width="66" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Parcial 
                1</font></strong></td>
              <td width="69" height="30"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Parcial 
                2</font></strong></td>
              <td width="69" height="30"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Parcial 
                3</font></strong></td>
              <td width="67" height="30"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ex. 
                Final</font></strong></td>
              <td width="70" height="30"><strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Ex. 
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
            <tr bgcolor="#DBECF2" height="25"> 
              <td width="66" height="8" align="center"><font face="Verdana" size="1"><%=RamosDebe("CodRamo")%></font></td>
              <td width="193" height="0" align="center"><font face="Verdana" size="1"><%=RamosDebe("Nombre")%></font></td>
              <td width="43" height="0" align="center"><font face="Verdana" size="1"><%=CodSecc%></font></td>
              <td width="46" height="8" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=Horario%></font></td>
              <td width="94" height="16" align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=RamosDebe("NombreReal")%></font></td>
              <td width="37" height="16" align="center"><font size="1" face="Arial, Helvetica, sans-serif"> 
                <%if RamosDebe("Tipocurso")="T"  then
			                                                                                                     response.write("Teorico")
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
          <table width="707" cellspacing="0" cellpadding="0" height="59" border="0" bordercolor="#FFFFFF">
            <td height="17" colspan="6"> 
              <div align="left"> 
                <p><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#000000">Ud. 
                  solicit&oacute; las siguientes asignaturas:</font></p>
              </div>
            </td>
            </tr>
            <tr background="Imagenes/fdo-cabecera-cel22.jpg"> 
              <td width="40" height="30"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">C&oacute;digo</font></b></font></div>
              </td>
              <td width="165" height="30"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Asignatura</font></b></font></div>
              </td>
              <td width="49" height="30"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">secci&oacute;n</font></b></font></div>
              </td>
              			  <td width="62" height="30"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Horario</font></b></font></div>
              </td>
              <td width="46" height="30"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Situación</font></b></font></div>
              </td>
              <td width="109" height="30"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Ramo</font></b></font></div>
              </td>
              
			  <%if paso=1 then%>
			   <td width="49" height="30"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">secci&oacute;n Solicitada</font></b></font></div>
              </td>
			  <%end if%>
              
			  <%if paso=1 then
			  
				  if Ucase(RamosSolici("Estado")) = "A" then %> 
				  <td width="163" height="30"> 
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
		   
            <tr bgcolor="#DBECF2" height="25"> 
              <td width="40" height="8" align="center"><font face="Verdana" size="1"><%=RamosSolici("CodRamo")%></font></td>
              <td width="165" height="0" align="center"><font face="Verdana" size="1"><%=RamosSolici("Nombre")%></font></td>
              <td width="49" height="0" align="center"><font face="Verdana" size="1"><%=CodSecc%></font></td>
			  
              <td width="62" height="0" align="center"><font face="Verdana" size="1"><%=Horario%></font></td>
              <td width="46" height="16" align="center"><font face="Verdana" size="1">
              <%If RamosSolici("Estado")="A" then
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
				    <td width="163" height="16" align="center"><font size="1" face="Arial, Helvetica, sans-serif">
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
        </td>
      </tr>
    </table>
  </form>
</div>

			</td>
		  </tr>
	  </table>
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

</script>
</html>
<%
RamosDebe.Close()
RamosSolici.Close()
%>
<!--#INCLUDE file="include/desconexion.inc" -->
