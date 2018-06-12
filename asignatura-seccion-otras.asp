<%
   Response.Buffer = false
   Response.Expires = -1
   'Response.Clear
%>
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
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

<script language="JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../../Desktop%20Folder/admalumnosx/imag/botones/volver-on.gif','imag/botones/aceptar-on.gif','imag/botones/volver-on.gif')">
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=profe name=profe>
</OBJECT> <OBJECT RUNAT=server PROGID=ADODB.Recordset id=seccion name=seccion>
</OBJECT> <html> <head> 
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<%
   Dim RamosSelec, Rst
   
   CodSede = Session("CodSede")
   Ano = AnoAdmision()
   Periodo =PeriodoAdmision()   
   CodCarr = Session("CodCarr")
   CodPestud = Session("CodPestud")
   Codcli = Session("Codcli") 
   Carrera = Request("CmbCarreras")
   	if Session("CodCli") = "" then
  	Response.Redirect("saltoinicio.htm")
	end if

   'Response.Write(Carrera)
   if Trim(Carrera) = "" then
     Carrera = Session("CodCarr")
   end if
%><body OnLoad="javascript:visualiza();" bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"> 
<table width="584" border="0" cellspacing="0" cellpadding="0" height="8">
  <tr> 
    <td valign="bottom" height="69"> 
      <div align="left"> 
        <table width="584" border="0" cellspacing="0" cellpadding="0" height="43">
          <tr> 
            <td width="336"><img name="fder_r1_c1" src="imag/f-der/f-der_r1_c1.gif" width="336" height="19" border="0"></td>
            <td width="248"><img name="fder_r1_c2" src="imag/f-der/f-der_r1_c2.gif" width="248" height="19" border="0"></td>
          </tr>
          <tr> 
            <td width="336"><img src="imag/f-der/otras-asignaturas.gif" width="336" height="28"></td>
            <td width="248"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="248" height="28">
                <param name="movie" value="swf/flecha.swf">
                <param name="quality" value="high">
                <embed src="swf/flecha.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="248" height="28">
                </embed> 
              </object></td>
          </tr>
          <tr> 
            <td width="336"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><a href="javascript:;" onMouseUp="MM_openBrWindow('text-otrasasignaturas.htm','otrasasignaturas','width=300,height=300')"><img src="imag/botones/instrucciones.gif" width="86" height="14" border="0"></a></font></td>
            <td width="248">
              <div align="right"><a href="inscrip-asigna.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','imag/botones/volver-on.gif',1)"><img src="imag/botones/volver-of.gif" width="145" height="19" name="Image1" border="0"></a></div>
            </td>
          </tr>
          <%
   ' Analizaremos si es una carrera de Bachillerato
   StrSql = "Select Distinct s.CodCarr, c.nombre_l "
   StrSql = StrSql & " From ra_optivo o, ra_seccio s, mt_carrer c "
   StrSql = StrSql & " Where o.CodSede = '" & CodSede & "' "
   StrSql = StrSql & " And o.CodCarr = '" & CodCarr & "' "
   StrSql = StrSql & " And o.CodSede = s.CodSede "
   StrSql = StrSql & " And o.CodRamo = s.CodRamo "
   'StrSql = StrSql & " And s.CodCarr = '" & CodCarr2 & "' "
   StrSql = StrSql & " And s.Ano = " & Ano 
   StrSql = StrSql & " And s.Periodo = '" & Periodo & "' "
   StrSql = StrSql & " And o.Ano = " & Ano 
   StrSql = StrSql & " And o.Periodo = '" & Periodo & "' "
   StrSql = StrSql & " And s.codCarr = c.CodCarr "
   StrSql = StrSql & " And (o.codOpt = '' or o.codOpt is null) "
   'Response.Write(StrSql) 
   
   if BCL_ADO(Strsql, Rst) then
     If trim(CodCarr2) = "" then
        CodCarr2 = Rst("CodCarr")
     end if
%>
          <%          
   end if 
%>
        </table>
        <font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font size="1">Listado 
        de asignaturas.</font></b><font size="1"> Para inscribirlas haz click 
        en la casilla &quot;inscribir&quot;</font></font></div>
    </td>
  </tr>
</table>
<form id=form1 name=form1 action="asignatura-seccion-otras.asp">
  <table width="584" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="317"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2">Listado 
        de Carreras:</font></b></font> 
        <select name="CmbCarreras" onChange="Javascript:document.form1.submit()">
          <% StrSql = "Select CodCarr, Nombre_L as nombre from mt_carrer Where Sede = '" & CodSede & "' "
       Set Rst = Session("Conn").Execute(StrSql)
       do while not Rst.eof
         if trim(Rst("CodCarr")) = Carrera then    %>
          <option value="<%=Rst("CodCarr")%>" selected><%=Rst("Nombre")%></option>
          <%   else  %>
          <option value="<%=Rst("CodCarr")%>"><%=Rst("Nombre")%></option>
          <%   end if 
    
         Rst.Movenext
       loop 
    %>
        </select>
      </td>
      <td width="244"> 
        <div align="right"></div>
      </td>
    </tr>
  </table>
  <table width="282" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td> 
        <table border="0" cellpadding="0" cellspacing="0" height="78" align="left" width="584">
          <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
          <tr valign="top"> 
            <td colspan="2" height="83"> 
              <table width="584" cellspacing="0" cellpadding="0" height="39" border="1" bordercolor="#FFFFFF">
                <tr bgcolor="4a5da1"> 
                  <td height="16" width="46" bgcolor="4a5da1"> 
                    <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">ver</font></b></font></div>
                  </td>
                  <td height="16" width="147"> 
                    <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">profesor</font></b></font></div>
                  </td>
                  <td height="16" width="36"> 
                    <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">s</font><font size="1">ecci&oacute;n</font></b></font></div>
                  </td>
                  <td height="16" width="102"> 
                    <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">horario</font></b></font></div>
                  </td>
                  <td height="16" width="138"> 
                    <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Ramo</font></b></font></div>
                  </td>
                  <td height="16" width="32"> 
                    <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1"> 
                      cupo</font></b></font></div>
                  </td>
                  <td height="16" width="47" valign="top"> 
                    <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Probar</font></b></font></div>
                  </td>
                </tr>
                <% ' Recorremos las secciones 
                 'Response.Write("Ramo = " & RamosSelec(i) & " Tipo = " & GetTipoRamo(RamosSelec(i), CodPestud) & " CodPestud = " & CodPestud )
                 'Response.End
                 StrSeccion = "Select CodRamo, CodSecc, CodProf from ra_seccio Where CodCarr = '" & Carrera & "' " 
                 StrSeccion = StrSeccion & " And Ano = " & ano & " And Periodo = " & Periodo
                 StrSeccion = StrSeccion & " Order by CodRamo, convert(numeric, CodSecc) "
                 
                 'Response.Write(StrSeccion)
                 Set RstSeccion = Session("Conn").Execute (StrSeccion)
                 do while not RstSeccion.eof
                 
              %>
                <tr bgcolor="ffc172"> 
                  <td width="46" height="16" align="center" bgcolor="ffc172"> 
                    <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="detalle%20asignatura.htm" target="_top">detalle</a></font></div>
                  </td>
                  <td width="147" height="16" align="center"> 
                    <div align="center"><font face="Verdana" size="1"><%=GetNombreProfe(RstSeccion("CodProf"))%> 
                      </font></div>
                  </td>
                  <td width="36" height="16" align="center"> 
                    <div align="center"><font face="Verdana" size="1"><%=RstSeccion("CodSecc")%></font></div>
                  </td>
                  <td width="102" height="16" align="center"> 
                    <div align="center"><font face="Verdana" size="1"><%=GetHorario(RstSeccion("CodRamo"), CodSede, RstSeccion("CodSecc"), Ano, Periodo)%></font></div>
                  </td>
                  <td width="138" height="16" align="center"> 
                    <div align="center"><font face="Verdana" size="0"><%=GetNombreRamo(RstSeccion("CodRamo"))%></font></div>
                  </td>
                  <td width="32" height="16" align="center"> 
                    <div align="center"><font face="Verdana" size="1"> 
                      <%%>
                      </font></div>
                  </td>
                  <td width="47" height="16" bgcolor="ffc172" valign="top"> 
                    <div align="center"><font face="Verdana" size="1"> 
                      <% if not EstaInscrito(Codcli, RstSeccion("CodRamo"), RstSeccion("CodSecc"), Ano, Periodo) and _
                            not EstaCursado(Codcli, RstSeccion("CodRamo"), RstSeccion("CodSecc"))  then %>
                      <input type="checkbox" name="<%=RstSeccion("CodRamo") & "-" & RstSeccion("CodSecc")%>" OnClick="javascript:GrabaSel(this, '<%=RstSeccion("CodRamo")%>', '<%=RstSeccion("CodSecc")%>')">
                      <% else 
                           if  EstaInscrito(Codcli, RstSeccion("CodRamo"), RstSeccion("CodSecc"), Ano, Periodo) then %>
                      Inscrito 
                      <% else%>
                      Cursado 
                      <% end if%>
                      <% end if %>
                      </font> </div>
                  </td>
                </tr>
                <%
                 RstSeccion.Movenext
               loop
              %>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr> 
      <td>&nbsp;</td>
    </tr>
  </table>
</form>
<script languaje="javascript">
function visualiza() {
  var x=window.open ("text-otrasasignaturas.htm","otrasasignaturas","width=200,height=200");
 }

function GrabaSel(Obj, Ramo, Seccion)
{
   //alert(Obj.name);
   if (Obj.checked)
     {parent.Horario.location = "PruebaRamo.asp?A=I&R=" + Ramo + "&S=" + Seccion;}
   else
     {parent.Horario.location = "PruebaRamo.asp?A=D&R=" + Ramo + "&S=" + Seccion;}
}

</script>
<font size="2" face="Verdana, Arial, Helvetica, sans-serif"></font> <font size="2" face="Verdana, Arial, Helvetica, sans-serif"></font> 
</body>
</html>
