<%
  Response.Expires = -1 
  Response.Buffer = true
  
  ' Un gran parametro que indica si funciona como solicitud
  SoloPrueba = REQUEST("P")
  If Trim(SoloPrueba = "") then
     SoloPrueba = "N"
  end if
  
%>
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="include/cupos.inc" -->
<script language="JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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
<body onLoad="MM_preloadImages('imag/botones/volver-on.gif','imag/botones/aceptar-on.gif')" bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=profe name=profe>
</OBJECT> <OBJECT RUNAT=server PROGID=ADODB.Recordset id=seccion name=seccion>
</OBJECT> <html> <head> 
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"></head> 
<%
   Dim RamosSelec
   dim var
   var=request("P")
   CodSede = Session("CodSede")
   Ano = Session("Ano") & ""
   Periodo = Session("Periodo") & ""
   CodCarr = Session("CodCarr")
   CodCarr2 = Request("CodCarr2")
   CodPestud = Session("CodPestud")
   Codcli = Session("Codcli") 
   If trim("CodCli") = "" then
     Response.Write("Debe comenzar de nuevo")
     Response.end
   end if
   ' Debemos separar los ramos
  
  Dim PrioridadRamos(30)
  Dim PrimeraVez, LastRstSeccion
  
  PrimeraVez = true
%>
<body OnLoad="javascript:visualiza();" bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" > 
<table width="584" border="0" cellspacing="0" cellpadding="0" height="43">
  <tr> 
    <td width="336"><img name="fder_r1_c1" src="imag/f-der/f-der_r1_c1.gif" width="336" height="19" border="0"></td>
    <td width="248"><img name="fder_r1_c2" src="imag/f-der/f-der_r1_c2.gif" width="248" height="19" border="0"></td>
    <td height="28" rowspan="2"> 
      <div align="right"></div>
    </td>
  </tr>
  <tr> 
  
   
   <%if var="S" then %>
      
       <td width="336"><img name="fder_r2_c1" src="imag/f-der/otras-asignaturas.gif" width="336" height="28" border="0"></td>
     <%else%>
    <td width="336"><img name="fder_r2_c1" src="imag/f-der/solicitud.gif" width="336" height="28" border="0"></td>
	
    <%end if%>
	<td width="248"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="248" height="28">
        <param name=movie value="solicitud-especial-fle.swf">
        <param name=quality value=high>
        <embed src="solicitud-especial-fle.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="248" height="28">
        </embed> 
      </object></td>
  </tr>
  <tr> 
   <%if var="S" then %>
		    <td width="336" valign="bottom" height="10" class="tex-normales"><a href="javascript:;" onMouseUp="MM_openBrWindow('text-otrasasignaturas.htm','Otrasasignaturas','width=400,height=450')"><img src="imag/botones/instrucciones.gif" width="86" height="14" border="0"></a><br>    
	<%else%>
	<td width="336" valign="bottom" height="10" class="tex-normales"><a href="javascript:;" onMouseUp="MM_openBrWindow('tex-inscripcionespecial.htm','SolicitudEspecial','width=400,height=450')"><img src="imag/botones/instrucciones.gif" width="86" height="14" border="0"></a><br>

   <%end if%>
      Para inscribirlas haz click en la casilla &quot;inscribir&quot;</td>
    <td width="248" height="10"> 
      <div align="right"><a href="javascript:ConfirmaSoli()" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','imag/botones/aceptar-on.gif',1)"><img src="imag/botones/aceptar-of.gif" width="156" height="19" name="Image1" border="0"></a></div>
    </td>
    <td height="10">&nbsp;</td>
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
  <tr> 
    <td class="tex-normales" valign="bottom" height="15">Carrera : </td>
    <td colspan=2 height="15"> 
      <form name=frmCarrera>
        <select name="CodCarr2" onChange="javascript:submit()">
          <% Do while not Rst.eof%>
          <%   If trim(CodCarr2) = Rst("CodCarr") then %>
          <option value=<%=Rst("CodCarr")%> selected><%=Rst("Nombre_l")%></option>
          <%   else %>
          <option value=<%=Rst("CodCarr")%>><%=Rst("Nombre_l")%></option>
          <%   end if 
             Rst.movenext
           loop 
           Rst.close()  
         %>
        </select>
      </form>
    </td>
  </tr>
  <%          
   end if 
%>
</table>
<%
 
   StrSql = " OtrosRamos '" & CodCli & "' "
   Set Rst = Conn.execute(StrSql)
   
   Do while not Rst.eof
 '  for i = 0 to ubound(RamosSelec, 1) - 1
      PrioridadRamos(i) = PrioridadRamo(Codcli, Rst("CodRamo"), Ano, Periodo)
%>
<% ' Recorremos las secciones 
                 'Response.Write("Ramo = " & RamosSelec(i) & " Tipo = " & GetTipoRamo(RamosSelec(i), CodPestud) & " CodPestud = " & CodPestud )
                 'Response.End
                 EsRepit = EsRepitente(Rst("CodRamo"), CodCli) 
                 TipoRamo = GetTipoRamo(Rst("CodRamo"), CodPestud)      
                
                 'Response.Write("Tipo Ramo = " & TipoRamo)
                 'Response.End 
                 select Case TipoRamo
                   case "OBLIGATORIO"
                     if EsPlanificado(Rst("CodRamo"), CodSede, Ano, Periodo) then
                       StrSeccion = "Select CodRamo, CodSecc, CodProf from ra_seccio Where CodRamo = '" & Rst("CodRamo") & "' and CodSede = '" & CodSede & "' " 
                       StrSeccion = StrSeccion & " And Ano = " & ano & " And Periodo = " & Periodo
                       StrSeccion = StrSeccion & " Order by convert(numeric, CodSecc) "
                     else
                       Equivalente = GetEquivalente(Rst("CodRamo"), CodSede, Ano, Periodo)
                       if Trim(Equivalente) <> "" then
                          StrSeccion = "Select CodRamo, CodSecc, CodProf from ra_seccio Where CodRamo = '" & Equivalente & "' and CodSede = '" & CodSede & "' " 
                          StrSeccion = StrSeccion & " And Ano = " & ano & " And Periodo = " & Periodo
                          StrSeccion = StrSeccion & " Order by convert(numeric, CodSecc) "                          
                       else
                          StrSeccion = "Select CodRamo from ra_ramo where Codramo is null "
                          ' Inconsistencia ....
                          'Response.Write("Inconsistencia en la Información ...")
                          'Response.End    
                       end if
                     end if 
                     EsEspecial = false
                   case "OPCIONAL"
                       StrSeccion = "Select distinct a.codramo,b.nombre,a.codsecc,a.cupo, e.CodProf " 
                       StrSeccion = StrSeccion & " from ra_optivo a,ra_ramo b,ra_seccio e" 
                       StrSeccion = StrSeccion & " where a.codcarr  = '" & CodCarr & "' " 
                       StrSeccion = StrSeccion & " and a.codramo=b.codramo" 
                       StrSeccion = StrSeccion & " and a.ano=" & Ano & "" 
                       StrSeccion = StrSeccion & " and a.periodo=" & Periodo & "" 
                       StrSeccion = StrSeccion & " and a.codramo=e.codramo" 
                       StrSeccion = StrSeccion & " and a.codsecc=e.codsecc" 
                       StrSeccion = StrSeccion & " and a.ano=e.ano" 
                       StrSeccion = StrSeccion & " and a.periodo=e.periodo" 
                       StrSeccion = StrSeccion & " and e.CodSede = '" & CodSede & "' "
                       StrSeccion = StrSeccion & " and a.codopt = '" & Rst("CodRamo") & "' "
                       EsEspecial = false
                       'Response.Write(StrSeccion)
                     
                   case "OPTATIVO ESPECIAL"  
                       StrSeccion = "Select distinct a.codramo,b.nombre,a.codsecc,a.cupo, e.CodProf" 
                       StrSeccion = StrSeccion & " from ra_optivo a,ra_ramo b,ra_seccio e" 
                       StrSeccion = StrSeccion & " where a.codcarr  = '" & CodCarr & "' " 
                       StrSeccion = StrSeccion & " and a.codramo=b.codramo" 
                       StrSeccion = StrSeccion & " and a.ano=" & Ano & "" 
                       StrSeccion = StrSeccion & " and a.periodo=" & Periodo & "" 
                       StrSeccion = StrSeccion & " and a.codramo=e.codramo" 
                       StrSeccion = StrSeccion & " and a.codsecc=e.codsecc" 
                       StrSeccion = StrSeccion & " and a.ano=e.ano" 
                       StrSeccion = StrSeccion & " and a.periodo=e.periodo" 
                       StrSeccion = StrSeccion & " and e.CodSede = '" & CodSede & "' "
                       StrSeccion = StrSeccion & " and a.CodSede = '" & CodSede & "' "
                       if trim(CodCarr2) <> "" then
                         StrSeccion = StrSeccion & " and e.CodCarr = '" & CodCarr2 & "' "
                       end if
                       StrSeccion = StrSeccion & " and (a.codopt is null or a.codopt='')"
                       EsEspecial = True
                       'Response.Write(StrSeccion)
                       'Response.end 
                 end select
                 
                 If EsEspecial then
                    if PrimeraVez then
                      Set RstSeccion = Conn.Execute (StrSeccion)
                      Set LastRstSeccion = RstSeccion
                      PrimeraVez = false
                    else
                      Set RstSeccion = LastRstSeccion
                      RstSeccion.movefirst
                    end if
                 else
                    Set RstSeccion = Conn.Execute (StrSeccion)
                 end if
                 'Set RstSeccion = Conn.Execute (StrSeccion)

                if not RstSeccion.eof then
%>
<form  id=<%=Rst("CodRamo")%>  name=<%=Rst("CodRamo")%> action="javascript:Nada()">
  <table width="282" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td> 
        <table border="0" cellpadding="0" cellspacing="0" height="78" align="left" width="584">
          <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
          <tr valign="top"> 
            <td colspan="2" height="83"> 
              <table width="584" cellspacing="0" cellpadding="0" height="19" border="1" bordercolor="#FFFFFF">
                <% If SoloPrueba = "N" then %>
                <tr bgcolor="83a3d0"> 
                  <td height="16" colspan="6"> 
                    <div align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2">Asignatura: 
                      <% =Rst("CodRamo") + " " + GetNombreRamo(Rst("CodRamo"))%>
                      </font></b></font></div>
                  </td>
                  <td height="16"> 
                    <div align="center"></div>
                  </td>
                </tr>
                <tr bgcolor="83a3d0"> 
                  <td height="16" colspan="1"> 
                    <div align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2">Motivo: 
                      </font></b></font></div>
                  </td>
                  <td height="16" colspan="6"> 
                    <input type="text" name="<%="T" & Rst("CodRamo")%>" size="90" maxlength="90" OnBlur="javascript:GrabaTexto(this)">
                  </td>
                </tr>
                <% end if%>
                <tr bgcolor="4a5da1"> 
                  <td height="16" width="46" bgcolor="4a5da1"> 
                    <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1" face="Geneva, Arial, Helvetica, san-serif">Lista 
                      Curso</font></b></font></div>
                  </td>
                  <td height="16" width="147"> 
                    <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1" face="Geneva, Arial, Helvetica, san-serif">profesor</font></b></font></div>
                  </td>
                  <td height="16" width="36" valign="middle"> 
                    <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1" face="Geneva, Arial, Helvetica, san-serif">secci&oacute;n</font></b></font></div>
                  </td>
                  <td height="16" width="102"> 
                    <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1" face="Geneva, Arial, Helvetica, san-serif">horario</font></b></font></div>
                  </td>
                  <td height="16" width="138"> 
                    <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1" face="Geneva, Arial, Helvetica, san-serif">Ramo</font></b></font></div>
                  </td>
                  <td height="16" width="32"> 
                    <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1"> 
                      <font face="Geneva, Arial, Helvetica, san-serif">cupo</font></font></b></font></div>
                  </td>
                  <td height="16" width="47">
                    <div align="center"><font face="Geneva, Arial, Helvetica, san-serif" size="1" color="#FFFFFF"><b>Solicitar</b></font></div>
                    </td>
                </tr>
                <%                 
                end if
                 do while not RstSeccion.eof
                 
                     TipoCupo = ""
                     Select Case TipoRamo
                       Case "OBLIGATORIO"
                          'Response.Write("Obligatorio")
                          Cupos = SacaCupoAntiguo ( RstSeccion("CodRamo"), RstSeccion("CodSecc"), Ano, Periodo, CodCarr, CodSede, EsRepit, TipoCupo )
                          'Response.Write("TipoCupo = " & TipoCupo)
                          'Response.End
                          Inscritos = CantidadInscritos ( RstSeccion("CodRamo"), RstSeccion("CodSecc"), CodSede, Ano, Periodo, TipoCupo )
                       Case "OPCIONAL"
                          'Response.Write("Opcional")
                          'Response.End
                          'Response.Write(SacaCupoOptativo(RamosSelec(i), RstSeccion("CodRamo"), RstSeccion("CodSecc"), Ano, Periodo, CodCarr) )
                          'Response.Write(Ano)
                          'Response.Write(Periodo)
                          'Response.end
                          Cupos = SacaCupoOptativo(Rst("CodRamo"), RstSeccion("CodRamo"), RstSeccion("CodSecc"), Ano, Periodo, CodCarr) 
                          Inscritos = CantidadInscritosOptativos(Rst("CodRamo"), RstSeccion("CodRamo"), RstSeccion("CodSecc"), CodSede, Ano, Periodo)
                       Case "OPTATIVO ESPECIAL"
                          'Response.Write("Opcional Especial")
                          'Response.End
                          'Cupos=0
                          Inscritos=0
                          Cupos = SacaCupoOptativoEspecial(RstSeccion("CodRamo"), RstSeccion("CodSecc"), Ano, Periodo, CodCarr) 
                          Inscritos = CantidadInscritosOptativoEspecial(RstSeccion("CodRamo"), RstSeccion("CodSecc"), CodSede, Ano, Periodo)
                     End Select

              %>
                <tr bgcolor="ffc172"> 
                  <td width="46" height="16" align="center" bgcolor="ffc172"> 
                    <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="detalle-asignatura.asp?R=<%=RstSeccion("CodRamo")%>&S=<%=RstSeccion("CodSecc")%> " target="_blank"><%=RstSeccion("CodRamo")%></a></font></div>
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
                      <%Response.Write("C=" & Cupos & ",I=" & Inscritos)%>
                      </font></div>
                  </td>
                  <td width="47" height="16" bgcolor="ffc172" valign="top"> 
                    <div align="center"><font face="Verdana" size="1"> 
                      <%
                         EstadoReal = EstadoRamoReal(CodCli, RstSeccion("CodRamo"))
                         
                         If (EstadoReal = "A") or (EstadoReal = "E") or (EstadoReal = "I") then
                           Cursado = True
                         else
                           Cursado = False
                         end if
                      %>
                      <% if EstaPreInscrito(Codcli, Rst("CodRamo"), RstSeccion("CodRamo"), RstSeccion("CodSecc"), Ano, Periodo) then %>
                      <input type="radio" name="<%=Rst("CodRamo")%>" value="<%=RstSeccion("CodRamo")%>" OnFocus="javascript:marca(this, '<%=trim(RstSeccion("CodRamo"))%>')" OnClick="javascript:GrabaSel(this,'<%=Rst("CodRamo")%>', '<%=RstSeccion("CodRamo")%>', '<%=RstSeccion("CodSecc")%>')" CHECKED <%If (Inscritos >= Cupos) or (cursado) then Response.Write("DISABLED") End if%> >
                      <% else %>
                      <input type="radio" name="<%=Rst("CodRamo")%>" value="<%=RstSeccion("CodRamo")%>" OnFocus="javascript:marca(this, '<%=trim(RstSeccion("CodRamo"))%>')" OnClick="javascript:GrabaSel(this, '<%=Rst("CodRamo")%>', '<%=RstSeccion("CodRamo")%>', '<%=RstSeccion("CodSecc")%>')" <%If  (Inscritos >= Cupos) or Cursado then Response.Write("DISABLED") End if%> >
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
  </table>
</form>
<%
    Rst.movenext
  loop 
%>
<script languaje="javascript">
function visualiza() {
  alert(P)
  var x=window.open ("tex-inscripcionespecial.htm","SolicitudEspecial","width=300,height=300");
 }
function marca(Obj, valor)
{
   if (Obj.checked) {Obj.value = "false"} 
   else {Obj.value = valor}
  //Obj.checked = !Obj.checked;
}

function RevisaOpcion(Obj, Ramo)
{ var j, k;
  //alert(Obj.value);
  for (j=0; j< document.forms.length; j++)
  {
     //if (document.forms.elements[i].type == "option")
     for (k=0; k < document.forms[j].elements.length ;k++)
     {
        if (document.forms[j].name != Obj.form.name)
        {
          //alert(document.forms[j].name); 
          if (document.forms[j].elements[k].type == "radio")
          { 
            //alert(Ramo + " : " + document.forms[j].elements[k].value);
            if ((Ramo == document.forms[j].elements[k].value) && (document.forms[j].elements[k].checked) )
            { alert("Ramo : " + document.forms[j].elements[k].value + " inscrito con " + document.forms[j].name ); 
              Obj.checked = false;
              return false;}
          }
        }         
     }
  }
  return true;
}


function GrabaTexto(Obj)
{
   var k, Ramo, Seccion, Curso, Texto;
   
   for (k = 0; k < Obj.form.elements.length; k++)
   {
      if (Obj.form.elements[k].type == "radio")
      {
         if (Obj.form.elements[k].checked)
         { 
            Ramo = Obj.form.elements[k].value;
            Curso = Obj.form.name;
            Seccion = "0";
            //alert("Curso = " + Curso);
            //parent.Horario.location = "GrabaSol.asp?A=I&R=" + Ramo + "&S=" + Seccion + "&T=" + Obj.value + "&C=" + Curso;            
         }
      } 
   }
}

function GrabaSel(Obj, Curso, Ramo, Seccion)
{ var j, k;

   if ((Obj.checked) && (Obj.value == "false") )
   {
      Obj.checked = false;
      //RamoInscrito[i] = false;
      <% if SoloPrueba = "N" then %>
      parent.Horario.location = "GrabaSol.asp?A=D&R=" + Ramo + "&S=" + Seccion + "&C=" + Curso;
      <% else %>
     parent.Horario.location = "PruebaRamo.asp?A=D&R=" + Ramo + "&S=" + Seccion;
      <% end if %>
   }   
   else {
   if ((Obj.checked) && (Obj.value != "false") )
     {
      <% if SoloPrueba = "N" then %>
       parent.Horario.location = "GrabaSol.asp?A=I&R=" + Ramo + "&S=" + Seccion + "&T=" + Obj.form.elements["T" + Curso].value + "&C=" + Curso;
      <% else %>
     parent.Horario.location = "PruebaRamo.asp?A=I&R=" + Ramo + "&S=" + Seccion;
      <% end if %>
     }
   else
     {
      <% if SoloPrueba = "N" then %>
       parent.Horario.location = "GrabaSol.asp?A=I&R=" + Ramo + "&S=" + Seccion + "&T=" + Obj.form.elements["T" + Curso].value + "&C=" + Curso;
      <% else %>
     parent.Horario.location = "PruebaRamo.asp?A=I&R=" + Ramo + "&S=" + Seccion;
      <% end if %>
     }
   }
}


function EscondeHor()
{
   //alert("paso");
   //parent.location = "asignatura-seccion.asp?Ramos="
   parent.Seleccion.resizeBy(1000, 1000);
}

function ConfirmaSoli()
{ 
  if (confirm("Ud solicitará las asignaturas de forma definitiva"))
  { 
     window.location =  "confirmacion-solici.asp" 
  }
}

function Nada()
{
}

</script></p>
</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->