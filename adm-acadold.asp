<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<%  response.buffer = false %>
<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if
%>
<link rel="stylesheet" href="css/tex-normales.css" type="text/css">

<script language="JavaScript" type="text/JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}
//-->
</script>
<body onLoad="MM_preloadImages('imag/botones/salir-on.gif','imag/f-der/ficha-on.gif','imag/f-der/concentracion-on.gif','imag/f-der/encuesta-on.gif')">
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=consulta name=consulta >
   </OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=consulta1 name=consulta > </OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=consulta2 name=consulta > </OBJECT>
<%
dim strSql,strSql1,strSql2
ano=session("ano")
codcarr=session("Codcarr")

strSql ="select texto_promocion " &_
        "From promocion_carrera " &_
		"where codcarr='" & codcarr & "' and promocion= '" & ano & "'"

strSql1 ="select texto_promocion " &_
        "From mt_parame "

strSql2 ="select texto_promocion " &_
        "From mt_carrer " &_
		"where codcarr='" & codcarr & "'"

consulta.Open strSql,Conn
consulta1.Open strSql1,Conn
consulta2.Open strSql2,Conn

GetPeriodoActivo

dim enlaces(2)
  enlaces(1) = "inscrip-asigna.htm"
  enlaces(2) = "solicitud.htm"
  if InscritoNormal() then
    enlaces(1)="resultado.htm" 
  end if
  if InscritoEspecial() then
    enlaces(2)="resultado.htm" 
  end if
  'Response.write("ID = " & SESSION("PER_ID"))
  if SESSION("PER_ID") = "-1" then
    enlaces(1)="" 
    enlaces(2)="" 
  end if

'response.write(strsql)
'response.end
%>
<html>
<head>
<title>Registro Académico en línea</title>
<meta http-equiv="Content-Type" content="text/html;">
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
<link rel="stylesheet" href="untitled.css" type="text/css">
<link rel="stylesheet" href="css/tex-normales.css" type="text/css">
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

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('imag/f-der/inscripcion-on.gif','imag/f-der/solicitud-on.gif','imag/f-der/malla-on.gif','imag/f-der/ficha-on.gif','imag/f-der/concentracion-on.gif','imag/botones/resumen-on.gif')">
<div align="left">
  <table border="0" cellpadding="0" cellspacing="0" height="293" align="left" width="584">
    <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
    <tr> 
      <td width="409"><img name="fder_r1_c1" src="imag/f-der/f-der_r1_c1.gif" width="336" height="19" border="0"></td>
      <td width="248"><img name="fder_r1_c2" src="imag/f-der/f-der_r1_c2.gif" width="248" height="19" border="0"></td>
    </tr>
    <tr> 
      <td width="409"><img name="fder_r2_c1" src="imag/f-der/registro.gif" width="336" height="28" border="0"></td>
      <td width="248"> <font color="#2166B1"><img src="imag/f-der/no-ayuda.gif" width="248" height="28"></font></td>
    </tr>
    <tr> 
      <td colspan="2" height="0">&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="2" height="10"> 
        <table width="584" height="1" cellpadding="0" cellspacing="0" border="0">
          <tr> 
            <td valign="top"> 
              <p class="tex-normales">Bienvenido al sistema de inscripci&oacute;n 
                de ramos de la Universidad Academia de Humanismo Cristiano<br>
                A trav&eacute;s de este sistema puedes realizar los procesos de 
                inscripci&oacute;n normal de asignaturas y proceso de solicitudes 
                especiales de inscripci&oacute;n de asignaturas. Adem&aacute;s 
                puedes tener informaci&oacute;n sobre el resultado de tu inscripci&oacute;n 
                de asignaturas y de tu situaci&oacute;n acad&eacute;mica. Si necesitas 
                ayuda por favor selecciona &quot;ayuda animada, y sigue las instrucciones 
                que se te indican.</p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr valign="top"> 
      <td colspan="2" height="12"> 
        <table width="584" height="1" cellpadding="0" cellspacing="0" border="0">
          <tr> 
            <td valign="top"> 
              <p>&nbsp;</p>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr valign="top"> 
       <% if not consulta1.eof then%>
      <td colspan="2" height="0" class="tex-normales"><span class="tex-normales"><font face="Arial, Helvetica, sans-serif"><i> 
        <font color="#2661C1"><%=Normaliza(consulta1("texto_promocion"))%></font> </i></font></span></td>
        <%end if %>
    </tr>
    <tr valign="top">
	 <% if not consulta2.eof then%>
      <td colspan="2" height="0" class="tex-normales"><font face="Arial, Helvetica, sans-serif" color="#2661C1"><%=Normaliza(consulta2("texto_promocion"))%></font></td>
	  <%end if%>
    </tr>
    <tr valign="top"> 
      <% if not consulta.eof then%>
      <td colspan="2" height="0" class="tex-normales"><span class="tex-normales"><font face="Arial, Helvetica, sans-serif"><i><font color="#2661C1"><%=Normaliza(consulta("texto_promocion"))%></font></i></font></span></td>
	  <%end if%>
    </tr>
    <tr valign="top"> 
      <td colspan="2" height="0">&nbsp;</td>
    </tr>
    <tr valign="top"> 
      <td colspan="2" height="21"> 
        <table width="584" border="0" cellspacing="0" cellpadding="0" height="21">
          <tr> 
            <%if trim(enlaces(1)) <> "" then %>
            <td width="175"><a href="<%=enlaces(1)%>" onMouseOver="MM_swapImage('Image1','','imag/f-der/inscripcion-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/inscripcion-of.gif" width="218" height="20" name="Image1" border="0"></a></td>
            <%else %>
            <td width="175"><img src="imag/f-der/inscripcion-of.gif" width="218" height="20" name="Image1" border="0"></td>
            <%end if%>
            <td rowspan="2" height="2" width="5">&nbsp;</td>
            <td rowspan="2" width="399" valign="top" class="tex-normales"> Este 
              proceso te permite realizar tu inscripci&oacute;n normal de asignaturas, 
              en los per&iacute;odos fijados para ello.</td>
          </tr>
          <tr> 
            <td width="175">&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
    <tr valign="top"> 
      <td colspan="2" height="21"> 
        <table width="584" border="0" cellspacing="0" cellpadding="0" height="21">
          <tr> 
            <%if trim(enlaces(2)) <> "" then %>
            <td width="175"><a href="<%=enlaces(2)%>" onMouseOver="MM_swapImage('Image2','','imag/f-der/solicitud-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/solicitud-of.gif" width="298" height="20" name="Image2" border="0"></a></td>
            <%else%>
            <td width="175"><img src="imag/f-der/solicitud-of.gif" width="298" height="20" name="Image2" border="0"></td>
            <%end if%>
            <td rowspan="2" height="2" width="5">&nbsp;</td>
            <td rowspan="2" width="399" valign="top" class="tex-normales"> Te 
              permite solicitar una inscripci&oacute;n especial. Debes usar esta 
              opci&oacute;n s&oacute;lo para las asignaturas que no puedas inscribir 
              a trav&eacute;s de la inscripci&oacute;n normal.</td>
          </tr>
          <tr> 
            <td width="175" valign="top">&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
    <tr valign="top"> 
      <td colspan="2" height="21"> 
        <table width="584" border="0" cellspacing="0" cellpadding="0" height="21">
          <tr> 
            <td width="175"><a href="resultado.htm" onMouseOver="MM_swapImage('Image3','','imag/botones/resumen-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/botones/resumen-of.gif" width="255" height="20" name="Image3" border="0"></a></td>
            <td rowspan="2" height="2" width="5">&nbsp;</td>
            <td rowspan="2" width="399" valign="top" class="tex-normales"> Te 
              permite ver el detalle de las asignaturas inscritas en este per&iacute;odo.</td>
          </tr>
          <tr> 
            <td width="175">&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
    <tr valign="top"> 
      <td colspan="2" height="21"> 
        <table width="584" border="0" cellspacing="0" cellpadding="0" height="15">
          <tr> 
            <td width="105"><a href="malla-curri.asp" onMouseOver="MM_swapImage('Image4','','imag/f-der/malla-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/malla-of.gif" width="114" height="20" name="Image4" border="0"></a></td>
            <td rowspan="2" height="2" width="5">&nbsp;</td>
            <td rowspan="2" width="451" valign="top" class="tex-normales"> Te 
              permite ver el plan de estudios de tu carrera y los ramos cursados.</td>
          </tr>
          <tr> 
            <td width="105">&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
    <tr valign="top"> 
      <td colspan="2" height="21"> <table width="584" border="0" cellspacing="0" cellpadding="0" height="17">
          <tr> 
            <td width="93"><a href="ficha.asp" onMouseOver="MM_swapImage('Image5','','imag/f-der/ficha-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/ficha-of.gif" width="114" height="20" name="Image5" border="0"></a></td>
            <td rowspan="2" height="2" width="5">&nbsp;</td>
            <td rowspan="2" width="451" valign="top" class="tex-normales"> Te 
              permite ver tu historia acad&eacute;mica.</td>
          </tr>
          <tr> 
            <td width="93">&nbsp;</td>
          </tr>
        </table>
        <table width="584" border="0" cellspacing="0" cellpadding="0" height="17">
          <tr> 
            <td width="93"><a href="concent-notas.asp" onMouseOver="MM_swapImage('Image51','','imag/f-der/concentracion-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/concentracion-of.gif" name="Image51" width="162" height="20" border="0" id="Image51"></a></td>
            <td rowspan="2" height="2" width="5">&nbsp;</td>
            <td rowspan="2" width="451" valign="top" class="tex-normales"> Te 
              permite ver tu concentracion y asistencia acomulada.</td>
          </tr>
          <tr> 
            <td width="93">&nbsp;</td>
          </tr>
        </table>
        <table width="584" border="0" cellspacing="0" cellpadding="0" height="17">
          <tr> 
            <td width="93"><a href="javascript:Enviar('ed.htm','<%=Encuestado(session("Codcli"))%>')" onMouseOver="MM_swapImage('Image52','','imag/f-der/encuesta-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/encuesta-of.gif" name="Image52" width="115" height="20" border="0" id="Image52"></a></td>
            <td rowspan="2" height="2" width="5">&nbsp;</td>
            <td rowspan="2" width="451" valign="top" class="tex-normales"> Te 
              permite evaluar a tus docentes.</td>
          </tr>
          <tr> 
            <td width="93">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr valign="top"> 
      <td colspan="2" height="21"> 
        <table width="584" border="0" cellspacing="0" cellpadding="0" height="21">
          <tr> 
            <td width="137"><a href="javascript:Terminar()"><img src="imag/botones/salir-of.gif" width="114" height="20" name="Image6" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image6','','imag/botones/salir-on.gif',1)" border="0"></a></td>
            <td rowspan="2" height="2" width="5">&nbsp;</td>
            <td rowspan="2" width="403" valign="top" class="tex-normales">&nbsp; </td>
          </tr>
          <tr> 
            <td width="137">&nbsp;</td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
  <p class="tex-normales">&nbsp;</p>
</div>
</body>

<script languaje="javascript">
function Terminar()
{
<% if session("logueado")="1" then %>
   if (confirm("Desea abandonar?") )
   {
      window.location = "salir.asp";
        //window.top.location = "alumn-udd.asp";
   }
<% end if%>   
}

function Enviar(pag,valor)
{
 if (valor=="True") {
    alert("Encuesta Realizada");
    return;
 }
  window.location.href = pag;	
}

</script>

</html>
