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
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#FFFFFF" onLoad="MM_preloadImages('../../../../Desktop%20Folder/admalumnosx/imag/botones/todas-las-asig-on.gif','../../../../Desktop%20Folder/admalumnosx/imag/botones/seleccionar-on.gif','../../../../Desktop%20Folder/admalumnosx/imag/botones/terminar-inscrip-on.gif')">
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=RamosDebe name=RamosDebe>
</OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=RamosPuede name=RamosPuede></OBJECT>
<%
Dim strRamosDebe, strRamosPuede, Rut, CodCli
CodCli = Session("CodCli")
CodSede = Session("CodSede")
IF CodCli = "" then
   Response.Redirect "alumn-udd.htm"
   'Response.Write( "Me voy : " + CodCli + " " + session("CodCli") )
end if
Ano = Session("Ano")
Periodo = Session("Periodo")
' desde la tabla mt_parame obtener el año y periodo
strRamosDebe = "SELECT b.codramo, b.nombre, a.inscrito, a.codsecc, a.Ramoequiv FROM ra_carga a, ra_ramo b " & _
               "WHERE a.codramo = b.codramo and " & _
               "a.Preinscrito = 'S' and " & _
               "a.codcli ='" & CodCli & "' and " & _
               "a.ano = '" & ano & "' and " & _
               "a.periodo = '" & periodo & "' Order By a.prioridad "
'agregar el año y periodo a esta query
	   
RamosDebe.Open strRamosDebe, Session("Conn")

%>
<Script Language="JavaScript">
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
</Script>
<html>
<head>
<title>f-der.gif</title>
<meta http-equiv="Content-Type" content="text/html;">
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div align="left">
<form name="form1" method="post" action="asignatura-seccion.asp">
  <input type = "hidden" name = "Ramos" value = "">
    <table border="0" cellpadding="0" cellspacing="0" align="left" width="530">
      <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
      <tr valign="top"> 
        <td colspan="2"><img name="fder_r1_c1" src="imag/f-der/f-der_r1_c1.gif" width="336" height="19" border="0"></td>
        <td width="579"><img name="fder_r1_c2" src="imag/f-der/f-der_r1_c2.gif" width="248" height="19" border="0"></td>
      </tr>
      <tr> 
        <td colspan="2"><img name="fder_r2_c1" src="imag/f-der/confirmacion.gif" width="336" height="28" border="0"></td>
        <td width="509"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="248" height="28">
            <param name="movie" value="swf/flecha.swf">
            <param name="quality" value="high">
            <embed src="swf/flecha.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="248" height="28">
            </embed> 
          </object></td>
      </tr>
      <tr valign="bottom"> 
        <td colspan="3"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Los 
          alumnos al inicio de cada semestre debe inscribir los ramos que clases 
          correspondientes a la carrera que cursas podr&aacute;s revisarlos permanentemente 
          las siguientes.</font></td>
      </tr>
      <tr valign="bottom"> 
        <td colspan="3">&nbsp;</td>
      </tr>
      <tr valign="bottom"> 
        <td colspan="2">&nbsp;</td>
        <td width="579">
<div align="right"><a href="javascript:ConfirmaInscrip();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image3','','imag/botones/terminar-inscrip-on.gif',1)"><img src="imag/botones/terminar-inscrip-of.gif" width="225" height="19" name="Image3" border="0"></a></div>
        </td>
      </tr>
      <tr valign="top"> 
        <td colspan="3" height="80"> 
          <table width="584" cellspacing="0" cellpadding="0" height="59" border="1" bordercolor="#FFFFFF">
            <td height="17" colspan="6"> 
              <div align="left"> 
                <p><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#000000">Ud. 
                  inscribir&aacute; las siguientes asignaturas:</font></p>
              </div>
            </td>
            </tr>
            <tr bgcolor="4a5da1"> 
              <td height="16" width="48" bgcolor="4a5da1"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">C&oacute;digo</font></b></font></div>
              </td>
              <td height="16" width="286"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Asignatura</font></b></font></div>
              </td>
              <td height="16" width="65"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">secci&oacute;n</font></b></font></div>
              </td>
              <td height="16" width="91"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Horario</font></b></font></div>
              </td>
              <td height="16" width="45"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Curso</font></b></font></div>
              </td>
            </tr>
            <% 
		  If RamosDebe.Eof Then
		  %>
            <%
		  Else
		   While Not RamosDebe.Eof
		     if Ucase(RamosDebe("Inscrito")) = "S" then
		       CodSecc = RamosDebe("CodSecc")
		       Horario = GetHorario(RamosDebe("RamoEquiv"), CodSede, RamosDebe("CodSecc"), Ano, Periodo)
		     else
		       CodSecc = ""
		       Horario = ""
		     end if
		   %>
            <tr bgcolor="ffc172"> 
              <td width="48" height="8" align="center"><font face="Verdana" size="1"><%=RamosDebe("CodRamo")%></font></td>
              <td width="286" height="0" align="center"><font face="Verdana" size="1"><%=RamosDebe("Nombre")%></font></td>
              <td width="65" height="0" align="center"><font face="Verdana" size="1"><%=CodSecc%></font></td>
              <td width="91" height="0" align="center"><font face="Verdana" size="1"><%=Horario%></font></td>
              <td width="45" height="0" align="center"><font face="Verdana" size="1">&nbsp; 
                <%%>
                </font></td>
            </tr>
            <%
		     RamosDebe.MoveNext
		   Wend
		 End If %>
          </table>
        </td>
      </tr>
    </table>
 </form>

</div>
</body>
<script languaje = "javascript">
function ConfirmaInscrip()
{
  document.form1.action = "confirmacion.asp"
 // alert("Hola");
  document.form1.submit();
} 

</script>
</html>
<%
RamosDebe.Close()
'Conn.Close()
%>
<!--#INCLUDE file="include/desconexion.inc" -->
