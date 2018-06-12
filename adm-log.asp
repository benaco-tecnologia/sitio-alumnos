<html>
<head>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript"  SRC="include/validarut.js"></SCRIPT>
<title></title>
<meta http-equiv="Content-Type" content="text/html;">
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div align="left">
  <table border="0" cellpadding="0" cellspacing="0" height="231" align="left" width="584">
    <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
    <tr> 
      <td width="409"><img name="fder_r1_c1" src="imag/f-der/f-der_r1_c1.gif" width="336" height="60" border="0"></td>
      <td width="248"><img name="fder_r1_c2" src="imag/f-der/f-der_r1_c2.gif" width="248" height="60" border="0"></td>
    </tr>
    <tr> 
      <td width="409"><img name="fder_r2_c1" src="imag/f-der/f-der_r2_c1.gif" width="336" height="28" border="0"></td>
      <td width="248">
        <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="248" height="28">
          <param name="movie" value="swf/flecha.swf">
          <param name="quality" value="high">
          <embed src="swf/flecha.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="248" height="28"></embed></object></td>
    </tr>
    <tr> 
      <td colspan="2" height="52"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">El 
        Administrador Acad&eacute;mico tiene por objetivo poner a disposici&oacute;n 
        del alumnado, ocho &uacute;tiles herramientas que permitir&aacute;n hacer 
        m&aacute;s eficiente las distintas instancias administrativas generadas 
        por los distintos niveles del proceso acad&eacute;mico.<br>
        Esta secci&oacute;n require el uso de Login y Password del estudiante:</font></td>
    </tr>
    <tr> 
      <td colspan="2" height="5">&nbsp;</td>
    </tr>
    <form name="login" method="post" action="ValidaClave.asp" onSubmit="javascript:return ValidaDatos();">
	<input type="hidden" name="rut" value="">
	<input type="hidden" name="pin" value="">	
    <tr valign="top"> 
      <td colspan="2" height="100"> 
        <table width="584" height="15" cellpadding="0" cellspacing="0" border="0">
          <tr> 
            <td width="58" height="11">
              <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">CI:&nbsp;</font></div>
            </td>
            <td width="155" height="11"> 
              <input name="logrut" type="text" size="12" maxlength="10">
             </td>
            <td height="16" rowspan="5">&nbsp;</td>
          </tr>
          <tr> 
            <td width="58" height="1">&nbsp;</td>
            <td width="155" height="1">&nbsp; </td>
          </tr>
          <tr> 
            <td width="58" height="15">
              <div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Password:&nbsp;</font></div>
            </td>
            <td width="155"> 
              <input name="logclave" type="password" size="12" maxlength="10">
            </td>
          </tr>
          <tr> 
            <td width="58" height="1">&nbsp;</td>
            <td width="155" height="1"> 
              <div align="right"> </div>
            </td>
          </tr>
          <tr> 
            <td width="58" height="2">&nbsp;</td>
            <td width="155" height="2"> 
              <div align="right">
                <input type="submit" name="Submit" value="Aceptar">
              </div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    </form>
  </table>
</div>
</body>
</html>
