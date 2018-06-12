<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=Datos name=Datos></OBJECT>
<%
' ****** Esta pag revisa la situación por la cual se impidió el logueo al alumno.

Dim Rut, strSql, Estado, GlosaError

Rut = RutSinDV(Session("RutAlum"))
strSql = "Select estacad from mt_alumno where rut = '" & Rut & "'"

Datos.Open strSql, Session("Conn")
'Estado = Trim(Datos(0))
Select Case Estado
	Case "ELIMINADO":
	     GlosaError = "Ud. se encuentra Eliminado de los registros de la  Universidad."
	Case "EGRESADO"	 
	     GlosaError = "Su estado actual es de Egresado en nuestros registros."
	Case "TITULADO"	 
	     GlosaError = "Su estado actual es de Titulado en nuestros registros."
	Case "SUSPENDIDO"	 		
	     GlosaError = "Ud. se encuentra suspendido de la Universidad."
	Case Else
		 GlosaError = "Ud. no se encuentra en los registros de la  Universidad."
End Select
Datos.Close()
'Conn.Close()


%>
<html>
<head>
<title>f-der.gif</title>
<meta http-equiv="Content-Type" content="text/html;">
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div align="left"></div>
<div align="left">
  <table border="0" cellpadding="0" cellspacing="0" height="58" width="584" align="left">
    <tr> 
      <td width="336"><img name="fder_r1_c1" src="imag/f-der/f-der_r1_c1.gif" width="336" height="60" border="0"></td>
      <td width="248"><img name="fder_r1_c2" src="imag/f-der/f-der_r1_c2.gif" width="248" height="60" border="0"></td>
    </tr>
    <tr> 
      <td width="336"><img name="fder_r2_c1" src="imag/f-der/atencion.gif" width="336" height="28" border="0"></td>
      <td width="248"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="248" height="28">
          <param name=movie value="../../escritorio/JobProgress%AA/UDD/www.udd.cl/swf/flecha.swf">
          <param name=quality value=high>
          <embed src="../../escritorio/JobProgress%AA/UDD/www.udd.cl/swf/flecha.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="248" height="28">
          </embed> 
        </object></td>
    </tr>
    <tr> 
      <td width="336">&nbsp;</td>
      <td width="248">&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="2" height="5"><font color="#FF0000" size="3" face="Verdana, Arial, Helvetica, sans-serif"><%=GlosaError%></font></td>
    </tr>
    <tr valign="top"> 
      <td colspan="2" height="0" align="rigth">&nbsp; <form><input type="button" name=Volver value="Aceptar" onClick="window.location.href='adm-log.htm';"></form></td>
    </tr>
  </table>
</div>
</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->