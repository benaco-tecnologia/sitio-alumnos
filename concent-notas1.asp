<%  response.buffer = false 
    Response.Expires = -1
%>
<!-- saved from url=(0022)http://internet.e-mail -->
<!--<OBJECT RUNAT=server PROGID=ADODB.Recordset id=malla name=malla></OBJECT> -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--<OBJECT RUNAT=server PROGID=ADODB.Recordset id=Detalle name=Detalle > </OBJECT>-->
<%

dim strsql,strsql2
codcli=session("codcli")
'codcli=200011001
'response.Write(codcli)
'response.end
strsql="select distinct a.codsecc,a.ramoequiv AS CODRAMO,c.nombre,a.periodo,a.ano,a.np,a.nf,a.escala ,a.nej,a.ncert1,a.ncert2,a.ncert3,a.ncert4,a.ncert5,a.ncert6,a.ncert7,a.ncert8,a.ncert9,a.ncert10,a.asistencia,a.ne,a.ner,a.nfcontrol"
strsql=strsql & " from ra_nota a, ra_ramo c "
strsql=strsql & " where a.codcli='" & (codcli) & "' and "
'strsql=strsql & " a.codramo=b.codramo and "
strsql=strsql & " a.ramoequiv=c.codramo " 
'strsql=strsql & " b.codramo=c.codramo "
strsql=strsql & " group by a.codsecc,a.ramoequiv,c.nombre,a.periodo,a.ano,a.np,a.nf,a.escala,a.nej,a.ncert1,a.ncert2,a.ncert3,a.ncert4,a.ncert5,a.ncert6,a.ncert7,a.ncert8,a.ncert9,a.ncert10,a.asistencia,a.ne,ner,a.nfcontrol"
strsql= strsql & " order by a.ano, a.periodo  "

Set Rst = Session("Conn").Execute(StrSql)

%>
<html>
<head>
<title>f-der.gif</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
<link rel="stylesheet" href="file://///Gotzilla/uddweb/admalumnos/css/tex-normales.css" type="text/css">
<link href="css/tex-normales.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<div align="left"> 
  <table border="0" cellpadding="0" cellspacing="0" height="138" align="left" width="877">
    <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
    <tr> 
      <td width="583">&nbsp;</td>
      <td width="294"><img src="imag/f-der/no-ayuda.gif" width="248" height="28"></td>
    </tr>
    <tr> 
      <td colspan="2"><img src="imag/f-der/concentracion.gif" width="336" height="28"><img src="imag/f-der/no-ayuda.gif" width="248" height="28"></td>
    </tr>
    <tr valign="bottom"> 
      <td colspan="2" height="30" class="tex-normales"><font color="#333333" size="2" face="Arial, Helvetica, sans-serif">Es 
        el listado de tus asignaturas cursadas a la fecha.</font></td>
    </tr>
    <tr valign="top"> 
      <td colspan="2" height="8">&nbsp;</td>
    </tr>
    <tr valign="top"> 
      <td colspan="2" height="25"> <table width="915" cellspacing="0" cellpadding="0" height="40" border="1" bordercolor="#FFFFFF">
          <tr bgcolor="4a5da1"> 
            <td height="17" width="43" bgcolor="4a5da1"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Codigo 
                Ramo </font></b></font></div></td>
            <td width="101" bgcolor="4a5da1"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Nombre 
                del Ramo</font></b></font></div></td>
            <td height="17" width="47"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Per&iacute;odo</font></b></font></div></td>
            <td height="17" width="26"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">A&ntilde;o</font></b></font></div></td>
            <td height="17" width="58"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Nota 
                Controles </font></b></font></div></t>
              
            <td height="17" width="58"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Nota 
                Ejercicios </font></b></font></div></td>
            <td height="17" width="27"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Cert1</font></b></font></div></td>
            <td height="17" width="27"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Cert2</font></b></font></div></td>
            <td height="17" width="27"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Cert3</font></b></font></div></td>
            <td height="17" width="27"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Cert4</font></b></font></div></td>
            <td height="17" width="27"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Cert5</font></b></font></div></td>
            <td height="17" width="27"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Cert6</font></b></font></div></td>
            <td height="17" width="27"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Cert7</font></b></font></div></td>
            <td height="17" width="27"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Cert8</font></b></font></div></td>
            <td height="17" width="27"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Cert9</font></b></font></div></td>
            <td height="17" width="33"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Cert10</font></b></font></div></td>
            <td height="17" width="35"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Nota 
                Pres.</font></b></font></div></td>
            <td height="17" width="35"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Nota 
                Ex.</font></b></font></div></td>
            <td height="17" width="41"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Nota 
                Ex.Rep.</font></b></font></div></td>
            <td width="36"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Puntaje 
                Final</font></b></font></div></td>
            <td height="17" width="36"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Nota 
                Final</font></b></font></div></td>
            <td height="17" width="77"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Asistencia</font></b></font></div></td>
          </tr>
          <%do while not rst.eof %>
          <%
		    if not rst.eof then
		    codigoramo=rst("Codramo")
			Nombre=rst("Nombre")
   			'response.Write(codigoramo)
			'response.End			
			periodo=rst("periodo")
   			ano=rst("ano")
   			notacont=rst("nfcontrol")
   			notaejer=rst("nej")
   			ncert1=rst("ncert1")
   			ncert2=rst("ncert2")
   			ncert3=rst("ncert3")
   			ncert4=rst("ncert4")  
   			ncert5=rst("ncert5")  			
			Escala=rst("escala")  
			ncert5=rst("ncert5")  
			ncert7=rst("ncert7")
			ncert8=rst("ncert8")
			ncert9=rst("ncert9")
			ncert10=rst("ncert10")
			notaPres=rst("np")  
			notaEx=rst("ne")  
			notaExRep=rst("ner")  
			notafinal=rst("nf")  
			
   			Asistencia=rst("asistencia")  
'	  
'else
  
			end if
'rst.close
			%>
          <tr bgcolor="ffc172"> 
            <td width="43" height="1"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=codigoramo%></font></div></td>
            <td width="101" height="1"><font face="Arial, Helvetica, sans-serif" size="1"><%=Nombre%> 
              </font> <div align="center"></div></td>
            <td width="47" height="1"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=periodo%></font></div></td>
            <td width="26" height="1"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=ano%></font></div></td>
            <td width="58" height="1"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=notacont%></font></div></td>
            <td width="58" height="1"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=notaejer%></font></div></td>
            <td width="27" height="1"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=ncert1%></font></div></td>
            <td width="27" height="1"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=ncert2%></font></div></td>
            <td width="27" height="1"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=ncert3%></font></div></td>
            <td width="27" height="15"><div align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=ncert4%></font></div></td>
            <td width="27" height="1"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=ncert5%></font></div></td>
            <td width="27" height="1"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=ncert6%></font></div></td>
            <td width="27" height="1"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=ncert7%></font></div></td>
            <td width="27" height="1"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=ncert8%></font></div></td>
            <td width="27" height="1"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=ncert9%></font></div></td>
            <td width="33" height="1"><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=ncert10%></font></div></td>
            <td width="35" height="1"><div align="center"><font face="Arial, Helvetica, sans-serif" color="#2E31B8" size="1"><b><%=notaPres%></b></font></div></td>
            <td width="35" height="1"><div align="center"><font face="Arial, Helvetica, sans-serif" color="#2E31B8" size="1"><b><%=notaEx%></b></font></div></td>
            <td width="41" height="1"><div align="center"><font face="Arial, Helvetica, sans-serif" color="#2E31B8" size="1"><b><%=notaExRep%></b></font></div></td>
            <td width="36"><div align="center"><font face="Arial, Helvetica, sans-serif" color="#2E31B8" size="1"><b><%=notafinal%></b></font></div></td>
            <td width="36" height="1"><div align="center"><font face="Arial, Helvetica, sans-serif" color="#2E31B8" size="1"><b><%=Escala%></b></font></div></td>
            <td width="77" height="1"><div align="center"><font size="1" face="Arial, Helvetica, sans-serif"><%=Asistencia%></font></div></td>
            <%  rst.movenext %>
          </tr>
          <% loop%>
        </table></td>
    </tr>
    <tr valign="top"> 
      <td colspan="2" height="2"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#666666">(Este 
        documento NO constituye certificado)</font> </td>
    </tr>
  </table>
</div>
</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
