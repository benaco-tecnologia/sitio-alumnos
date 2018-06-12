<%  response.buffer = false 
    Response.Expires = -1
%>
<!-- saved from url=(0022)http://internet.e-mail -->
<!--<OBJECT RUNAT=server PROGID=ADODB.Recordset id=malla name=malla></OBJECT> -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/NombreRamo.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->

<%

dim strsql,strsql2
codcli=session("codcli")
RamoEquiv = Request("R")
Seccion = Request("S")
CodProf = request("P")
ano=session("ano")
Periodo=session("periodo")
CodSede=session("codsede")

strsql="Select count (*) as cifra from ra_carga c, mt_carrer cr "
strsql=strsql & " where c.Inscrito = 'S'"
strsql=strsql & " and c.ano = '" & ano & "'"
strsql=strsql & " and c.Periodo = '" & Periodo & "'"
strsql=strsql & " and c.RamoEquiv = '" & RamoEquiv & "'"
strsql=strsql & " and c.CodSecc = " & Seccion & ""
strsql=strsql & " and c.CodCarr = cr.codcarr "
strsql=strsql & " and cr.Sede = '" & CodSede & "' "
'strsql=strsql & "order by fecmod "
Set Rst = Session("Conn").Execute(StrSql)
if not rst.eof then
  num_total=rst("cifra")
else
  num_total=0
end if
rst.close

strsql="Select c.codcli, cl.Paterno, cl.Materno, cl.Nombre from ra_carga c, mt_carrer cr, mt_client cl, mt_alumno a "
strsql=strsql & " where c.Inscrito = 'S'"
strsql=strsql & " and c.ano = '" & ano & "'"
strsql=strsql & " and c.Periodo = '" & Periodo & "'"
strsql=strsql & " and c.RamoEquiv = '" & RamoEquiv & "'"
strsql=strsql & " and c.CodSecc = " & Seccion & ""
strsql=strsql & " and c.CodCarr = cr.codcarr "
strsql=strsql & " and cr.Sede = '" & CodSede & "' "
strsql=strsql & " and c.Codcli = a.CodCli "
strsql=strsql & " and a.Rut = cl.Codcli "

'strsql="Select codcli from ra_carga "
'strsql=strsql & " where Inscrito = 'S'"
'strsql=strsql & " and ano = '" & ano & "'"
'strsql=strsql & " and Periodo = '" & Periodo & "'"
'strsql=strsql & " and RamoEquiv = '" & RamoEquiv & "'"
'strsql=strsql & " and CodSecc = " & Seccion & ""

Set Rst = Session("Conn").Execute(StrSql)
'Do while not rst.eof
'   response.write(rst("codcli"))
'   rst.movenext
'loop
'response.end
		%>
<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/tablas.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.Estilo1 {color: #FFFFFF}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="400" cellspacing="0" cellpadding="0" height="135" border="1" bordercolor="#000000">
   <tr>
   <td height="30" colspan="2" align="center"valign="middle" background="Imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda" id="lblDetAsignatura">
      Detalles de la Asignatura </td> 
   </tr>
  <tr> 
    <td width="133" height="20" background="Imagenes/fdo-cabecera-cel22.jpg"> 
      <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><font color="#FFFFFF" id="lblAsignatura">Asignatura:</font>&nbsp;</font></b></div>    </td>
    <td height="20" bgcolor="#DBECF2" width="261"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;<%=GetNombreRamo(RamoEquiv)%></font></b></td>
  </tr>
  <tr bgcolor="4a5da1"> 
    <td width="133" height="20"  background="Imagenes/fdo-cabecera-cel22.jpg" > 
      <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><font color="#FFFFFF">Secci&oacute;n:&nbsp;</font></font></b></div>    </td>
    <td height="20" bgcolor="#DBECF2" width="261"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;<%=Seccion%> 
    </font></b></td>
  </tr>
  <tr bgcolor="ffc172"> 
    <td width="133" height="20" align="center" background="Imagenes/fdo-cabecera-cel22.jpg" > 
      <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><font color="#FFFFFF" id="lblProfesor">Profesor:</font>&nbsp;</font></b></div>    </td>
    <td height="20" align="left" bgcolor="#DBECF2" width="261"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;<%=(CodProf)%> 
    </font></b></td>
  </tr>
  <tr bgcolor="ffc172"> 
    <td width="133" height="20" align="center" background="Imagenes/fdo-cabecera-cel22.jpg" > 
      <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><font color="#FFFFFF">Horario:&nbsp;</font></font></b></div>    </td>
    <td height="20" align="left" bgcolor="#DBECF2" width="261"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;<%=GetHorario(RamoEquiv, CodSede, Seccion, Ano, Periodo)%> 
    </font></b></td>
  </tr>
  <tr bgcolor="ffc172"> 
    <td width="133" height="20" align="center" background="Imagenes/fdo-cabecera-cel22.jpg"> 
      <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><font color="#FFFFFF" id="lblAlumnosInscritos">Alumnos 
        inscritos:</font>&nbsp;</font></b></div>    </td>
    <td height="23" align="center" bgcolor="#DBECF2" width="261"> 
      <div align="left"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;<%=num_total%></font></b></div>    </td>
  </tr>
  <%
  if not rst.eof then
    Alumnos = rst.getrows
    pos = 0
    'do while not rst.eof	  
    do while pos <= Ubound(Alumnos, 2)
  %>
  <tr bgcolor="ffc172"> 
    <td width="133" height="20" align="center" background="Imagenes/fdo-cabecera-cel22.jpg"> 
      <div align="right"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><font color="#FFFFFF">Nombres&nbsp;</font></font></b></div>    </td>
    <td height="23" align="center" bgcolor="#ffe3e3" width="261"> 
      <div align="left"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp;&nbsp;<%=Alumnos(3,pos) + " " + Alumnos(2, pos) + " " + Alumnos(1, pos)%></font></b></div>    </td>
  </tr>  
  <%  
     'rst.movenext 
     pos = pos + 1
   loop
 end if
%>
  <!--<tr bgcolor="ffc172"> 
    <td height="1" align="center" bgcolor="ffe5c3" width="279">&nbsp;</td>
  </tr>
  <tr bgcolor="ffc172"> 
    <td height="2" align="center" bgcolor="ffe5c3" width="279">&nbsp;</td>
  </tr>
  <tr bgcolor="ffc172"> 
    <td height="2" align="center" bgcolor="ffe5c3" width="279">&nbsp;</td>
  </tr>
  <tr bgcolor="ffc172"> 
    <td height="2" align="center" bgcolor="ffe5c3" width="279">&nbsp;</td>
  </tr>
  <tr bgcolor="ffc172"> 
    <td height="2" align="center" bgcolor="ffe5c3" width="279">&nbsp;</td>
  </tr>
  <tr bgcolor="ffc172"> 
    <td height="2" align="center" bgcolor="ffe5c3" width="279">&nbsp;</td>
  </tr>
  <tr bgcolor="ffc172"> 
    <td height="2" align="center" bgcolor="ffe5c3" width="279">&nbsp;</td>
  </tr>
  <tr bgcolor="ffc172"> 
    <td height="2" align="center" bgcolor="ffe5c3" width="279">&nbsp;</td>
  </tr>
  <tr bgcolor="ffc172"> 
    <td height="2" align="center" bgcolor="ffe5c3" width="279">&nbsp;</td>
  </tr>
  <tr bgcolor="ffc172"> 
    <td height="2" align="center" bgcolor="ffe5c3" width="279">&nbsp;</td>
  </tr>
  <tr bgcolor="ffc172"> 
    <td height="2" align="center" bgcolor="ffe5c3" width="279">&nbsp;</td>
  </tr>
  <tr bgcolor="ffc172"> 
    <td height="2" align="center" bgcolor="ffe5c3" width="279">&nbsp;</td>
  </tr>
  <tr bgcolor="ffc172"> 
    <td height="2" align="center" bgcolor="ffe5c3" width="279">&nbsp;</td>
  </tr>
  <tr bgcolor="ffc172"> 
    <td height="2" align="center" bgcolor="ffe5c3" width="279">&nbsp;</td>
  </tr>-->
</table>
	<td height="2"  bgcolor="ffe5c3" width="279">	
		<form>
			<input type=button value="Cerrar" onClick="cerrarse()"> 
		</form> 
</td>
</body>
<%ObjetosLocalizacion("detalle-asignatura.asp")%>
</html>
<script> 
	function cerrarse()
	{ 
	window.close() 
	} 
</script>
<!--#INCLUDE file="include/desconexion.inc" -->