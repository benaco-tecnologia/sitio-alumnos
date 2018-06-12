
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#INCLUDE FILE="include/conexion.inc" -->

<html>
<head>
<title>::Contador::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<br>
<b>Contador de Toma de Ramos Otoño 2011</b>
<br>
<table>
<tr>
  <td>Alumnos Conectados<td>
  <td><% response.write (application("uactivos"))%></td>
</tr>
<tr>
  <td>Inscritos WEB<td>
  <td><% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count ( * ) as valor from  sis_reg_inscripcion where fecha >= '20101231' ",Conn
if rst.recordcount > 0 then
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing
%></td>
</tr>
<tr>
  <td>Solicitudes Web<td>
  <td><%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count ( * ) as valor from  sis_reg_solicitud where fecha >= '20101231' ",Conn
if rst.recordcount > 0 then
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%></td>
</tr>
<tr>
  <td>Total Inscripción <td>
  <td><%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2011 and estacad='VIGENTE' and periodo_mat = 1",Conn
if rst.recordcount > 0 then
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%></td>
</tr>
<tr>
  <td>Encuestas realizadas<td>
  <td><%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count ( * ) as valor from mt_alumno where ano_ed=2010 and periodo_ed=2 and encdoc='S' ",Conn
if rst.recordcount > 0 then
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%></td>
</tr>
</table>

</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
