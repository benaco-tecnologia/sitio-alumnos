<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->

<%

Dim StrSql,strRamosSolici
Dim Rst,Rstc
dim prueba	
Ano = 2003
Periodo = 2
CodCli =10087199

CodSede = Session("CodSede")
IF CodCli = "" then
   Response.Redirect "alumn-udd.htm"
end if
strRamosSolici = "SELECT b.codramo, b.nombre, a.inscrito, a.codsecc, a.Ramoequiv FROM ra_carga a, ra_ramo b " & _
               "WHERE a.codramo = b.codramo and " & _
               "a.inscrito = 'S' and " & _
               "a.codcli ='" & CodCli & "' and " & _
               "a.ano = '" & ano & "' and " & _
               "a.periodo = '" & periodo & "' Order By a.prioridad "

set Rstc=Session("Conn").execute(strRamosSolici)
%> 
	

<body bgcolor="#FFFFFF" text="#000000">
<table width="500" border="0">
  <tr>
    <td>Codigo</td>
    <td>Asignatura</td>
    <td>Seccion</td>
    <td>Horario</td>
  </tr>
  <tr>
<%
	  If Rstc.Eof Then
		  Else
		   While Not Rstc.Eof
		     if Ucase(Rstc("Inscrito")) = "S" then
		       CodSecc = Rstc("CodSecc")
		       Horario = GetHorario(Rstc("RamoEquiv"), CodSede, Rstc("CodSecc"), Ano, Periodo)
		     else
		       CodSecc = ""
		       Horario = ""
		     end if

%>
    <td><%=rstc("CodRamo")%></td>
    <td><%=rstc("nombre")%></td>
    <td><%=rstc("CodSecc")%></td>
    <td><%=Horario%></td>
<%
rstc.MoveNext
Wend
End If 
%>
  </tr>
</table>
</body>
</html> 
<!--#INCLUDE file="include/desconexion.inc" -->