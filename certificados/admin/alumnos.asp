<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	Response.Buffer = TRUE 
    Response.Expires = -1
	
	If Request.ServerVariables("HTTP_REFERER") = "" Then
		Response.Redirect("../v2/error.html")
		Response.Clear()
		Response.End()
	End If
%>
<!--#INCLUDE VIRTUAL="/certificados/conexion.inc"  -->
<!--#INCLUDE VIRTUAL="/include/funciones.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ECAS | Administración de certificados</title>
<style>
body {
	font-family: Arial, Helvetica, sans-serif;
	font-size: .9em;
}
label, input {
	display: block;
	float: left;
	width: 200px;
	margin-bottom: 10px;
}
label {
	width: 150px;
	text-align: right;
	padding-right: 10px;
	margin-top: 2px;
}
br {
	clear: left;
}
</style>

</head>

<body>
<p><a href="./generar.asp">Generar</a> <a href="validar.asp">Validar</a> <a href="listar.asp">Listar</a> <a href="alumnos.asp"><strong>Alumnos</strong></a> <a href="excepciones.asp">Excepciones</a></p>
<p>Filtros</p>
<p>
<form action="alumnos.asp" method="post">
  <label>Matriculado</label>
  <select name="matriculado">
    <option value="<> ''"<%If Request.Form("matriculado") = "<> ''" Then Response.Write(" selected=""selected""") End If %>>Todos</option>
    <option value="= 'S'"<%If Request.Form("matriculado") = "= 'S'" Then Response.Write(" selected=""selected""") End If %>>Sí</option>
    <option value="= 'N'"<%If Request.Form("matriculado") = "= 'N'" Then Response.Write(" selected=""selected""") End If %>>No</option>
  </select>
  <br />
  <label>Año matrícula</label>
  <select name="ano_mat">
    <option value="<> 0"<%If Request.Form("ano_mat") = "<> 0" Then Response.Write(" selected=""selected""") End If %>>Todos</option>
  	<option value="= 2015"<%If Request.Form("ano_mat") = "= 2015" Then Response.Write(" selected=""selected""") End If %>>2015</option>
	<option value="= 2014"<%If Request.Form("ano_mat") = "= 2014" Then Response.Write(" selected=""selected""") End If %>>2014</option>
    <option value="= 2013"<%If Request.Form("ano_mat") = "= 2013" Then Response.Write(" selected=""selected""") End If %>>2013</option>
  </select>
  <br />
  <label>Período matrícula</label>
  <select name="periodo_mat">
    <option value="<> 0"<%If Request.Form("periodo_mat") = "<> 0" Then Response.Write(" selected=""selected""") End If %>>Todos</option>
    <option value="= 1"<%If Request.Form("periodo_mat") = "= 1" Then Response.Write(" selected=""selected""") End If %>>Otoño</option>
    <option value="= 2"<%If Request.Form("periodo_mat") = "= 2" Then Response.Write(" selected=""selected""") End If %>>Primavera</option>
  </select>
  <br />
  <label>Estado académico</label>
  <select name="estacad">
    <option value="<> ''"<%If Request.Form("estacad") = "<> ''" Then Response.Write(" selected=""selected""") End If %>>Todos</option>
    <option value="= 'VIGENTE'"<%If Request.Form("estacad") = "= 'VIGENTE'" Then Response.Write(" selected=""selected""") End If %>>VIGENTE</option>
    <option value="= 'ELIMINADO'"<%If Request.Form("estacad") = "= 'ELIMINADO'" Then Response.Write(" selected=""selected""") End If %>>ELIMINADO</option>
    <option value="= 'SUSPENDIDO'"<%If Request.Form("estacad") = "= 'SUSPENDIDO'" Then Response.Write(" selected=""selected""") End If %>>SUSPENDIDO</option>
    <option value="= 'EGRESADO'"<%If Request.Form("estacad") = "= 'EGRESADO'" Then Response.Write(" selected=""selected""") End If %>>EGRESADO</option>
    <option value="= 'TITULADO'"<%If Request.Form("estacad") = "= 'TITULADO'" Then Response.Write(" selected=""selected""") End If %>>TITULADO</option>
  </select>  
  <br />
  <label>Deuda</label>
  <select name="deuda">  	
    <option value="--"<%If Request.Form("deuda") = "--" Then Response.Write(" selected=""selected""") End If %>>Todos</option>
    <option value=" and rut in"<%If Request.Form("deuda") = " and rut in" Then Response.Write(" selected=""selected""") End If %>>Si</option>
    <option value=" and rut not in"<%If Request.Form("deuda") = " and rut not in" Then Response.Write(" selected=""selected""") End If %>>No</option>
  </select>
  <br />
  <label></label>
  <input type="submit" name="btnFiltrar" id="btnFiltrar" value="Aceptar" /><br />
  </form>
</p>
<p>&nbsp;</p>
<%
	If Request.Form("btnFiltrar") <> "" Then
		StrSql = "select distinct rut, matriculado, ano_mat, periodo_mat, estacad from MT_ALUMNO where ano_mat "&Request.Form("ano_mat")&" and periodo_mat "&Request.Form("periodo_mat")&" and matriculado "&Request.Form("matriculado")&" and estacad "&Request.Form("estacad")&""&Request.Form("deuda")&" (select CODCLI from MT_CTADOC where (DATEDIFF(day, FECVEN, GETDATE()) > 60) AND (SALDO > 0) and datepart(yyyy, fecven) = " &year(now())& " and feccancel is null);"
		Set rs = Server.CreateObject("ADODB.RecordSet") 	
		rs.Open strSql, Conn,3,3	
%>
<p align="center">Mostrando <%=rs.recordcount%> alumnos.</p>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="0">
  <tr>
    <td><strong><a href="?order=rut">Rut</a></strong></td>
    <td align="center"><strong>Cuotas sin pago</strong></td>
    <td align="center"><strong>Matriculado</strong></td>
    <td align="center"><strong>Año<br />matrícula</strong></td>
    <td align="center"><strong>Período<br />
    matrícula</strong></td>
    <td align="center"><strong>Estado<br />académico</strong></td>
  </tr>
<%
		do until rs.EOF
%>
  <tr style="font-size:.75em;">
    <td><%=rs.Fields(0)%></td>
    <td align="center"><acronym style="cursor: help;"<%=cuotas(rs.Fields(0))%></acronym></td>
    <td align="center"><%=rs.Fields(1)%></td>
    <td align="center"><%=rs.Fields(2)%></td>
    <td align="center"><%=rs.Fields(3)%></td>
    <td align="center"><%=rs.Fields(4)%></td>
  </tr>
<%
			rs.MoveNext
		loop
		rs.Close()
		Set rs = Nothing
	End If
%>
</table>
</body>
</html>
