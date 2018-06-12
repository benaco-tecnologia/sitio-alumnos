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
<%
	Function validado(codigo_verificacion)
		Set rstmp = Conn.Execute("select fecha from certificados_validados where codigo_verificacion = '" & codigo_verificacion & "';")
		If not rstmp.eof Then
			validado = "<img src='../v2/success.png' width='25' height='25' title='Validado el "&rstmp.Fields(0)&"' />"
		Else
			validado = "<img src='../v2/error.png' width='25' height='25' />"
		End If
		rstmp.Close()
		Set rstmp = Nothing	
	End Function
%>
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
<p><a href="./generar.asp">Generar</a> <a href="validar.asp">Validar</a> <a href="listar.asp"><strong>Listar</strong></a> <a href="alumnos.asp">Alumnos</a> <a href="excepciones.asp">Excepciones</a></p>
<%
	order = "g.fecha desc"
	if (Request.QueryString("order") <> "") Then
		order = Request.QueryString("order")
	End if
	StrSql = "select g.rut, t.tipo, f.fin, g.fecha, g.codigo_verificacion from certificados_tipo t, certificados_fin f, certificados_generados g where g.id_tipo = t.id_tipo	and g.id_fin = f.id_fin order by "&order&";"
	Set rs = Server.CreateObject("ADODB.RecordSet") 	
	rs.Open strSql, Conn		
%>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="0">
  <tr>
    <td width="15"><strong>#</strong></td>
    <td><strong><a href="?order=g.rut">Rut</a></strong></td>
    <td><strong><a href="?order=t.tipo">Certificado</a></strong></td>
    <td><strong><a href="?order=f.fin">Fin</a></strong></td>
    <td><strong><a href="?order=g.fecha">Emisión</a></strong></td>
    <td width="75" align="center"><strong>Validado</strong></td>
    <td width="40" align="center"><strong>PDF</strong></td>
  </tr>
<%
	do until rs.EOF
%>
  <tr style="font-size:.75em;">
    <td>&nbsp;</td>
    <td><%=rs.Fields(0)%></td>
    <td><%=rs.Fields(1)%></td>
    <td><%=rs.Fields(2)%></td>
    <td><%=rs.Fields(3)%></td>
    <td align="center"><%Response.Write(validado(rs.Fields(4)))%></td>
    <td align="center"><a href="http://www.ecasvirtual.cl/certificados/Certificado Ecas (Cod. <%Response.Write(rs.Fields(4))%>).pdf" target="_blank"><img src="pdf_icon.gif" width="25" height="28" border="0" /></a></td>
  </tr>
<%
		rs.MoveNext
	loop
	rs.Close()
	Set rs = Nothing
%>
</table>
</body>
</html>
