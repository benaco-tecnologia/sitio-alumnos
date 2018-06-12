<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
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
	If (Request.Form("btnExcepcion") <> "") Then
		StrSql = "INSERT INTO certificados_excepciones (rut) VALUES ('"&Request.Form("rut")&"');"
		Conn.Execute(StrSql)
	End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>ECAS | Administraci&oacute;n de certificados</title>
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
<p><a href="./generar.asp">Generar</a> <a href="validar.asp">Validar</a> <a href="listar.asp">Listar</a> <a href="alumnos.asp">Alumnos</a> <a href="excepciones.asp"><strong>Excepciones</strong></a></p>
<p>Aqu&iacute; se listan todos los alumnos que, independiente de su estado acad&eacute;mico o financiero, podr&aacute;n solicitar certificados durante el plazo m&aacute;ximo de 7 d&iacute;as.</p>
<%
	StrSql = "SELECT rut, DATEDIFF(day, GETDATE(), fecha) FROM certificados_excepciones WHERE DATEDIFF(day, GETDATE(), fecha) < 8 AND DATEDIFF(day, GETDATE(), fecha) > 0;"
	Set rs = Conn.Execute(StrSql)
	If Not rs.EOF Then
%>
<table width="90%" border="0" align="center" cellpadding="1" cellspacing="0">
  <tr>
    <td><strong>Rut</strong></td>
    <td><strong>Validez excepci&oacute;n</strong></td>
  </tr>
<%
	do until rs.EOF
%>
  <tr style="font-size:.75em;">
    <td><%=rs.Fields(0)%></td>
    <td><%=rs.Fields(1)%> días</td>
  </tr>
<%
		rs.MoveNext
	loop
	rs.Close()
	Set rs = Nothing
%>
</table>
<%
	Else
		Response.Write("<p><strong>No existen alumnos con excepci&oacute;n.</strong></p>")
	End If
%>
<p>&nbsp;</p>
<form id="form1" name="form1" method="post" action="excepciones.asp">
<p>Ingresar excepci&oacute;n:</p>
	<label for="rut">Rut de alumno <span style="font-size: .8em;">(sin puntos ni d&iacute;gito verificador)</span></label>
	<input type="text" name="rut" id="rut" maxlength="8" /> <input name="btnExcepcion" type="submit" value="Agregar" />
</form>
<p>&nbsp; </p>
</body>
</html>
