<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#INCLUDE FILE="include/funciones.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Desencripta</title>
</head>

<body>
<%
If Request.Form("campo") <> "" Then
	Response.Write("Desencriptado: " & Desencripta(Request.Form("campo")))
End if
%>
<form action="desencriptar.asp" method="post">
<input name="campo" type="text" /><br />

<input name="aceptar" value="Aceptar" type="submit" />
</form>
</body>
</html>