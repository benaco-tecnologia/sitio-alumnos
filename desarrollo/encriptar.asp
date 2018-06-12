<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#INCLUDE FILE="include/funciones.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Encripta</title>
</head>

<body>
<%
If Request.QueryString("usuario") <> "" Then
	Response.Write("Clave: " & Mid(request.QueryString("usuario"), 1, 5))
	Response.Write("<br />Encriptado: " & Encripta(Mid(Request.QueryString("usuario"), 1, 5)))
	''Response.Write("<br />Desencriptado: " & Desencripta(Mid(Request.QueryString("usuario"), 1, 5)))

	'Const servidor = "192.168.1.194"
	'Const nombreUsuario = "matricula"
	'Const password = "dtb01s"
	'Const baseDatos = "matricula"									
	
	'Set Conn = Server.CreateObject("ADODB.Connection")
	'Conn.Open "DRIVER={SQL Server}; SERVER="& servidor &"; DATABASE="& baseDatos &"; UID="& nombreUsuario &"; PWD="& password &""
	'Conn.Execute("UPDATE MT_WEBMAIL SET PASSWEB = '"&Encripta(Mid(Request.Form("campo"), 1, 5))&"', PASSCHANGED='S' WHERE RUT = '"&Request.Form("campo")&"'")
	'Response.write("UPDATE MT_WEBMAIL SET PASSWEB = '"&Encripta(Mid(Request.Form("campo"), 1, 5))&"', PASSCHANGED='S' WHERE RUT = '"&Request.Form("campo")&"'")
	'Conn.Close
	'set Conn = Nothing
	
	Response.write("UPDATE MT_WEBMAIL SET PASSWEB = '"&Encripta("15422")&"', PASSCHANGED='S' WHERE RUT = '"&Request.QueryString("usuario")&"'")
End if
%>
<form action="encriptar.asp" method="post">
RUT: <input name="campo" type="text" /><br />
Clave Admin: <input name="clave" type="password" /><br />
<input name="aceptar" value="Aceptar" type="submit" />
</form>
</body>
</html>
