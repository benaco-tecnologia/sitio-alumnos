<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
	Const servidor = "192.168.1.194"
	Const nombreUsuario = "matricula"
	Const password = "dtb01s"
	Const baseDatos = "matricula"									
	
	Set Session("Conn") = Server.CreateObject("ADODB.Connection")
	Session("Conn").Open "DRIVER={SQL Server}; SERVER="& servidor &"; DATABASE="& baseDatos &"; UID="& nombreUsuario &"; PWD="& password &""

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Alumnos / Sincronizaci&oacute;n U+ con Ecas Virtual</title>
</head>

<body>
</body>
</html>
