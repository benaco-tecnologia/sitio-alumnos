<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#INCLUDE FILE="include/funciones.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Reiniciar Clave</title>
</head>

<body style="margin: 20px;">
<%
If Request.Form("rut_alumno") <> "" and Request.Form("clave_administrador") <> "" AND Request.Form("clave_administrador") = "ar2405" Then
	servidorMy = "192.168.4.10"
	nombreUsuarioMy = "matricula"
	passwordMy = "dtb01s"
	baseDatosMy = "ecasvirtual"

	
	servidor = "192.168.1.194"
	nombreUsuario = "matricula"
	password = "dtb01s"
	baseDatos = "matricula"
	
	strDsn = "ecasvirtual"
	
	'If (Session("ConnMySQL") = "") Then
	'	strConnMySQL = "DSN=" & strDsn &  ";UID="&nombreUsuarioMy&";PWD="&passwordMy&"" 
	'	Set Session("ConnMySQL") = Server.CreateObject("ADODB.Connection")
	'	Session("ConnMySQL").Open(strConnMySQL)								
	'End If
	
	aConnectionString = "DRIVER={SQL Server}; SERVER="& servidor &"; DATABASE="& baseDatos &"; UID="& nombreUsuario &"; PWD="& password &""
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open aConnectionString
	
	Set rs = Conn.Execute("SELECT RUT FROM MT_ALUMNO WHERE RUT = '"&UCase(CStr(Request.Form("rut_alumno")))&"';")
	
	
	
	If (Not rs.Eof) Then
		rut_alumno = CStr(rs("RUT"))
		Set rs = Nothing
		Dim cincoDigitos : cincoDigitos = Mid(rut_alumno, 1, 5)
		Conn.Execute("UPDATE MT_WEBMAIL SET PASSWEB = '"&Encripta(cincoDigitos)&"', PASSCHANGED='S' WHERE RUT = '"&rut_alumno&"'")
		Set infoUsuario = Conn.Execute("SELECT nombre FROM MT_WEBMAIL WHERE RUT = '"&rut_alumno&"'")
		nombre = infoUsuario("nombre")
		'response.write("aa")
		'response.end
		
		strConnMySQL = "DSN=" & strDsn &  ";UID="&nombreUsuarioMy&";PWD="&passwordMy&"" 
		Set ConnMySQL = Server.CreateObject("ADODB.Connection")
		ConnMySQL.Open(strConnMySQL)	
		
		Set rs = ConnMySQL.Execute("UPDATE mdl_user SET password = MD5('"&cincoDigitos&"') WHERE username = '"&rut_alumno&"';")
		Set rs = Nothing
		'response.write("aa")
		'response.end
		ConnMySQL.Close
		Set ConnMySQL = Nothing
	
		Response.Write("<h4>La clave del alumno " &Replace(nombre, "_", " ")& " ha sido reiniciada a los primeros 5 d&iacute;gitos de su RUT en el Portal de Alumnos y en EcasVirtual.</h4>")
		%><br /><input name="aceptar" value="Aceptar" type="button" onclick="window.location='reiniciar-clave.asp'" /><%
	Else
	%>
		<h4>Los datos ingresados no coinciden con ning&uacute;n alumno en nuestra base de datos, favor ingrese nuevamente.</h4>
		<input name="aceptar" value="Volver" type="button" onclick="window.location='reiniciar-clave.asp'" />
	<%
	End If
Else
%>
<h3>Reiniciar Clave</h3>
<form action="reiniciar-clave.asp" method="post">
	<table width="600" border="0" cellpadding="5">
		<tr>
			<td>RUT Alumno</td>
			<td><input name="rut_alumno" type="text" id="rut_alumno" maxlength="8" autocomplete="off" />
			(sin puntos ni d&iacute;gito verificador)</td>
		</tr>
		<tr>
			<td>Clave Administrador</td>
			<td><input name="clave_administrador" type="password" id="clave_administrador" autocomplete="off" maxlength="20" />
			</td>
		</tr>
		<tr>
			<td>&nbsp;</td>
			<td><input name="btnAceptar" type="submit" id="btnAceptar" value="Reiniciar Clave" /></td>
		</tr>
	</table>
</form>
<%
End if
%>
</body>
</html>
