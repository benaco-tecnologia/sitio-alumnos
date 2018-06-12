<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="md5.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title></title>
</head>

<body>
<% 
	servidorMy = "192.168.4.10"
	nombreUsuarioMy = "matricula"
	passwordMy = "dtb01s"
	baseDatosMy = "ecasvirtual"
	
	strDsn = "ecasvirtual"

	strConnMySQL = "DSN=" & strDsn &  ";UID="&nombreUsuarioMy&";PWD="&passwordMy&"" 
	Set objConnMySQL = Server.CreateObject("ADODB.Connection")
	objConnMySQL.Open(strConnMySQL)
	
	Set rs = objConnMySQL.Execute("SELECT fullname FROM mdl_course;")
	do until (rs.EOF)
		Response.Write(rs("fullname") & "<br />")
		rs.MoveNext
	loop
%>
</body>
</html>