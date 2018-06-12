<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	Response.Buffer = TRUE 
    Response.Expires = -1
	
	If Request.ServerVariables("HTTP_REFERER") = "" or Session("rut") = "" Then
		Response.Redirect("./error.html")
		Response.Clear()
		Response.End()
	End If
%>
<!--#INCLUDE VIRTUAL="/certificados/conexion.inc"  -->
<!--#INCLUDE VIRTUAL="/include/funciones.inc" -->
<%
	Dim numero_ecas

	If Not estacad(Session("rut")) Then 
		Response.Redirect("./wrn_201099.html")
	End If
	
	If Not estfinan(Session("rut")) Then
		Response.Redirect("./wrn_201096.html")
	End If
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Ecas | Solicitud de Certificados</title>
<style>
body {
	font-family: Arial, Helvetica, sans-serif;
	font-size: .9em;
}
.certificado {
	margin: 10px 50px;
}
input {
	width: 200px;
}
</style>

<body>
<%dim i
For Each i in Session.Contents
  Response.Write(i & ": "&Session.Contents(i)&"<br />")
Next%>

<%=Request.Form%>  
</body>
</html>
