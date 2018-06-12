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
	
	rut = Session("rut")
	fecha = Request.Form("fecha")
	certificado = Server.HTMLEncode(Request.Form("formato"))
	id_tipo = Request.Form("certificado")
	id_fin = Request.Form("fin")
	codigo_verificacion = Request.Form("codigo_verificacion")
	institucion = Request.Form("institucion")
	estacad_ = CStr(estacad_txt(rut))
	estfinan_ = CStr(estfinan_txt(rut))

	ano = ano_mat(rut)
	periodo = periodo_mat(rut)
	jornada_ = jornada(rut)
	nivel_ = nivel(rut)
	numero_ecas_ = numero_ecas(rut)
	validez_ = validez(id_tipo)
	
	correoelectronico = Request.Form("correoelectronico")

	sql = "INSERT INTO certificados_generados (rut, fecha, certificado, id_tipo, id_fin, codigo_verificacion, institucion, estacad, estfinan, ano, periodo, jornada, nivel, numero_ecas, validez) VALUES ('"&rut&"', '"&fecha&"', '"&certificado&"', "&id_tipo&", "&id_fin&", '"&codigo_verificacion&"', '"&institucion&"', '"&estacad_&"', '"&estfinan_&"', "&ano&", "&periodo&", '"&jornada_&"', "&nivel_&", '"&numero_ecas_&"', "&validez_&"); select TOP 1 rut from certificados_generados;"

	'Response.write(sql)
	'Response.end()

	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.Open sql,Conn
	'response.write(sql)
	'response.end()

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title></title>
</head>

<body>
<form name="frmGenerar" id="frmGenerar" action="http://www.ecasvirtual.cl/certificados/generar.php" method="post" target="_self"> 
    <input name="codigo_verificacion" type="hidden" value="<%Response.Write(Request.Form("codigo_verificacion"))%>" />
    <input name="formato" type="hidden" value="<%Response.Write(Server.HTMLEncode(Request.Form("formato")))%>" />
    <input name="correoelectronico" type="hidden" value="<%Response.Write(Request.Form("correoelectronico"))%>" />
    <input name="TimeZoneOffset" type="hidden" value="" />
</form>
<script type="text/javascript" language="JavaScript">
    <!--    
    var x = new Date();
    document.frmGenerar.TimeZoneOffset.value = x.getTimezoneOffset();
    document.frmGenerar.submit();
    //-->
</script>
</body>
</html>
