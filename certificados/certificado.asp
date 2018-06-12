<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	Dim nombre, rut, estacad, estfinan, tipo_certificado, fin_certificado, institucion, formato, titulo_certificado
	
	Const servidor = "192.168.1.192"
	Const nombreUsuario = "matricula"
	Const password = "dtb01s"
	Const baseDatos = "matricula"
	
	Dim StrSql, rs
	
	'"Provider=SQLOLEDB; " _  & "Data Source=source_name; " _  & "Database=db_name; " _  & "UID=userid; " _  & "PWD=password;"
	Dim aConnectionString : aConnectionString = "DRIVER={SQL Server}; "_
  											  & "SERVER="& servidor &"; "_
											  & "DATABASE="& baseDatos &"; "_
											  & "UID="& nombreUsuario &"; "_
											  & "PWD="& password &""
									
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open aConnectionString
	Set rs = Server.CreateObject("ADODB.RecordSet")
	
	tipo_certificado = Request.Form("tipo_certificado")
	fin_certificado = Request.Form("fin_certificado")
	
	StrSql = "select tipo, formato from certificados_tipo where id_tipo = " & tipo_certificado
	
	rs.Open strSql,Conn
	if not rs.EOF then
		titulo_certificado = rs.Fields(0)
		formato = rs.Fields(1)
	end if
	rs.Close()
	
	formato = "<p align=""center"">CERTIFICADO<br /></p>" & formato
	
	StrSql = "select frase from certificados_fin where id_fin = " & fin_certificado
	
	rs.Open strSql,Conn
	if not rs.EOF then
		formato = Replace(formato, "{FIN}", rs.Fields(0))
	end if
	rs.Close()
	
	formato = Replace(formato, "{NOMBRE}", Request.Form("nombre"))
	formato = Replace(formato, "{RUT}", Request.Form("rut"))
	
	rut = Request.Form("rut")
	numero_ecas = Request.Form("numero_ecas")
	id_tipo = Request.Form("tipo_certificado")
	id_fin = Request.Form("fin_certificado")
	institucion = Request.Form("institucion")
	estacad = Request.Form("estacad")
	estfinan = Request.Form("estfinan")
	ano = Request.Form("ano")
	periodo = Request.Form("periodo")
	jornada = Request.Form("jornada")
	nivel = Request.Form("nivel")
	
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ECAS › Solicitud de Certificados</title>
<style>
	body {
		/*background: url(bg.jpg) center top repeat-x;*/
		font-family: Verdana;
	}
	
	table {
		width: 70%;
		background-color: #FFF;
		font-size: 80%;
	}
	
	input, select {
		padding: 3px;
		font-size: 100%;
	}
</style>
</head>

<body>
<table border="1" align="center" cellpadding="5" cellspacing="0">
  <tr>
    <td>
    <%Response.Write(formato)%>
	</td>
  </tr>
</table>
<form action="" method="post">
<p align="center">
	<input name="certificado" type="hidden" value="<%=Server.HTMLEncode(formato)%>" />
    <input value="<%Response.Write(rut)%>" type="hidden" name="rut" />
    <input value="<%Response.Write(nombre)%>" type="hidden" name="nombre" />
    <input value="<%Response.Write(estacad)%>" type="hidden" name="estacad" />
    <input value="<%Response.Write(estfinan)%>" type="hidden" name="estfinan" />
    <input value="<%Response.Write(nivel)%>" type="hidden" name="nivel" />
    <input value="<%Response.Write(jornada)%>" type="hidden" name="jornada" />
    <input value="<%Response.Write(ano)%>" type="hidden" name="ano" />
    <input value="<%Response.Write(periodo)%>" type="hidden" name="periodo" />
    <input value="<%Response.Write(numero_ecas)%>" type="hidden" name="numero_ecas" />
    <input name="btnGenerar" type="submit" value="Confirmar emisión" />
    <input name="" type="button" value="Modificar" />
</p>
</form>
</body>
</html>
