<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	response.buffer = TRUE 
    Response.Expires = -1
	
	'If Session("rut") <> "" and Session("checado") <> "1" Then
	''	Response.Write("Ingreso prohibido")
		'Response.End()
	''End If
%>
<!--#INCLUDE VIRTUAL="/include/conexion.inc"  -->
<!--#INCLUDE VIRTUAL="/include/funciones.inc" -->
<%
	dim codcli
	dim rut
	dim nombre
	dim mail
	dim estacad
	dim estfinan
	
	codcli = Session("rut")
	codcli = "17464807"
	
	Const servidor = "192.168.1.8"
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
	
	StrSql = "select a.rut + '-' + c.dig as rut, c.nombre + ' ' + c.paterno + ' ' + c.materno, c.mail, a.estacad, a.estfinan, a.nivel, a.jornada, a.ano, a.periodo, a.codcli from mt_alumno a, mt_client c where a.rut = c.codcli and c.codcli = '" + codcli + "';"
	
	rs.Open strSql,Conn
	if not rs.EOF then
		rut = trim(rs.Fields(0))
		nombre = trim(rs.Fields(1))
		mail = trim(rs.Fields(2))
		estacad = trim(rs.Fields(3))
		estfinan = trim(rs.Fields(4))
		
		nivel = trim(rs.Fields(5))
		jornada = trim(rs.Fields(6))
		ano = trim(rs.Fields(7))
		periodo = trim(rs.Fields(8))
		numero_ecas = trim(rs.Fields(9))
		
	end if	
	rs.Close()
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ECAS &#8250; Solicitud de Certificados</title>
<style>
	body {
		/*background: url(bg.jpg) center top repeat-x;*/
		font-family: Verdana;
	}
	
	h3 {
		color: #000;
	}
	
	table {
		width: 95%;
		background-color: #FFF;
		font-size: 80%;
	}
	
	ul {
		font-size: 70%;
	}
	
	input, select {
		padding: 3px;
		font-size: 100%;
	}
</style>
</head>

<body>
<h3>ECAS › Solicitud de Certificados </h3>
<p>&nbsp;</p>
<% If estacad <> "VIGENTE" Then %>
<p>Estimado <%Response.Write(nombre)%>, su estado acad&eacute;mico no est&aacute; vigente.</p>
<% Elseif estfinan <> "AL DIA" Then %>
<p>Estimado <%Response.Write(nombre)%>, su situaci&oacute;n financiera no est&aacute; al d&iacute;a.</p>
<% Else %>
<form name="solicitud_certificado" method="post" action="certificado.asp">
<p>1. Estimado <%Response.Write(nombre)%> (<%Response.Write(rut)%>) verifique su correo electr&oacute;nico para recibir una copia del certificado.</p>
<p align="center"><input name="correo_electronico" type="text" size="30" value="<%Response.Write(mail)%>" />
<input value="<%Response.Write(rut)%>" type="hidden" name="rut" />
<input value="<%Response.Write(nombre)%>" type="hidden" name="nombre" />
<input value="<%Response.Write(estacad)%>" type="hidden" name="estacad" />
<input value="<%Response.Write(estfinan)%>" type="hidden" name="estfinan" />

<input value="<%Response.Write(nivel)%>" type="hidden" name="nivel" />
<input value="<%Response.Write(jornada)%>" type="hidden" name="jornada" />
<input value="<%Response.Write(ano)%>" type="hidden" name="ano" />
<input value="<%Response.Write(periodo)%>" type="hidden" name="periodo" />
<input value="<%Response.Write(numero_ecas)%>" type="hidden" name="numero_ecas" />
      
        </p>
<p>&nbsp;</p>
<p>2. Seleccione el tipo de certificado y para que lo necesita.</p>
<p align="center"><select name="tipo_certificado" style="margin-bottom: 2px;">
<%
	StrSql = "select id_tipo, tipo from certificados_tipo order by tipo;"
	
	rs.Open strSql,Conn
	do until rs.EOF
		Response.Write("<option value=""" & rs.Fields(0) & """>"&rs.Fields(1)&"</option>" & vbCrLf)
		rs.MoveNext
	loop
	rs.Close()
%>
</select>
<br />
<select name="fin_certificado">
<%
	StrSql = "select id_fin, fin from certificados_fin order by fin;"
	
	rs.Open strSql,Conn
	do until rs.EOF
		Response.Write("<option value=""" & rs.Fields(0) & """>"&rs.Fields(1)&"</option>" & vbCrLf)
		rs.MoveNext
	loop
	rs.Close()
%>
</select></p>
<p>&nbsp;</p>
<p>3. Ingrese el nombre completo de la institucion en donde presentar&aacute; el certificado.</p>
<p align="center"><input name="institucion" type="text" size="60" /></p>

<p align="center"><input type="submit" name="button" id="button" value="Aceptar" /> <input type="reset" name="btnCerrar" value="Cancelar" onClick="window.close();" />
  &nbsp;
  </p>
</form>
<% End If %>
</body>
</html>