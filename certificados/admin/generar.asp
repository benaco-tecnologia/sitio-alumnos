<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	Response.Buffer = TRUE 
    Response.Expires = -1
%>
<!--#INCLUDE VIRTUAL="/certificados/conexion.inc"  -->
<!--#INCLUDE VIRTUAL="/include/funciones.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ECAS | Administración de certificados</title>
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

.info {
color: #00529B;
background-color: #BDE5F8;
background-image: url('info.png');
}

.info, .success, .warning, .error, .validation {
    border: 1px solid;
    margin: 10px 100px;
    padding:15px 10px 15px 50px;
    background-repeat: no-repeat;
    background-position: 10px center;
}
</style>
</head>

<body>
<p><a href="./generar.asp"><strong>Generar</strong></a> <a href="validar.asp">Validar</a> <a href="listar.asp">Listar</a> <a href="alumnos.asp">Alumnos</a> <a href="excepciones.asp">Excepciones</a></p>
<form name="frmCertificado" method="post" action="./certificado.asp">
<label>Rut</label>
    <select name="rut">
    <option value="0">Seleccionar</option>
<%
	StrSql = "select distinct a.rut, cl.dig from MT_ALUMNO a, MT_CTADOC c, MT_CLIENT cl where a.matriculado = 'S' and ((ano_mat = 2018 and a.periodo_mat = 1) or (ano_mat = 2017 and a.periodo_mat = 2)) and a.estacad = 'VIGENTE' and c.CODCLI = a.rut and cl.codcli = a.rut order by a.rut;"

	Set rs = Server.CreateObject("ADODB.RecordSet") 	
	rs.Open strSql, Conn
	do until rs.EOF
		Response.Write("<option value=""" & rs.Fields(0) & """>"&FormatNumber(rs.Fields(0),0)&"-"&rs.Fields(1)&"</option>" & vbCrLf)
		rs.MoveNext
	loop
	rs.Close()
	Set rs = Nothing
%>
    </select><br />
	<label>Certificado</label>
    <select name="certificado">
      <option value="0">Seleccionar</option>
      <option value="1">Alumno Regular</option>
  	</select><br />
    <label>Fin</label>
    <select name="fin">
    <option value="0">Seleccionar</option>
<%
	StrSql = "select id_fin, fin from certificados_fin order by fin;"
	Set rs = Server.CreateObject("ADODB.RecordSet") 	
	rs.Open strSql, Conn
	do until rs.EOF
		Response.Write("<option value=""" & rs.Fields(0) & """>"&rs.Fields(1)&"</option>" & vbCrLf)
		rs.MoveNext
	loop
	rs.Close()
	Set rs = Nothing
%>
  	</select><br />
 	<label>Institución</label>
    <input name="institucion" type="text" /><br />
    <label>Correo electrónico</label>
<input name="correoelectronico" type="text" /><br />
    <label for="btnConfirmar"></label>
    <input type="submit" name="btnPrevisualizar" value="Aceptar" />&nbsp;
    <input type="reset" name="btnCerrar" value="Cancelar" onClick="window.close();" />
</form>
<p>&nbsp;</p>
<div class="info">El Certificado se generar&aacute; en pantalla con <strong>formato PDF</strong>.</div>
</body>
</html>
