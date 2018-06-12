<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	Response.Buffer = TRUE 
    Response.Expires = -1
	
	If Request.QueryString("usrcod") = "" Then
		Response.Redirect("./error.html")
		Response.Clear()
		Response.End()
	Else
		Session("rut") = Request.QueryString("usrcod")
	End If
%>
<!--#INCLUDE VIRTUAL="/certificados/conexion.inc"  -->
<!--#INCLUDE VIRTUAL="/include/funciones.inc" -->
<%
	If (Not estacad(Session("rut")) or Not matricula_actual(Session("rut"))) And Not excepcionCertificado(Session("rut")) Then 
		Response.Redirect("./wrn_201099.html")
		Response.End()
	End If
	
	If deuda_actual(Session("rut")) And Not excepcionCertificado(Session("rut")) Then
		'Response.Redirect("./wrn_201096.html")
		'Response.End()
	End If
	
	Session("genero") = genero(Session("rut")) 
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
.info, .success, .warning, .error, .validation {
    border: 1px solid;
    margin: 10px 100px;
    padding:15px 10px 15px 50px;
    background-repeat: no-repeat;
    background-position: 10px center;
}
.warning {
    color: #9F6000;
    background-color: #FEEFB3;
    background-image: url('warning.png');
}

.info {
color: #00529B;
background-color: #BDE5F8;
background-image: url('info.png');
}

</style>
</head>

<body>
<form name="frmCertificado" method="post" action="./certificado.asp">
<input name="rut" type="hidden" value="<%Response.Write(Session("rut"))%>" />
<img name="" src="http://www.ecasvirtual.cl/images/header-1.jpg" style="width:100px;" /><h3>Solicitud de certificados</h3>
<%If Session("error_datos") = "verdadero" and Request.QueryString("err") <> "" Then%>
<div class="warning">Debes completar todos los datos.</div>
<%End If%>
<p>Verifique su correo electr&oacute;nico e ingrese el nombre completo de la instituci&oacute;n en donde presentar&aacute; el certificado.</p>
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
<input name="correoelectronico" type="text" value="<%Response.Write(correo(Session("rut")))%>" /><br />
    <label for="btnConfirmar"></label>
    <input type="submit" name="btnPrevisualizar" value="Aceptar" />&nbsp;
    <input type="reset" name="btnCerrar" value="Cancelar" onClick="window.close();" />
</form>
<p>&nbsp;</p>
<div class="info">El Certificado se generar&aacute; en pantalla con <strong>formato PDF</strong>.</div>
</body>
</html>
