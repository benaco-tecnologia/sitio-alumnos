<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	Response.Buffer = TRUE 
    Response.Expires = -1
	
	If Request.ServerVariables("HTTP_REFERER") = "" Then
		Response.Redirect("../v2/error.html")
		Response.Clear()
		Response.End()
	End If
	
	rut = Request.Form("rut")
%>
<!--#INCLUDE VIRTUAL="/certificados/conexion.inc"  -->
<!--#INCLUDE VIRTUAL="/include/funciones.inc" -->
<%
	Dim numero_ecas_
	Dim sexo

	If Not estacad(rut) Then 
		Response.Redirect("../v2/wrn_201099.html")
	End If
	
	If Not estfinan(rut) Then
		Response.Redirect("../v2/wrn_201096.html")
	End If
	
	mes = array("", "Enero","Febrero","Marzo","Abril","Mayo","Junio", "Julio", "Agosto","Septiembre","Octubre","Noviembre","Diciembre")
	
	Set rs = Server.CreateObject("ADODB.RecordSet")
	certificado = Request.Form("certificado")
	fin = Request.Form("fin")
	StrSql = "select tipo, formato from certificados_tipo where id_tipo = " & certificado
	rs.Open strSql,Conn
	
	if not rs.EOF then
		titulo_certificado = rs.Fields(0)
		formato = rs.Fields(1)
	end if
	
	rs.Close()
	formato = "<h3 align=""center"">CERTIFICADO<br /></h3><p>&nbsp;</p>" & formato
	StrSql = "select fin, frase from certificados_fin where id_fin = " & fin
	rs.Open strSql,Conn
	
	if not rs.EOF then
		formato = Replace(formato, "{fin}", rs.Fields(1) & " <strong>" &rs.Fields(0)& "</strong>")
	end if
	
	rs.Close()
	formato = Replace(formato, "{institucion}", Request.Form("institucion"))
	StrSql = "select a.rut, c.dig, c.nombre + ' ' + c.paterno + ' ' + c.materno as nombre, a.codcli as numero_ecas, ano_mat as anio, c.sexo as sexo from mt_alumno a, mt_client c where a.rut = c.codcli and c.codcli = '" + rut + "';"
	rs.Open strSql,Conn
	sexo = ""
	For Each fField in RS.Fields
		If (fField.Name = "rut") Then
			formato = Replace(formato, "{"&fField.Name&"}", FormatNumber(RS(fField.Name), 0))
		Elseif (fField.Name = "numero_ecas") Then
			numero_ecas_ = RS(fField.Name)
		Elseif (fField.Name = "sexo") Then
			sexo = RS(fField.Name)
		Else
			formato = Replace(formato, "{"&fField.Name&"}", RS(fField.Name))
		End If
	Next
	
	rs.Close
	fecha = now()
	'formato = Replace(formato, "{anio}", "2011")
	formato = Replace(formato, "{fecha}", day(fecha) &" de "&mes(month(fecha))&" de "&year(fecha))
	
	consonantes = array("B","C","D","F","G","H","J","K","L","M","N","P","Q","R","S","T","V","W","X","Z")
	numeros = array("1","2","3","4","5","6","7","8","9")
	
	Randomize
	codigo_verificacion = ""
	For i = 1 to 5
		codigo_verificacion = codigo_verificacion & consonantes(int(Rnd * 19))
	Next

	For i = 1 to 3
		codigo_verificacion = codigo_verificacion & numeros(int(Rnd * 8))
	Next

	formato = formato & "<p>&nbsp;</p><p style=""font-size: .8em;"">Para verificar la integridad del presente certificado ingrese a <a href=""http://www.ecas.cl/"">www.ecas.cl</a> y utilice el código "& codigo_verificacion & ".</p>" 
	
	If sexo = "F" Then
		formato = Replace(formato, "{don}", "do&ntilde;a")
		formato = Replace(formato, "{alumno}", "alumna")
		formato = Replace(formato, "{matriculado}", "matriculada")
		formato = Replace(formato, "{interesado}", "de la interesada")
	Else
		formato = Replace(formato, "{don}", "don")
		formato = Replace(formato, "{alumno}", "alumno")
		formato = Replace(formato, "{matriculado}", "matriculado")
		formato = Replace(formato, "{interesado}", "del interesado")
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
</head>

<body>
<p><a href="./generar.asp"><strong>Generar</strong></a> <a href="validar.asp">Validar</a> <a href="listar.asp">Listar</a> <a href="alumnos.asp">Alumnos</a></p>
<form name="frmConfirmacion" action="./grabar.asp" method="post"> 
    <div class="certificado">
        <%Response.Write(formato)%>
        <input name="fin" type="hidden" value="<%Response.Write(Request.Form("fin"))%>" />
        <input name="certificado" type="hidden" value="<%Response.Write(Request.Form("certificado"))%>" />
        <input name="formato" type="hidden" value="<%Response.Write(Server.HTMLEncode(formato))%>" />
        <input name="codigo_verificacion" type="hidden" value="<%Response.Write(codigo_verificacion)%>" />
        <input name="fecha" type="hidden" value="<%Response.Write(fecha)%>" />
        <input name="correoelectronico" type="hidden" value="<%Response.Write(Request.Form("correoelectronico"))%>" />
        <input name="institucion" type="hidden" value="<%Response.Write(Request.Form("institucion"))%>" />
        <input name="rut" type="hidden" value="<%Response.Write(rut)%>" />
    <p align="center">    
        <input type="submit" name="btnConfirmar" value="Confirmar" />&nbsp;
        <input type="reset" name="btnVolver" value="Modificar" onClick="history.go(-1);" />
    </p>
	</div>
</form>

</body>
</html>
