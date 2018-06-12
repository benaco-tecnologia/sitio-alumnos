<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ECAS › Validaci&oacute;n de Certificados</title>
</head>

<body>
<h3>ECAS &#8250; Validación de certificados
</h3>
<form id="form1" name="form1" method="post" action="">
  <table border="1" cellpadding="5">
    <tr>
      <th width="148" scope="row">Código de verificación</th>
      <td width="3"><label>
        <input name="codigo_verificacion" type="text" id="codigo_verificacion" size="20" maxlength="10" />
      </label></td>
    </tr>
    <tr>
      <th scope="row">&nbsp;</th>
      <td><label>
        <input type="submit" name="button" id="button" value="Aceptar" />
      </label></td>
    </tr>
  </table>
</form>
<p>El Certificado no es válido, el código ingresado no existe.</p>
<p>El Certificado es válido para presentarlo en {INSTITUCION}.</p>
<table border="1" cellpadding="5">
  <tr>
    <th align="left" scope="row">RUT</th>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <th align="left" scope="row">Alumno</th>
    <td>&nbsp;</td>
  </tr>
</table>
<p>&nbsp;</p>
</body>
</html>
