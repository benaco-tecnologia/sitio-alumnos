<html xmlns="http://www.w3.org/1999/xhtml">
<head><title>Envia datos</title></head>   
<body>
 
<h1>Envia Datos</h1>         
 
<table>
 <form action="http://89.140.231.220/alumnosnet/alumnos.asp" method="post" name="formulario">
<tr>
  <td>Usuario</td>
  <td><input type="text" name="user" value="ValorUsuarioEncriptado"></td>
</tr> 
<tr>
  <td>Clave</td>
  <td><input type="text" name="pass" value="ValorClaveEncriptado"></td>
</tr> 
<tr>
  <td>Carrera</td>
  <td><input type="text" name="codcarr" value="ValorCarreraEncriptado"></td>
</tr>
<tr>
  <td>P&aacute;gina</td>
  <td><input type="text" name="redirect" value="ValorPaginaEncriptado"></td>
</tr>
<tr>
  <td></td>
  <td><input type="submit" name="submit" value="Enviar"></td>
</tr>

</form>

</table>
</body>
</html>