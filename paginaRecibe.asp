<%
Function desencripta_portal(valor)
dim str      
dim Rst

	str ="SELECT dbo.fn_desencripta_portal('"& valor &"')resultado"
	Set Rst = Session("Conn").Execute(str)
	if not Rst.eof then
		desencripta_portal = Rst("resultado")
	else
		desencripta_portal = ""
	end if
	Rst.close()

End Function
%>

<html xmlns="http://www.w3.org/1999/xhtml">
<head><title>Recibe datos</title></head>    
<body>
 
<h1>Recibe Datos</h1>    
 
<h1>Encriptados</h1>     
<table>
<tr>
  <td>Usuario</td>
  <td><%=request("user")%></td>
</tr> 
<tr>
  <td>Clave</td>
  <td><%=request("pass")%></td>
</tr> 
<tr>
  <td>Carrera</td>
  <td><%=request("codcarr")%></td>
</tr>
<tr>
  <td></td>
  <td></td>
</tr>
</table>

<h1>Desencriptados</h1>     
<table>
<tr>
  <td>Usuario</td>
  <td><%=desencripta_portal(request("user"))%></td>
</tr> 
<tr>
  <td>Clave</td>
  <td><%=desencripta_portal(request("pass"))%></td>
</tr> 
<tr>
  <td>Carrera</td>
  <td><%=desencripta_portal(request("codcarr"))%></td>
</tr>
<tr>
  <td></td>
  <td></td>
</tr>
</table>
</body>
</html>