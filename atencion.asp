<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%
Dim Error, Logueo, Pagina,msg

Error = Request("Error")
Logueo = Request("Logueo")
msg=request("msg")
msg=session("mensaje")
'response.Write(msg)
'response.End

If Error = 1 Then
   img = "Usuario no Existe"
   msg = "usuario no Existe"
End If   
If Error = 2 Then
   img = "Clave No Valida"
   msg = "Clave No Valida"
End If   
If Error = 3 Then
   img = "Bloqueo Financiero"
   msg = "Bloqueo Financiero"
End If   
If Error = 4 Then
   img = "No Matriculado"
   msg = "No Matriculado ..."
End If  
if error =5 Then

   Img= " Posee Bloqueo " & msg
   msg= " Posee Bloqueo " & msg
end if
if error =6 Then
   Img= " No Existe en Maestro Bloqueos "
   msg= " No Existe en Maestro Bloqueos "
end if
if error=7 then
   Img= "No Permiso Consulte al Administrador del Sitio "
   msg= msg
end if
If Logueo = 1 Then
   'Pagina = "adm-log.htm"
   Pagina = "alumnos.asp"
End If   
If Logueo = 2 Then
   Pagina = "CambioClave.asp"
End If   
If Logueo = 3 Then
   'Pagina = "adm-log.htm"
   Pagina = "alumnos.asp"
End If   

%>
<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html;">
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right"><% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td valign="top" >
			<% CargarTop1()%><% SubMenu()%>
				  <table border="0" cellpadding="0" cellspacing="15" height="84" width="584" align="left">
					<tr> 
					  <td width="584"><img name="fder_r2_c1" src="Imagenes/titulos/T-atencion.gif" width="271" height="38" border="0"></td>
					</tr>
					<tr>
					  <td height="1">&nbsp;</td>
					</tr>
					<tr> 
					  <td height="3" class="Tit-celdas"><span style="font-size: 18px"></span><font color="#DD0202" size="+1" face="Arial, Helvetica, sans-serif"><%=msg%></font></td>
					</tr>
					<tr valign="top"> 
					  <td height="0" align="rigth">&nbsp; <form>
						  <div align="left"> 
						  <!--<input type="button" name=Volver value="Volver" onClick="javascript:history.go(-1)" > -->
						  </div>
						</form></td>
					</tr>
			  </table>
			</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>
</body>
</html>
