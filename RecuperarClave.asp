<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top12.asp" -->
<!--#INCLUDE file="f-izq.asp" -->

<%
	Session("RutCli")=""
	Session("RutCliente")= ""
	Session("MiClave")=""
	Session("ClaveEncriptada")=""
	Session("MiValor")=0
	Session("NomAlum") = ""
	Session("RutAlum") = ""
	Session("Rut") = ""
	Session("CodCli") = ""
	session("Codsede")=""
	session("Codcarr")=""
	Session("codpestud")=""
	Session("anoalumno") = ""
	Session("Jornada") = ""
	Session("Nivel") = ""
    Session("Logueado") = ""
	Session("BloqueoA") = ""
	Session("BloqueoF") = ""
	Session("estado") = ""
	Session("EstadoMatriculado") = ""
	session("Cerrada") = ""
	Session("GetCarr") = ""
	session("CarreraAlumno") = ""
	session("NumeroMaximoPrueba") = ""
	
	Session("destino")= Request("destino")

%>
<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title>Portal de Alumnos en L&iacute;nea </title>
<body >
<table width="1024" border="0" align="left" cellpadding="0" cellspacing="0">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0">
		  <tr>
			<td width="0" valign="top" align="left" ><% CargarTop()%><% CargarMenu()%></td>
		    <td valign="top" align="left" ><% CargarTop1()%>
		  
            
            
              <table align="center" width="100%">
            <tr></tr>
            <tr>
            <td></td>
            <td>
            <td valign="top" align="left"><img src="imagenes/titulos/T-cambio-password.gif" width="558" height="66"> <br>
                    <form name="claves" >
                      <table width="335" height="125" border="0" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
                        <tr align="center" valign="middle" bgcolor="#FFFFFF">
                          <td width="208" height="30" class="text-normal-celdas"><div align="right">Ingrese su password actual :</div></td>
                          <td width="163" height="30" valign="middle" background="imagenes/fondo-password.gif"><div align="center">
                              <input type="password" name="Clave" size="12" maxlength="9" class="casillas-form">
                          </div></td>
                        </tr>
                        <tr align="center" valign="middle">
                          <td height="30" class="text-normal-celdas"><div align="right">Ingrese su nueva password :</div></td>
                          <td height="30" background="imagenes/fondo-password.gif"><div align="center">
                              <input type="password" name="NuevaClave" size="12" maxlength="5" class="casillas-form">
                          </div></td>
                        </tr>
                        <tr align="center" valign="middle" bgcolor="#FFFFFF">
                          <td height="30" class="text-normal-celdas"><div align="right">Confirme su nueva password :</div></td>
                          <td height="30" background="imagenes/fondo-password.gif"><input type="password" name="RepClave" size="12" maxlength="5" class="casillas-form"></td>
                        </tr>
                        <tr align="center" valign="middle" bgcolor="#FFFFFF">
                          <td height="30" class="text-normal-celdas"><div align="right"></div></td>
              <td><div align="center"><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('_r11_c21','','Imagenes/botones/A-_r12_c2_f2alt.jpg',1)" >
                              <input type="hidden" name="CambioClave" value="<%=CambioClave%>">
                              <img src="Imagenes/botones/A-_r12_c2alt.jpg" name="_r11_c21" width="130" height="21" border="0" align="left" id="_r11_c21" onClick="javascript: Validar(window.document.claves.NuevaClave.value, window.document.claves.RepClave.value, window.document.claves.Clave.value)"></a></div></td>
                        </tr>
                      </table>
                      <div align="left"></div>
                      <div align="left"></div>
                </form></td>
            
            </td>
            <td></td>
            </tr>
            
            </table>
            
            
            </td>
           
		  </tr>
          <tr>
         
          </tr>
		</table>
	
	</td>
  </tr>
</table>
</body>
</html>

