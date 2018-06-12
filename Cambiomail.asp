<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenuExterno.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->

<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if
%>
<html>

<script language="JavaScript">
function enviar(){
	var email = document.getElementById("mail");	
	if (email.value != "")
	{
		window.document.mail.action = "cambiomail_graba.asp";
		window.document.mail.submit();
	}
	else
		alert("Debe ingresar un valor.");
}

function validarEmail(email, campo) {
	 if (email.value != "")
	 {
		 expr = /^([a-zA-Z0-9_\.\-])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
		 if (!expr.test(email.value)){
			 alert("El valor ingresado " + email.value + " es incorrecto.");
			 document.getElementById(campo).value = "";
			 document.getElementById(campo).focus();
		 }
	 }
}

</script>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1"> 
<title>Portal de Alumnos en L&iacute;nea</title>
<body >
<table border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right" ><% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td valign="top">
			<% CargarTop1()%><% SubMenu()%>
			<table width="800" border="0" cellspacing="0" cellpadding="15" height="550" bgcolor="#FFFFFF">
			  <tr> 
                <td valign="top" align="left"><span style="font-size: 14px"><p style="font-size: 25px" class="text-menu">Actualizaci&oacute;n de correo electr&oacute;nico</p></span> <br><br>
                <form name="mail">
                    <table width="544" height="120" border="0" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
                        <tr align="center" valign="middle"	bgcolor="#FFFFFF">
                            <td width="100" height="55" class="text-normal-celdas">Correo electr&oacute;nico</td>
                            <td align="center"><table width="256"><tr>
                            <td  width="248" height="30" valign="middle" align="left" background="imagenes/fondo3.gif">
                            &nbsp;&nbsp;&nbsp;<input type="text" name="mail" id="mail" size="30" border="0" style="background-color:transparent;"  
                            maxlength="100" class="casillas-form" onblur="return validarEmail(this,'mail')">
                            </td>
                            </tr></table></td>
                            <td width="134" height="55" class="text-normal-celdas"><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('_r11_c21','','Imagenes/botones/A-_r12_c2_f2alt.jpg',1)" >
                            <img src="Imagenes/botones/A-_r12_c2alt.jpg" name="_r11_c21" width="130" height="21" border="0" align="center" id="_r11_c21" onClick="javascript:enviar()"></a></td>
                        </tr>
                        <tr align="center" valign="middle"	bgcolor="#FFFFFF">
                            <td width="100" height="55" class="text-normal-celdas"></td>
                            <td align="center"><table width="256"><tr>
                            <td  width="248" height="30" valign="middle" align="left">
                            </td>
                            </tr></table></td>
                            <td width="134" height="55" class="text-normal-celdas"></td>
                        </tr>
                        
                    </table>
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
<!--#INCLUDE file="include/desconexion.inc" -->