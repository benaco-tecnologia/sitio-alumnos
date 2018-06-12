<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if
%>
<script type="text/javascript">
function submitform()
{
	RbMatriculaBanco = document.getElementById('RbMatriculaBanco').checked;
	RbMatriculaSevipag = document.getElementById('RbMatriculaSevipag').checked;
	
	RbArancelBanco = document.getElementById('RbArancelBanco').checked;
	RbArancelSevipag = document.getElementById('RbArancelSevipag').checked;
	RbArancelPagare = document.getElementById('RbArancelPagare').checked;
	
	if (RbMatriculaBanco == false && RbMatriculaSevipag == false)
	{
		alert("Debe Seleccionar una opci\u00f3n de pago")
	}
	else if (RbArancelBanco == false && RbArancelSevipag == false && RbArancelPagare == false)
	{
		alert("Debe Seleccionar una opci\u00f3n de pago")
	}
	else
	{
		document.myform.submit();
	}	
}
</script>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/css/parrafo.css" rel="stylesheet" type="text/css">
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<body >
<table border="0" cellpadding="0" cellspacing="0"  align="left">
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
				  <table width="100%" border="0" cellspacing="0" cellpadding="15" height="550" bgcolor="#FFFFFF">
			  <tr> 
				<td valign="top"><p style="font-size:25px" class="text-menu">Matr&iacute;cula y Arancel.</p>
				  <table width="800" border="0" cellpadding="0" cellspacing="0">
                  <form name="myform" id="myform" action="Procesaarancelymatricula.asp" method="POST">
                   <tr valign="top" bgcolor="#FFFFFF">
					  <td height="24" colspan="4">
					     <b class="text-menu">Seleccione un medio de pago.</b>
					  </td>
					</tr>  
					<tr valign="top" bgcolor="#FFFFFF">
					  <td width="99" height="26">
					    <p class="text-menu">&nbsp;</p>
					  </td>
					  <td width="165">&nbsp;</td>
				      <td width="184">&nbsp;</td>
					</tr>
                    <tr valign="top" bgcolor="#FFFFFF">
					  <td height="24">
					    <p class="text-menu">Matr&iacute;cula</p>
					  </td>
				<td><input id ="RbMatriculaBanco" type="radio" name="rbMatricula" value="BANCO" checked><b class="Tit-celdas">Banco (Contado).</b></td>
				<td><input id ="RbMatriculaSevipag" type="radio" name="rbMatricula" value="SERVIPAG"><b class="Tit-celdas">Servipag (Contado).</b></td>                      
					</tr>
                    <tr valign="top" bgcolor="#FFFFFF">
					  <td height="24">
					    <p class="text-menu">&nbsp;</p>
					  </td>
					  <td>&nbsp;</td>
				      <td></td>
					</tr>
                    <tr valign="top" bgcolor="#FFFFFF">
					  <td height="24" >
					    <p class="text-menu">Arancel</p>
					  </td> 
					  <td><input id ="RbArancelBanco" type="radio" name="rbArancel" value="BANCO" checked><b class="Tit-celdas">Banco (Contado).</b> </td>
				      <td><input id ="RbArancelSevipag" type="radio" name="rbArancel" value="SERVIPAG"><b class="Tit-celdas">Servipag (Contado).</b></td>
                      <td width="352"><input id ="RbArancelPagare" type="radio" name="rbArancel" value="PAGARE"> <b class="Tit-celdas">Pagar&eacute; (Cuotas).</b> </td>
					</tr>
                    <tr valign="top" bgcolor="#FFFFFF">
					  <td height="24">
					  </td> 
					  <td></td>
				      <td></td>                      
					</tr>
                    <tr valign="top" bgcolor="#FFFFFF">
					  <td height="24">
					  </td>
					  <td></td>
				      <td></td>                      
					</tr>
                    </form>
                    <tr valign="top" bgcolor="#FFFFFF">
					  <td height="24">
                      </td>
					  <td></td>
                      <td></td> 
				      <td><A onMouseOver="MM_swapImage('Image7','','Imagenes/botones/continuar-on.gif',1)" onmouseout=MM_swapImgRestore() href="javascript: submitform()"><IMG src="Imagenes/botones/continuar-of.gif" border="0"  name="Image7"></A></td>                      
					</tr>
			    </table>      
			      <table width="738">
                    <tr>
                    </tr>
                    <tr>
                    </tr>
                  </table>
		      	</td>
                 
			  </tr>
			</table>
			</td>
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
<!--javascript:parent.mainframe.history.go(0)-->