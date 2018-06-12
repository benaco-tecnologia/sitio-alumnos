<%@LANGUAGE="JAVASCRIPT" CODEPAGE="65001"%>

<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->


<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
-->
</style></head>

<%@Language="VBSCRIPT"%>
<%
  Option Explicit
  On Error Resume Next
  Response.Clear
  Dim objError
  Set objError = Server.GetLastError()
%>

<body>
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
                <table width="722" height="162" border="0" cellpadding="0" cellspacing="0">
                  <tr>
                    <td width="250" background="Imagenes/atencion.jpg"></td>
                    <td width="472" align="left" valign="bottom">
                    <table width="473" height="136" border="0" cellpadding="0" cellspacing="1">
                      <tr>
                         <td width="122"><img src="Imagenes/descripcion.jpg" width="122" height="27"></td>
                      </tr>
                      <% If Len(CStr(objError.ASPCode)) > 0 Then %>
                      <tr>
                         <td width="122"><img src="Imagenes/iss_error_number.jpg" width="122" height="17"></td>
                         <td width="348" bgcolor="#F3F3F3"><%=objError.ASPCode%></td>
                      </tr>
                      <% End If %>
                      <% If Len(CStr(objError.Number)) > 0 Then %>
                      <tr>
                            <td bgcolor="#E6E6E6"><img src="Imagenes/com_error_number.jpg" width="122" height="17"></td>
                            <td width="348" bgcolor="#F3F3F3"><%=objError.Number%><%=" (0x" & Hex(objError.Number) & ")"%></td>
                      </tr>
                      <% End If %>
                      <% If Len(CStr(objError.Source)) > 0 Then %>
                      <tr>
                        <td bgcolor="#F3F3F3"><img src="Imagenes/error_source.jpg" width="121" height="17"></td>
                        <td bgcolor="#F3F3F3"><%=objError.Source%></td>
                      </tr>
                      <% End If %>
                      <% If Len(CStr(objError.File)) > 0 Then %>
                      <tr>
                        <td bgcolor="#E6E6E6"><img src="Imagenes/file_name.jpg" width="122" height="17"></td>
                        <td bgcolor="#E6E6E6"><%=objError.File%></td>
                      </tr>
                      <% End If %>
                      <% If Len(CStr(objError.Line)) > 0 Then %>
                      <tr>
                        <td bgcolor="#E6E6E6"><img src="Imagenes/line_number.jpg" width="122" height="17"></td>
                        <td bgcolor="#F3F3F3"><%=objError.Line%></td>
                       </tr>
                      <% End If %>
                      <% If Len(CStr(objError.Description)) > 0 Then %>
                      <tr>
                        <td bgcolor="#E6E6E6"><img src="Imagenes/brief_descripcion.jpg" width="122" height="17"></td>
                        <td bgcolor="#E6E6E6"><%=objError.Description%></td>
                      </tr>
                      <% End If %>
                      <% If Len(CStr(objError.ASPDescription)) > 0 Then %>
                      <tr>
                        <td bgcolor="#E6E6E6"><img src="Imagenes/full_descripcion.jpg" width="122" height="17"></td>
                        <td bgcolor="#F3F3F3"><%=objError.ASPDescription%></td>
                      </tr>
                      <% End If %>
                    </table></td>
                  </tr>
                </table>
                </td>
			  </table>
			</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>                
</body>
</html>

