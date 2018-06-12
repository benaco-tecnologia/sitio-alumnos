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
				<td valign="top"><p style="font-size:25px" class="text-menu">Comprobante de Matr&iacute;cula y Arancel.</p>
				  <table width="800" border="0" cellpadding="0" cellspacing="0">
                  <tr valign="top" bgcolor="#FFFFFF">
                    </tr>                    
			    </table>   
                <table border="1">
                  <tr>
                    <th scope="col">Producto &nbsp;&nbsp;</th>
                    <th scope="col">Tipo de Pago</th>
                    <th scope="col">Cup&oacute;n</th>
                    <th scope="col">Total</th>
                  </tr>
                 
                  <tr>
                    <td>Matr&iacute;cula</td>
                    <td><%=session("Matricula")%></td>
                    <td> 
					<%if session("MatriculaBanco")<>"" then%>
                    	<a  target="_blank"href="<%=session("MatriculaBanco")%>">Descargar</a>
                    <%else%>
                    	N/A
					<%end if %>
                    </td>
                    <td>
                    <%if session("Matricula")="SERVIPAG" then%>
                    	Monto
                    <%else%>
                    	N/A
					<%end if %>
                    </td>
                  </tr>
                 
                  <tr>    
                    <td>Arancel</td>
                    <td><%=session("Arancel")%></td>
                    <td>
                    <%if session("ArancelBanco")<>"" then%>
                    	<a  target="_blank"href="<%=session("ArancelBanco")%>">Descargar</a>
                    <%elseif session("ArancelPagare")<>"" then%>
                    	<a  target="_blank"href="<%=session("ArancelPagare")%>">Descargar</a>
                    <%else%>
                    	N/A
					<%end if %>
                    </td>
                    <td>
                     <%if session("Arancel")="SERVIPAG" then%>
                    	Monto
                    <%else%>
                    	N/A
					<%end if %>
                    </td>
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