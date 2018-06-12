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

<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title><%=Session("NombrePestana")%></title>
<body >
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
			<table width="100%" border="0" cellspacing="0" cellpadding="20" height="277" bgcolor="#FFFFFF">
			  <tr> 
				<td height="78" valign="top"><p><img src="imagenes/titulos/T-ayuda.gif" width="400" height="38"></p>
		        </td>
			  </tr>
               <%
					strLink="sp_traeFwlinkPortalAlumno 'menu_ayuda.asp'"
					Set rsLink = Session("Conn").execute(strLink)
							
					if not rsLink.eof then       
						while not rsLink.eof
						%>
						
						 <tr valign="top" bgcolor="#FFFFFF">
						  <td height="100"><a href="<%="RedirectNet.asp?PaginaRedirect="&rsLink("id")%>" target="<%=rsLink("propiedad")%>">
							<p class="text-menu"><%=rsLink("nombre_link")%></p>
							</a>
							  <p align="justify" class="text-normal-celdas"><span class="tex-normales"><%=rsLink("descripcion")%></span></p></td>
							  
						  <td>&nbsp;</td>
						  <%if rsLink("imagen") <> "" then
								RutaFoto = Server.MapPath("Imagenes") & rsLink("imagen")
								if ExisteArchivo(RutaFoto) then%>
									<td><a href="<%="RedirectNet.asp?PaginaRedirect="&rsLink("id")%>" target="<%=rsLink("propiedad")%>"><img src="<%="Imagenes\"&rsLink("imagen")%>" height="100" width="150"></a></td>
								<%else%>
									<td>&nbsp;</td>    
								<%end if%>	
						  <%else%>
							<td>&nbsp;</td>    			  
						  <%end if%>
						</tr>
							
						<%rsLink.movenext		
						wend
					end if %>
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