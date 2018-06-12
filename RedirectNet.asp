<!--#INCLUDE FILE="include/funciones.inc" -->

<%
	RedirectDestino = request("PaginaRedirect")
		
	if RedirectDestino = "Certificados.aspx" OR RedirectDestino = "DocumentosEmitidos.aspx" OR RedirectDestino = "listadocertificado.aspx" OR RedirectDestino = "matriculaweb.aspx"then
		pagina = GetPaginaCertifOnline()
		enviaDestino = 1
	else
		if isnumeric(request("PaginaRedirect")) then
			pagina = GetPaginaLink(request("PaginaRedirect"))
			esArchivo = GetEsArchivoPaginaLink(request("PaginaRedirect"))
		else
			pagina =""
			esArchivo =""
		end if
		enviaDestino = 0
	end if   
 
%>
<% 
if esArchivo ="SI" then
	response.Redirect(pagina)
end if 


if pagina<> "" then %> 
    <html>
    <head>
    <title> Redirect </title>
    </head> 
    <body onLoad="document.formulario.submit()">  
    
    <form action="<%=pagina%>" method="post" name="formulario" target="_top"> 
    <input type="hidden" name="user" value="<%=encripta_portal(session("RutClienteR"))%>">
    <input type="hidden" name="pass" value="<%=encripta_portal(Session("MiClave"))%>">
    <input type="hidden" name="codcarr" value="<%=encripta_portal(Session("codcarr"))%>">
    <%if enviaDestino = 1then%>
    	<input type="hidden" name="destino" value="<%=RedirectDestino%>">
    <%end if%>
    </form>
    </body>
    </html>
<%end if %>