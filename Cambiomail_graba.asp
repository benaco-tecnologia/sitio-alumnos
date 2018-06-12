<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%

if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if
mail = request("mail")

str ="SP_ACTUALIZA_MAIL_PA '"& session("Rutcli") &"','"& mail &"'"
Session("Conn").execute(str)

Audita 580,"Ingresa a Portal de Alumnos Login"
%>

<%if Session("SituacionActual") = "N" then %>
	<script>
		alert('Se ha actualizado correctamente el correo electr\u00f3nico.');
        window.top.location.href = 'menu_tomaderamos.asp';
    </script>
<%else%>
	<script>
		alert('Se ha actualizado correctamente el correo electru\u00f3nico.');
        window.top.location.href = 'SituActual.asp';
    </script>
<%end if%>
