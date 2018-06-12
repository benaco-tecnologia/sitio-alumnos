<% 
    Response.Expires = -1
%>   
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
<script languaje="javascript">
  //alert("Prueba");
  <% if not InscritoNormal() then  %>
      if (confirm("No ha inscrito aún los ramos ..."))
        { window.location = "alumn-udd.htm"; }
      else {window.close() }
  <% else  %>
      window.close();
  <% end if %>
</script>
<P>&nbsp;</P>

</BODY>
</HTML>
<!--#INCLUDE file="include/desconexion.inc" -->