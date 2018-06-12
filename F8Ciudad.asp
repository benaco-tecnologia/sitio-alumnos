<%
    Response.Expires = -1
    Session.Timeout = 999
    Response.Buffer = false
    Server.ScriptTimeout = 99999
%>
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%
dim str
dim rst
dim Ciudad 
Ciudad= trim(request.form("txtcodigociudad"))
if Ciudad <> "" then
	str="select * from mt_ciudad where codciudad  like '%" & Ciudad & "%'"  
else
	str="select * from mt_ciudad  order by codciudad asc "
end if
set rst=Session("Conn").execute(str)

%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script>
function filtra(){
	TheForm.action="F8Ciudad.asp"
	TheForm.submit();
}
function consulta(Ciudad)
{ window.opener.document.TheForm.txtciudadact.value=Ciudad;
   window.close();
}
</script>
<title>Consulta Comunas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/ed.css" rel="stylesheet" type="text/css">
</head>

<body background="ima/iso-uhc.gif">
<form name="TheForm" method="post" action="">
  <table width="530" border="0">
    <tr> 
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr class="titulo"> 
      <td colspan="3">Consulta de <span class="titulo">Ciudad</span></td>
    </tr>
    <tr class="titulo">
      <td colspan="3"><div align="right">
          <input name="Cerrar" type="submit" id="Cerrar" value="Cerrar" onClick="javascript:window.close();">
        </div></td>
    </tr>
  </table>
  <table width="530" border="0">
    <tr>
      <td width="222" class="tex">Busqueda por Codigo de Ciudad</td>
      <td width="298"><input name="txtcodigociudad" type="text" id="txtcodigociudad" value="<%=Ciudad%>" onblur="javascript:filtra();"> </td>
    </tr>
    <tr>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
  </table>
  <p>&nbsp;</p>
  <table width="530" height="27" border="0">
    <tr> 
      <td width="172" class="titulo2">Codigo Ciudad</td>
      <td width="348" class="titulo2">Nombre Ciudad</td>
    </tr>
  </table>
  <table width="530"   border="0">
    <%while not rst.eof %>
    <tr class="sub-titulo"> 
      <td width="172"><a href="javascript:consulta('<%=trim(rst("codciudad"))%>')"> <%=rst("codciudad")%>&nbsp; </a></td>
      <td width="348"><%=rst("nomciudad")%></td>
	  <%rst.movenext%>
	 </tr>
    <%wend%> 
  </table>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
</form>
</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
