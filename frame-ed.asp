<%
  Response.Expires = -1
%>
<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE file="analytics.asp" -->

<%

'Obtengo numero de encuesta docente vigente		
sqlenc="select coalesce(numero,0)numero from RA_ENCUESTA Where CONVERT(DATE,getDate())>=FecIni And CONVERT(DATE,getDate()) <=FecFin AND TIPO_ENCUESTA=1"
set rstenc= Session("Conn").Execute(sqlenc)	
if rstenc.eof then
	response.Redirect("menu_encuestas.asp")
end if


strParame="SELECT coalesce(dbo.Fn_ValorParame('USAANALYTICS'),'NO')Parame"
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
	USAANALYTICS=rstParame("Parame")
end if

if USAANALYTICS="SI" then
	analitycs() 
end if 
%>

<script languaje="javascript">
function Enviar3(pag,valor,validaencuesta) 
{
	if (validaencuesta=='SI') 
	{	
		if (valor==0) { 
			alert("Para acceder a esta opci\u00f3n debe responder sus Encuestas Pendientes");
			return;
		}
		top.location.href = pag;
	}
	else
	{
		top.location.href = pag;
	}
}
</script>
<%
	if Session("CodCli") = "" then
	  Response.Redirect("saltoinicio.htm")
	end if
	
	if EstaHabilitadaNW (475)="S" then 
		if GetPermisoNW(475) ="N" then
			response.Redirect("MensajesBloqueos.asp")		
		end if
	else
		response.Redirect("MensajeBloqueoHabilita.asp")
	end if 
%>
<html>
<head>
<title><%=Session("NombrePestana")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset cols="198,*" frameborder="no" border="0" framespacing="0" rows="*"> 
  <frame name="leftFrame" scrolling="NO" src="frame-izq.asp" frameborder="NO">
  <frame name="mainFrame" src="ed.htm" frameborder="NO" marginwidth="0" marginheight="0" scrolling="yes" noresize>
</frameset>
<noframes> 
<body bgcolor="#FFFFFF" text="#000000">
</body>
</noframes> 
</html>
