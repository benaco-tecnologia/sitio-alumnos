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
if SESSION("PER_ID") = "-1" then
	response.Redirect("menu_tomaderamos.asp")
end if 

if EstaHabilitadaNW (473)="S" then 
	if GetPermisoNW(473) ="N" then
		response.Redirect("MensajesBloqueos.asp")	
	else
		if TomadeRamoED(session("Codcarr"),session("Codcli"),session("ano_ed"),session("periodo_ed"))=0 then
			session("MensajeBloqueosVarios") ="Para acceder a esta opción debe responder la Encuesta Docente."
			response.Redirect("MensajeBloqueo.asp")
		end if 
		
		strParame="SELECT dbo.Fn_ValorParame('BLOQUEAPAENCUESTAS')Parame"
		set rstParame= Session("Conn").Execute(strParame)		
		if not rstParame.eof then
				BLOQUEAPAENCUESTAS=rstParame("Parame")
			else
				BLOQUEAPAENCUESTAS="" 
		end if
		
		if BLOQUEAPAENCUESTAS="SI" then
			'valida si contesta o no la encuesta
			if TomadeRamo(session("Codcarr"),session("Codcli"),session("ano_ed"),session("periodo_ed"))=0 then
				session("MensajeBloqueosVarios") ="Para acceder a esta opción debe responder sus Encuestas Pendientes."
				response.Redirect("MensajeBloqueo.asp")
			end if 
		end if
	end if
else
	response.Redirect("MensajeBloqueoHabilita.asp")
end if

'MEJORA BLOQUEA ENCUESTASEK
strPar="SELECT dbo.Fn_ValorParame('VALIDAENCUESTASEK')Parame" 
set rstPar= Session("Conn").Execute(strPar)		
if not rstPar.eof then
		VALIDAENCUESTASEK=rstPar("Parame")
	else
		VALIDAENCUESTASEK="" 
end if

if VALIDAENCUESTASEK = "SI" then
	strAl="SP_VALIDAENCUESTA_PA_SEK '"& session("codcli") &"' "
	set rstAl= Session("Conn").Execute(strAl)		
	if not rstAl.eof then
		BloqueaTRRep=rstAl(0) 
	end if
	
	if BloqueaTRRep="NO" then  
		session("MensajeBloqueosVarios") ="Para proceder a ver sus calificaciones de sus asignaturas debe realizar todas las encuestas dispuestas en su portal. </br> Si presentas algún inconveniente en la realización de la encuesta, por favor comunícate con el departamento de Informática de la Universidad, o envíanos un mail a <a href='mailto:soporte.uisek@usek.cl'>soporte.uisek@usek.cl</a>"
		
		response.Redirect("MensajeBloqueo.asp")
	end if
end if 
'FIN MEJORA
%>
<html>
<head>
<title><%=Session("NombrePestana")%></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset cols="198,*" frameborder="no" border="0" framespacing="0" rows="*"> 
  <frame name="leftFrame" scrolling="NO" src="frame-izq.asp" frameborder="NO">
  <frame name="mainFrame" src="solicitud.htm" frameborder="NO" marginwidth="0" marginheight="0" scrolling="yes" noresize>
</frameset>
<noframes> 
<body bgcolor="#FFFFFF" text="#000000">
</body>
</noframes> 
</html>
