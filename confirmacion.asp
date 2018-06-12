<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="include/mail.inc" -->
<!--#INCLUDE FILE="include/audita.inc" -->
<!--#INCLUDE FILE="include/Cupos.inc" -->
<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%
dim CodCli,CodSede,CodPestud,Codcarr,ano,Periodo, Confirma

CodCli = Session("CodCli")
CodSede = Session("CodSede")
CodPestud = Session("CodPestud")
Codcarr=session("Codcarr")
ano=session("ano")
Periodo=session("Periodo")


if Session("CodCli") = "" then
   Response.Redirect("saltoinicio.htm")
end if
   
'response.Write(Confirma)
'Response.Write("Por favor, disculpe las molestias...Estamos trabajando para usted...Bettersoft...")'lol
'Response.End()


 if ValidaCuposPreinscritos(Codcli,Ano,Periodo,Codpestud,Codcarr,Codsede)=1 and ValidaCuposPreinscritosActividad(Codcli,Ano,Periodo,Codpestud,Codcarr,Codsede)=1 and ValidaInscripcionLPT(Codcli,Ano,Periodo,Codpestud,Codcarr,Codsede)=1  and ValidaInscripcionTeorico(Codcli,Ano,Periodo,Codpestud,Codcarr,Codsede)=1 then
     Confirma=1
	  
	   Codcli = Session("Codcli")
	   Ano = Session("Ano")
	   Periodo = Session("Periodo")
	
	   'Mejora Confirmacion tr
	   StrSql = "SP_CONFIRMA_TR_PORTAL '" & Codcli & "', " & Ano & "," & Periodo & ""
	   Session("Conn").execute StrSql
	
 else
' 	redireccionar a la Pagina anterior...
	Confirma=0
 end if
%>
<body bgcolor="#FFFFFF" text="#000000">

<%
strParame="SELECT coalesce(dbo.Fn_ValorParame('USAMATRICULAWEB_TR'),'NO')Parame"
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
	USAMATRICULAWEB_TR=rstParame("Parame")
end if

if Confirma = 1 and USAMATRICULAWEB_TR ="SI" then
	Session("RutClienteR")=Session("RutClienteR") & "0"
	destinoRedirect = "RedirectNet.asp?PaginaRedirect=matriculaweb.aspx"
	response.Redirect(destinoRedirect)
end if%>

<%if Confirma = 1 then%>
<script>


	window.top.location.href = "resultado.asp"
    //parent.location = "resultado.htm"
    //parent.parent.frames("leftFrame").location = "f-izq.asp";

</script>
<%else%>
<script>
  //alert("Esta asignatura ya no tiene cupo, debes revisar y modificar tu inscripcion rte de Seccion en el Ramo <%= session("RamoSinCupo")%>  ya que la vacante, fue Confirmado por otro alumno ....!!!! ")
  //window.top.mainFrame.location.href="inscrip-asigna.htm"
// parent.location = "frame-inscrip-asigna.htm"
  window.top.location.href='frame-inscrip-asigna.asp';
  //window.top.mainFrame.location.href="asignatura-seccion.asp"
  //parent.parent.frames("leftFrame").location = "f-izq.asp";
</script>
<%end if%>
</body>
</html>
