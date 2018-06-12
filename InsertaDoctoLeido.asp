<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#include file="vars.inc.asp" --> 
<%
codramo = request("CR")
codprof = request("CP")
codsecc = request("CS")

Str = "SP_INSERTA_DOCTO_LEIDO '" & session("codcli") & "','" & codramo & "','" & codsecc & "','" & codprof & "'" 
Session("Conn").Execute(Str) 

Respuesta="OK"

Respuesta="'Respuesta':'" & Respuesta & "' "	
response.Write("({" & Respuesta & "})")
response.End()


%>

