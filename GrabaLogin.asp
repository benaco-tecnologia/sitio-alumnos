<!--#include file="lib.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<%


Session("AnoClases") = request.Form("ano")
Periodo = request.Form("Periodo")
CodigoSeccion = request.Form("CodigoSeccion")
CodigoRamo = request.Form("CodigoRamo")
CodigoCarrera = request.Form("CodigoCarrera")
DiaAsist = request.Form("Dia")
ModulosClases = request.Form("elModulo")
Entrada = request.Form("Entrada")
Salida = request.Form("Salida")

'response.Write("xD")
'response.End()

if Entrada ="SI" then
	str="sp_InsertaLoginProfe '"& CodigoRamo &"',"& Session("AnoClases") &","& Periodo &",'"& session("codprof") &"' ,"& CodigoSeccion &",'"& DiaAsist &"','"& ModulosClases &"','ENTRADA'"
	'response.Write(str)
	'response.End()
	Session("Conn").execute(str)
	response.Redirect("Selsede.asp")
end if

if Salida ="SI" then 
	if pasoAsistenciaAlumnos = "SI" then
		if estaentiempodecerrar = "SI" then 
			str="sp_InsertaLoginProfe '"& CodigoRamo &"',"& Session("AnoClases") &","& Periodo &",'"& session("codprof") &"' ,"& CodigoSeccion &",'"& DiaAsist &"','"& ModulosClases &"','SALIDA'"
			Session("Conn").execute(str)
			
			%><script languaje="javascript"> 
				alert("Estimado Docente, se ha registrado correctamente su hora de salida.");
				window.top.location.href = "selsede.asp"; 
			</script><% 
		else
			strM ="SELECT CONVERT(INT,COALESCE(MINUTOSDECIERRE,0))MINUTOSDECIERRE  FROM mt_parame"
			if BCL_ADO(strM,rstM) then   
				tiempo = rstM("MINUTOSDECIERRE")
			else
				tiempo = "0"
			end if
			
			%><script languaje="javascript"> 
				alert("Debe esperar al menos <%=tiempo%> minutos para proceder a cerrar la clase.");
				window.top.location.href = "selsede.asp"; 
			</script><% 
		end if
	else
		%><script languaje="javascript"> 
			alert("Debe pasar asistencia a los alumnos de este m\u00f3dulo para proceder a cerrar la clase.");
			window.top.location.href = "selsede.asp"; 
		</script><%
	end if 
end if 

function estaentiempodecerrar()	 
	
	str="sp_ValidaTiempoLoginProfe '"& CodigoRamo &"',"& Session("AnoClases") &","& Periodo &",'"& session("codprof") &"' ,"& CodigoSeccion &",'"& DiaAsist &"','"& ModulosClases &"'"
	if BCL_ADO(str,rst) then
		if ValNulo(rst("tiempo"), STR_)="SI" then
			estaentiempodecerrar="SI"
		else
			estaentiempodecerrar="NO"
		end if 
	else 
		estaentiempodecerrar="NO"
	end if
	
end function

function pasoAsistenciaAlumnos()	
	
	'str="sp_pasoAsistenciaAlumnos '"& CodigoRamo &"',"& Session("AnoClases") &","& Periodo &",'"& session("codprof") &"' ,"& CodigoSeccion &",'"& DiaAsist &"','"& ModulosClases &"'" 
	
	'if BCL_ADO(str,rst) then
	'	if ValNulo(rst("Asistencia"), STR_)="SI" then
	'		pasoAsistenciaAlumnos="SI"
	'	else
	'		pasoAsistenciaAlumnos="NO"
	'	end if 
	'else 
	'	pasoAsistenciaAlumnos="NO"
	'end if
	pasoAsistenciaAlumnos="SI"
end function

'response.Redirect("Selramo.asp?Fecha="& Session("FechaActiva") &"&Activa="& Session("Activa") &"")
'response.Redirect("Selsede.asp")
%>
