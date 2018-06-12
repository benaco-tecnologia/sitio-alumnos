<!--#include file="include/lib.asp" -->
<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if
NroEncuesta= Session("NroEncuesta")
'AnnoEncuesta = getEncuestaVigenteAnno()
AnnoEncuesta = Session("AnnoEncuesta")
'PeriodoEncuesta = getEncuestaVigentePeriodo()
PeriodoEncuesta = Session("PeriodoEncuesta") 

Set rsBloques = getBloques(NroEncuesta)

'NUEVO BORRA ENCUESTA ANTES DE INSERTAR
strBR ="delete from ra_resultados WHERE CODCLI ='" & Session("CodCli") & "' AND ano='"& Session("AnnoEncuesta") &"' AND periodo='"& Session("PeriodoEncuesta") &"' AND encuesta='"& Session("NroEncuesta") &"'"

Session("Conn").Execute(strBR)


while not rsBloques.Eof
	NroBloque = rsBloques("Orden")
	 Set rsPreguntas = getPreguntas(NroEncuesta, rsBloques("Orden"))
			while not rsPreguntas.Eof
			NroPregunta = rsPreguntas("Numero")
			
			
			'if rsPreguntas("TipoControlRespuesta") = "MULTIPLE" then
			if 0 = 1 then
				Cadena = Split(request.form("Check_"& rsBloques("Orden") & "_" & rsPreguntas("Numero")), ",")
				
				for i = 0 to ubound(Cadena)
					tmp = Split(Cadena(i), "_")
					Respuesta = tmp(2)
					
					SqlQry = ""
					SqlQry = "Insert Into RA_RESULTADOS " & _ 
						" (Ano, Periodo, CodCli, Rut, Codramo, " &_ 
						" CodSecc, CodProf, TipoDocente, NroPregunta, Respuesta, " & _ 
						" Encuesta, Fecha , TextoLibre) "
					SqlQry = SqlQry & "Values('" & AnnoEncuesta & "', "
					SqlQry = SqlQry & "	    '" & PeriodoEncuesta & "', "
					SqlQry = SqlQry & "		'" & Session("CodCli") & "', "
					SqlQry = SqlQry & "		'" & Session("Rut") & "', "
					SqlQry = SqlQry & "		Null, "
					
					SqlQry = SqlQry & "		Null, "
					SqlQry = SqlQry & "		Null, "
					SqlQry = SqlQry & "		Null, "			
					SqlQry = SqlQry & "		'" & NroPregunta & "', "
					'SqlQry = SqlQry & "		'" & NroBloque & "', "
					if Respuesta = "" then 
					SqlQry = SqlQry & "		Null, "
					else
					SqlQry = SqlQry & "		'" & Respuesta & "', "
					end if
					SqlQry = SqlQry & "		'" & NroEncuesta & "', getdate(),"
					if TextoLibre = "" then 
					SqlQry = SqlQry & "		Null) "
					else
					SqlQry = SqlQry & "		'" & TextoLibre & "') "
					end if
					
					'response.Write(SqlQry)
					Set Trash = InsertLineaResultado(SqlQry) 
				next

			else
						
				if request.form("Select_" & rsBloques("Orden") & "_" & rsPreguntas("Numero")) = "0" then 
					Respuesta = ""
				else
					tmp = Split(request.form("Select_" & rsBloques("Orden") & "_" & rsPreguntas("Numero")), "_")
					Respuesta = tmp(2)								
				end if
				
				TextoLibre = replace(request.Form("txtTextoLibre_" & rsBloques("Orden") & "_" & rsPreguntas("Numero")), "'", "")
				
				SqlQry = ""
				SqlQry = "Insert Into RA_RESULTADOS " & _ 
					" (Ano, Periodo, CodCli, Rut, Codramo, " &_ 
					" CodSecc, CodProf, TipoDocente, NroPregunta, Respuesta, " & _ 
					" Encuesta, Fecha , TextoLibre) "
				SqlQry = SqlQry & "Values('" & AnnoEncuesta & "', "
				SqlQry = SqlQry & "	    '" & PeriodoEncuesta & "', "
				SqlQry = SqlQry & "		'" & Session("CodCli") & "', "
				SqlQry = SqlQry & "		'" & Session("Rut") & "', "
				SqlQry = SqlQry & "		Null, "
				
				SqlQry = SqlQry & "		Null, "
				SqlQry = SqlQry & "		Null, "
				SqlQry = SqlQry & "		Null, "			
				SqlQry = SqlQry & "		'" & NroPregunta & "', "
				'SqlQry = SqlQry & "		'" & NroBloque & "', "
				if Respuesta = "" then 
				SqlQry = SqlQry & "		Null, "
				else
				SqlQry = SqlQry & "		'" & Respuesta & "', "
				end if
				SqlQry = SqlQry & "		'" & NroEncuesta & "', getdate(),"
				if TextoLibre = "" then 
				SqlQry = SqlQry & "		Null) "
				else
				SqlQry = SqlQry & "		'" & TextoLibre & "') "
				end if
				
				'response.Write(SqlQry)
				Set Trash = InsertLineaResultado(SqlQry) 
			end if 
			
			rsPreguntas.MoveNext	
		wend
	rsBloques.MoveNext
wend
'response.End()
'response.Write("HOLA")
'response.End()
txtObservaciones = replace(request.Form("txtObservaciones"), "'", "")
'set trash = InsertEncuestados(AnnoEncuesta, PeriodoEncuesta, Session("CodCli"), Session("CodCli"), "", 0, "", "", Session("CodCarr"), NroEncuesta, txtObservaciones)

'NUEVO BORRA ENCUESTADOS ANTES DE INSERTAR
'strBE ="delete from ra_encuestados WHERE CODCLI ='" & Session("CodCli") & "' AND ano='"& Session("AnnoEncuesta") &"' AND periodo='"& Session("PeriodoEncuesta") &"' AND encuesta='"& Session("NroEncuesta") &"'"

'Session("Conn").Execute(strBE) 

set trash2 = InsertEncuestados(Session("CodCli"), "", 0, AnnoEncuesta, PeriodoEncuesta, "", Session("CodCarr"), txtObservaciones, "", "S", NroEncuesta)


'CodCli, CodRamo, CodSecc, Ano, Periodo, CodProf, CodCarr, Obs, Fecha, Evaluado, NroEncuesta

'response.Redirect("adm-acad.asp")
response.Redirect("menu_encuestas.asp")

response.End()

%>