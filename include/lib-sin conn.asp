<%
Function f_NormSql(sQry)
Dim sNormQry
Dim i
   sQry = UCase(sQry)
   If (InStr(1, sQry, "SELECT") > 0) And (InStr(1, sQry, "FROM") > 0) Then
     posini = InStr(1, sQry, "FROM")
     posfin = InStr(posini + 4, sQry, "WHERE") - 1
     If posfin <= 0 Then posfin = InStr(posini + 4, sQry, "GROUP BY") - 1
     If posfin <= 0 Then posfin = InStr(posini + 4, sQry, "ORDER BY") - 1
     If posfin <= 0 Then posfin = Len(sQry)
     sNormQry = ""
     For i = posini + 4 To posfin
        If Mid(sQry, i, 1) = "," Then
          sNormQry = sNormQry & " WITH (NOLOCK),"
        Else
          sNormQry = sNormQry & Mid(sQry, i, 1)
        End If
     Next
     sNormQry = sNormQry & " WITH (NOLOCK) "
      f_NormSql = Mid(sQry, 1, posini + 4) + _
                  sNormQry + _
                  Mid(sQry, posfin + 1)
   Else
      f_NormSql = sQry
   End If
End Function

Function IIf(condition,value1,value2)
	If condition _
		Then 
			IIf = value1 
		Else 
			IIf = value2
	end if
End Function

function Conexion()
	Dim oCmd
	'Set oCmd = Server.CreateObject("ADODB.Connection")
	'Conn = "Provider=SQLOLEDB; Data Source=(local)" & _
	'		  "; Initial Catalog=uisek" & _
	'		  "; User ID=matricula" & _
	'		  "; Password=dtb01s;"
	'oCmd.open Conn
	Set Conexion = Session("Conn")
end function

function getEncuestaVigente()
	Set Sql = Conexion()
	
	StrSql = ""
	StrSql = StrSql & "Select Nroencuesta "
	StrSql = StrSql & "From ra_encuesta_ei "
	StrSql = StrSql & "Where getDate() >= fecIni "
	StrSql = StrSql & "	and getDate() <= fecFin"

	'response.Write(strsql)
	'response.End()
	
	Set rs = Sql.Execute(f_NormSql(StrSql))
	
	if not rs.Eof _
		then
			getEncuestaVigente = rs("Nroencuesta")
		else
			getEncuestaVigente = 1
	end if
    
    'response.Write(getEncuestaVigente)
	'response.End()
	
	Set rs = Nothing
	Set Sql = Nothing
end function

function getEncuestaVigenteAnno()
	Set Sql = Conexion()
	
	StrSql = ""
	StrSql = StrSql & "Select Ano "
	StrSql = StrSql & "From ra_encuesta_ei "
	StrSql = StrSql & "Where getDate() >= fecIni "
	StrSql = StrSql & "	and getDate() <= fecFin"
	
	Set rs = Sql.Execute(f_NormSql(StrSql))
	
	if Not rs.Eof _
		then
			getEncuestaVigenteAnno = rs("Ano")
		else
			getEncuestaVigenteAnno = 0
	end if

	Set rs = Nothing
	Set Sql = Nothing
end function

function getEncuestaVigentePeriodo()
	Set Sql = Conexion()
	
	StrSql = ""
	StrSql = StrSql & "Select Periodo "
	StrSql = StrSql & "From ra_encuesta_ei "
	StrSql = StrSql & "Where getDate() >= fecIni "
	StrSql = StrSql & "	and getDate() <= fecFin"
	
	Set rs = Sql.Execute(f_NormSql(StrSql))
	
	if Not rs.Eof _
		then
			getEncuestaVigentePeriodo = rs("Periodo")
		else
			getEncuestaVigentePeriodo = 0
	end if

	Set rs = Nothing
	Set Sql = Nothing
end function

function getBloques(NroEncuesta)
	Set Sql = Conexion()
	
	StrSql = ""
	StrSql = StrSql & "Select CodTipoPregunta as Orden, Descripcion "
	StrSql = StrSql & "From RA_TIPOPREGUNTA_EI "
	StrSql = StrSql & "Where NroEncuesta='" & Nroencuesta & "' "
	StrSql = StrSql & "Order By CodTipoPregunta"
	
	'response.Write(strsql)
	'response.End()
	
	Set getBloques = Sql.Execute(f_NormSql(StrSql))
	Set Sql = Nothing
	
end function

function getPreguntas(NroEntrevista, NroBloque)
	Set Sql = Conexion()
	
	StrSql = ""
	StrSql = StrSql & "Select * "
	StrSql = StrSql & "From RA_PREGUNTAS_EI "
	StrSql = StrSql & "Where COD_TIPOPREGUNTA='" & NroBloque & "' "
	StrSql = StrSql & "Order By Numero"
	
	Set getPreguntas = Sql.Execute(f_NormSql(StrSql))
	
	Set Sql = Nothing
end function

function getRespuestas(TipoRespuesta)
	Set Sql = Conexion()
	
	StrSql = ""
	StrSql = StrSql & "Select * "
	StrSql = StrSql & "From RA_RESPUESTAED_EI "
	StrSql = StrSql & "Where TipoRespuesta='" & TipoRespuesta & "'"
	StrSql = StrSql & "Order By Respuesta"
	
	Set getRespuestas = Sql.Execute(f_NormSql(StrSql))
	
	Set Sql = Nothing
end function

function InsertLineaResultado(s)
	Set Sql = Conexion()
	
	StrSql = s
	response.Write(StrSql)
	Set InsertLineaResultado = Sql.Execute(f_NormSql(StrSql))
	
	Set Sql = Nothing
end function

function InsertEncuestados(CodCli, CodCarr, AnnoEncuesta, PeriodoEncuesta, NroEntrevista, txtObs)
	Set Sql = Conexion()
	
	StrSql = ""
	StrSql = StrSql & "insert into RA_ENCUESTADOS_EI(CodCli, CodCarr, ANO, PERIODO, NROENCUESTA, Fecha, Observaciones) "
	StrSql = StrSql & "Values('" & CodCli & "', '" & CodCarr & "', '" & AnnoEncuesta & "', '" & PeriodoEncuesta & "', '" & NroEntrevista & "', getDate(), '" & txtObs & "')"
		
	Set InsertEncuestados = Sql.Execute(f_NormSql(StrSql))
	
	Set Sql = Nothing
end function

function getFaltaPorResponder(thisCodCli, Nro)
	Set Sql = Conexion()
	
	StrSql = ""
	StrSql = StrSql & "Select Count(CodCli) as hay "
	StrSql = StrSql & "From ra_encuestados_ei "
	StrSql = StrSql & "Where CodCli='" & thisCodCli & "' And NroEncuesta='" & Nro & "'"
	
	Set rs = Sql.Execute(f_NormSql(StrSql))
	getFaltaPorResponder = rs("hay")	
	
	Set Sql = Nothing
end function%>
