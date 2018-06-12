

<%
<!--#INCLUDE file="include/funciones.inc" -->

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
	'Dim oCmd
	'Dim ConnectStr

	'Set oCmd = Server.CreateObject("ADODB.Connection")
	'Conn = "Provider=SQLOLEDB; Data Source=(local)" & _
	'		  "; Initial Catalog=uisek" & _
	'		  "; User ID=matricula" & _
	'		  "; Password=dtb01s;"
	'oCmd.open Conn
	'Set Conexion = oCmd

end function

function getEncuestaVigente(tipo_encuesta)
	'Set Sql = Conexion()
	dim strsql
	dim rs
	
	
	StrSql = ""
	StrSql = StrSql & "Select Numero "
	StrSql = StrSql & "From ra_encuesta "
	StrSql = StrSql & "Where CONVERT(DATE,GETDATE()) >= fecIni "
	StrSql = StrSql & "	and CONVERT(DATE,GETDATE()) <= fecFin" & _ 
	    " and tipo_encuesta = " & tipo_encuesta &" " 
	
	if 	tipo_encuesta = 7 then 
		sqlNuev="select coalesce(nuevo,'')nuevo from mt_alumno where codcli ='"& session("codcli") &"'"
		if bcl_ado(sqlNuev,rstNuev) then
			if rstNuev("nuevo") = "S" then
				TipoAlumno="Nuevo"
			else
				TipoAlumno="Antiguo"
			end if				
		else
			TipoAlumno=""
		end if
		
		StrSql = StrSql & " AND (TIPO_ALUMNO ='"& TipoAlumno &"' OR COALESCE(TIPO_ALUMNO,'Todos') = 'Todos')"
	
	end if 

	
	'Set rs = Sql.Execute(f_NormSql(StrSql))
	'Set rs = Session("Conn").Execute(f_NormSql(StrSql))
	if bcl_ado(strsql,rs) then
			getEncuestaVigente = rs("Numero")
		else
			getEncuestaVigente = 0
	end if
    
    'response.Write(getEncuestaVigente)
	'response.End()
	
	'Set rs = Nothing
	'Set Sql = Nothing
	'Set StrSql = Nothing
	
end function

function getEncuestaVigenteAnno()
	'Set Sql = Conexion()
	dim strsql
	dim rs
	
	StrSql = ""
	StrSql = StrSql & "Select Ano "
'	StrSql = StrSql & "From ra_encuesta_ei "
	StrSql = StrSql & "From ra_encuesta "
	StrSql = StrSql & "Where getDate() >= fecIni "
	StrSql = StrSql & "and getDate() <= fecFin"
	
	'response.Write(strsql)
	'response.End()
	
	'Set rs = Sql.Execute(f_NormSql(StrSql))
	'Set rs = Session("Conn").Execute(f_NormSql(StrSql))
	
	if bcl_ado(strsql,rs) then
			getEncuestaVigenteAnno = rs("ano")
		else
			getEncuestaVigenteAnno = 0
	end if

	'Set rs = Nothing
	'Set Sql = Nothing

end function

function getEncuestaVigentePeriodo()
	'Set Sql = Conexion()
	dim Strsql
	dim rs
	
	StrSql = ""
	StrSql = StrSql & "Select Periodo "
'	StrSql = StrSql & "From ra_encuesta_ei "
	StrSql = StrSql & "From ra_encuesta "
	StrSql = StrSql & "Where getDate() >= fecIni "
	StrSql = StrSql & "	and getDate() <= fecFin"
	
	'response.Write(StrSql)
	'Set rs = Sql.Execute(f_NormSql(StrSql))
	'Set rs = Session("Conn").Execute(f_NormSql(StrSql))
	
	'if Not rs.Eof _
	'	then
	if bcl_ado(strsql, rs) then
			getEncuestaVigentePeriodo = rs("Periodo")
		else
			getEncuestaVigentePeriodo = 0
	end if

	'Set rs = Nothing
	'Set Sql = Nothing
end function

function getBloques(NroEncuesta)
	'Set Sql = Conexion()
	dim strsql
	dim rs
	
	StrSql = ""
	'StrSql = StrSql & "Select CodTipoPregunta as Orden, Descripcion "
'	StrSql = StrSql & "From RA_TIPOPREGUNTA_EI "
	'StrSql = StrSql & "From RA_TIPOPREGUNTA "
	'StrSql = StrSql & "Where NroEncuesta='" & Nroencuesta & "' "
	'StrSql = StrSql & "Order By CodTipoPregunta"
	
	
	
	'StrSql = StrSql & "Select distinct a.CodTipoPregunta as Orden, a.Descripcion "
	'StrSql = StrSql & "From RA_TIPOPREGUNTA a,RA_ENCUESTA_DETALLE b "
	'StrSql = StrSql & "Where CodTipoPregunta = Tipo_pregunta and nro_encuesta='" & Nroencuesta & "' "
	'StrSql = StrSql & "Order By a.CodTipoPregunta"
	
	StrSql = StrSql & "SELECT a.CodTipoPregunta AS Orden , "
    StrSql = StrSql & "	a.Descripcion,MAX(NRO_LINEA)AS NRO_LINEA ,"
    StrSql = StrSql & " a.TIPOCONTROLRESPUESTA AS CONTROL_RESP "
	StrSql = StrSql & "FROM    RA_TIPOPREGUNTA a , "
	StrSql = StrSql & "RA_ENCUESTA_DETALLE b "
	StrSql = StrSql & "WHERE   CodTipoPregunta = Tipo_pregunta "
	StrSql = StrSql & "AND nro_encuesta ='" & Nroencuesta & "' "
	StrSql = StrSql & "GROUP BY a.CodTipoPregunta,a.Descripcion, a.TIPOCONTROLRESPUESTA "
	StrSql = StrSql & "ORDER BY NRO_LINEA "
	
	'response.Write(strsql)
	'response.End()
	
	'Set getBloques = Sql.Execute(f_NormSql(StrSql))
	Set getBloques = Session("Conn").Execute(f_NormSql(StrSql))
	'if bcl_ado(strsql,rs) then 
	'	getBloques = rs
	
	'end if 
	'Set Sql = Nothing
	
end function

function getPreguntas(NroEncuesta, tipo_pregunta )
	'Set Sql = Conexion()
	Dim strsql
	dim rs
	
	'StrSql = ""
	'StrSql = StrSql & "Select * "
	'StrSql = StrSql & "From RA_PREGUNTAS "
	'StrSql = StrSql & "Where COD_TIPOPREGUNTA='" & NroBloque & "' "
	'StrSql = StrSql & "Order By Numero"
	
	StrSql = "SELECT	p.Numero , p.Pregunta, p.CodRespuesta as TipoRespuesta, d.Nro_Linea, coalesce(t.TipoControlRespuesta,'') as TipoControlRespuesta " & _
            " FROM	RA_ENCUESTA_DETALLE d, ra_preguntas p, ra_tipopregunta t " & _
            " WHERE	NRO_ENCUESTA = " & NroEncuesta & " AND TIPO_PREGUNTA = " & tipo_pregunta & _
	        " and d.Tipo_pregunta = t.CodTipoPregunta  AND d.NRO_PREGUNTA = p.NUMERO " & _
            " ORDER BY 	NRO_LINEA "

	'response.Write(StrSql)
	'Set getPreguntas = Sql.Execute(f_NormSql(StrSql))
	Set getPreguntas = Session("Conn").Execute(f_NormSql(StrSql))	
	'if bcl_ado(strsql) then 
	'	getPreguntas = rs
	'end if 
	
	'Set Sql = Nothing
end function

function getRespuestas(TipoRespuesta)
	'Set Sql = Conexion()
	dim strsql
	dim rs
	
	StrSql = ""
	StrSql = StrSql & "Select CodRespuesta as Respuesta, Respuesta as Descripcion "
'	StrSql = StrSql & "From RA_RESPUESTAED_EI "
 	StrSql = StrSql & "From RA_RESPUESTA "
	StrSql = StrSql & "Where CodTipoRespuesta='" & TipoRespuesta & "'"
	StrSql = StrSql & "Order By CodRespuesta"

	
	'Set getRespuestas = Sql.Execute(f_NormSql(StrSql))
	Set getRespuestas = Session("Conn").Execute(f_NormSql(StrSql))
	'if bcl_ado(strsql, rs) then
	 '  getRespuestas = rs
	'end if  
	
	'Set Sql = Nothing
end function

function InsertLineaResultado(s)
	'Set Sql = Conexion()
	dim strsql
	dim rs
	
	StrSql = s
	'response.Write(StrSql)
	'response.End()
	'Set InsertLineaResultado = Sql.Execute(f_NormSql(StrSql))
	Set InsertLineaResultado = Session("Conn").Execute(f_NormSql(StrSql))
	
	'if bcl_ado(strsql,rs) then 
	'	InsertLineaResultado = rs
	'end if 
	'Set Sql = Nothing
end function

function getTextoLibre(TipoRespuesta)
	'Set Sql = Conexion()
	dim strsql
	dim rs
	
	StrSql = ""
	StrSql = StrSql & "Select TextoLibre as Texto "
'	StrSql = StrSql & "From RA_RESPUESTAED_EI "
 	StrSql = StrSql & "From RA_TIPORESPUESTA "
	StrSql = StrSql & "Where CODTIPORESPUESTA='" & TipoRespuesta & "'"

	
	'Set getRespuestas = Sql.Execute(f_NormSql(StrSql))
	Set getTextoLibre = Session("Conn").Execute(f_NormSql(StrSql))
	'if bcl_ado(strsql, rs) then
	 '  getRespuestas = rs
	'end if  
	
	'Set Sql = Nothing
end function


function InsertEncuestados(CodCli, CodRamo, CodSecc, Ano, Periodo, CodProf, CodCarr, Obs, Fecha, Evaluado, NroEncuesta)
	'Set Sql = Conexion()
	dim strsql
	dim rs
	
	StrSql = ""
	StrSql = StrSql & "insert into RA_ENCUESTADOS " & _ 
	    " (CodCli, CodRamo, CodSecc, ANO, PERIODO, " & _ 
	    " CodProf, CodCarr, Observacion, Fecha, Evaluado, " &_ 
	    " Encuesta) "
	StrSql = StrSql & "Values('" & CodCli & "'," & iif(CodRamo <> "","'" & CodRamo & "'", "null") & ", " & iif(CodSecc <> 0,  "'" & CodSecc & "'", "null") & ", '" & Ano & "', '" & Periodo & _ 
	     "'," & iif(CodProf <> "", "'" & CodProf & "'", "null") & "," & iif(CodCarr<>"", "'" & CodCarr & "'", "null") & ", LEFT('" & Obs  & "',500), getdate(), 'SI'," & NroEncuesta & " ) "
	
	'response.Write(strsql)
	'response.End()	
	'Set InsertEncuestados = Sql.Execute(f_NormSql(StrSql))
	Set InsertEncuestados = Session("Conn").Execute(f_NormSql(StrSql))
	
	'Set Sql = Nothing
end function

function getFaltaPorResponder(thisCodCli, Nro)
	'Set Sql = Conexion()
	dim strsql 
	dim rs
	
	strSql = ""
	strSql = strSql & "Select Count(CodCli) as hay "
	strSql = strSql & "From ra_encuestados "
	strSql = strSql & "Where CodCli='" & thisCodCli & "' And Encuesta='" & Nro & "' And Evaluado <> 'NO' "
	
	'response.Write(strsql)
	'response.End()
	'Set rs = Sql.Execute(f_NormSql(StrSql))
	'Set rs = Session("Conn").Execute(f_NormSql(StrSql))
	if bcl_ado(strsql, rs) then 
		getFaltaPorResponder = rs("hay")	
	end if 
	'Set Sql = Nothing
end function%>
