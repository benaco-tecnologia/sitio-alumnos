<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%On Error Resume Next%>
<%
	Response.Clear()
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
 	Server.ScriptTimeout = 360
	Response.Buffer = true

	statRamos = 0
	statAlumnos = 0
	statProfesores = 0
	
	Const servidor = "192.168.1.194"
	Const nombreUsuario = "matricula"
	Const password = "dtb01s"
	Const baseDatos = "matricula"									
	
	Const servidorMy = "192.168.4.10"
	Const nombreUsuarioMy = "matricula"
	Const passwordMy = "dtb01s"
	Const baseDatosMy = "ecasvirtual"
	
	Const strDsn = "ecasvirtual"
	
	strConnMySQL = "DSN=" & strDsn &  ";UID="&nombreUsuarioMy&";PWD="&passwordMy&"" 
	Set Session("ConnMySQL") = Server.CreateObject("ADODB.Connection")
	Session("ConnMySQL").Open(strConnMySQL)								
	
	Set Session("Conn") = Server.CreateObject("ADODB.Connection")
	Session("Conn").Open "DRIVER={SQL Server}; SERVER="& servidor &"; DATABASE="& baseDatos &"; UID="& nombreUsuario &"; PWD="& password &""
	
	Function getContextID(CODRAMO)
		getContextID = ""
		strQuery = "select mdl_context.id as id " & _
					"from mdl_course, mdl_context " & _ 
					"where mdl_context.instanceid = mdl_course.id and (mdl_course.shortname = '"&CODRAMO&"' " & _
					"or mdl_course.equivalente like '"&CODRAMO&"' or mdl_course.equivalente like '%"&CODRAMO&"' " & _
					"or mdl_course.equivalente like '"&CODRAMO&"%') LIMIT 1;"
		Set rstmp = Session("ConnMySQL").Execute(strQuery)
		
		If (Not rstmp.EOF) Then
			getContextID = rstmp("id")
		End IF
		Set rstmp = Nothing
	End Function
	
	Function getUserID(RUT)
		getUserID = ""
		strQuery = "select id " & _
					"from mdl_user " & _ 
					"where username = '"&RUT&"';"
		Set rstmp = Session("ConnMySQL").Execute(strQuery)
		
		If (Not rstmp.EOF) Then
			If (CLng(rstmp("id")) > 0) Then
				getUserID = rstmp("id")
			End If
		End IF
		Set rstmp = Nothing
	End Function
	
	strAno = Request.Form("ano")
	strPeriodo = Request.Form("periodo")
	strQuery = "DELETE FROM mdl_role_assignments WHERE INSCRIPCION = 'UMAS';"
	Session("ConnMySQL").Execute(strQuery)
	
	strQueryAsignaciones = "select n.ramoequiv as CODRAMO, a.rut AS RUT, 5 AS ROLEID " & _
						   "from ra_nota n, mt_alumno a " & _
						   "where n.ano = "&strAno&" and n.periodo = "&strPeriodo&" and n.codcli = a.codcli " & _
								 "and a.estacad = 'VIGENTE' and a.matriculado = 'S' " & _
						   "UNION " & _
						   "SELECT CODRAMO, CODPROF AS RUT, 3 AS ROLEID " & _
						   "FROM RA_SECCIO " & _
						   "WHERE ANO = "&strAno&" AND PERIODO = "&strPeriodo&" " & _
						   "order by CODRAMO, ROLEID, RUT;"	
	strCodRamo = ""
	strContextID = ""
		
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open strQueryAsignaciones, Session("Conn"), 1, 1
	Response.Flush
	
	do while not rs.EOF
		If (strCodRamo <> rs("CODRAMO")) Then
			strContextID = getContextID(rs("CODRAMO"))
			strCodRamo = rs("CODRAMO")
			statRamos = statRamos + 1
		End If

		strUserID = getUserID(rs("RUT"))
		
		If (strUserID <> "") Then
			strQuery = "INSERT INTO mdl_role_assignments (roleid, contextid, userid, enrol, modifierid, INSCRIPCION) VALUES " & _
		 			   "('"&rs("ROLEID")&"', '"&strContextID&"','"&strUserID&"', 'manual', 2, 'UMAS');"
			Session("ConnMySQL").Execute(strQuery)
			
			If Err.Number <> 0 then
				'Response.Write(Err.Description & "<br />")
				Error.Clear
			Else
				If (rs("ROLEID") = "5") Then
					statAlumnos = statAlumnos + 1
				Else
					statProfesores = statProfesores + 1
				End If
			End If	
		End If	
		Response.Clear()
		rs.MoveNext	
	loop
	Response.Write(CStr(statRamos)&";"&CStr(statAlumnos)&";"&CStr(statProfesores))
%>