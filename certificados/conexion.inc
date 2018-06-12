<%
	Const servidor = "192.168.1.194"
	Const nombreUsuario = "sa"
	Const password = "reflex"
	Const baseDatos = "umasnet"
	
	Dim StrSql, rs
	
	'"Provider=SQLOLEDB; " _  & "Data Source=source_name; " _  & "Database=db_name; " _  & "UID=userid; " _  & "PWD=password;"
	Dim aConnectionString : aConnectionString = "DRIVER={SQL Server}; "_
  											  & "SERVER="& servidor &"; "_
											  & "DATABASE="& baseDatos &"; "_
											  & "UID="& nombreUsuario &"; "_
											  & "PWD="& password &""
									
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open aConnectionString
	
	Function matricula_actual(rut)
		ano = year(now())
		ano_anterior = ano - 1
		
		Set rstmp = Conn.Execute("select rut from MT_ALUMNO where (ano_mat = "&ano&" or ano_mat = "&ano_anterior&") and matriculado = 'S' and rut = '" & rut & "';")
		If not rstmp.eof Then
			matricula_actual = TRUE
		Else
			matricula_actual = FALSE
		End If
		rstmp.Close()
		Set rstmp = Nothing	
	End Function
	
	Function cuotas(codcli)
		Set rstmp = Conn.Execute("select count(*) as numero, min(fecven) as cuotas from (select codcli, fecven from MT_CTADOC where (DATEDIFF(day, FECVEN, GETDATE()) > 70) AND (SALDO > 0) and (datepart(yyyy, fecven) = 2011 or datepart(yyyy, fecven) = 2012) and feccancel is null) a group by a.codcli having a.codcli = '"& codcli &"';")
		if (Not rstmp.EOF) Then
			cuotas = "title=""Vencimiento: " & rstmp(1) & """>" & rstmp(0)
		Else
			cuotas = ">0"
		End If
		rstmp.Close()
		Set rstmp = Nothing
	End Function
	
	Function excepcionCertificado(codcli)
		Set rstmp = Conn.Execute("SELECT rut FROM certificados_excepciones WHERE rut = '" & codcli & "' AND DATEDIFF(day, GETDATE(), fecha) < 8 AND DATEDIFF(day, GETDATE(), fecha) > 0;")
		If (Not rstmp.EOF) Then
			excepcionCertificado = TRUE
		Else
			excepcionCertificado = FALSE
		End If
	End Function
	
	Function deuda_actual(rut)
		Set rstmp = Conn.Execute("select CODCLI from MT_CTADOC where (DATEDIFF(day, FECVEN, GETDATE()) > 70) AND (SALDO > 0) and (datepart(yyyy, fecven) = 2014 or datepart(yyyy, fecven) = 2015) and feccancel is null AND (CODCLI =  '" & rut & "');")
		If not rstmp.eof Then
			deuda_actual = TRUE
		Else
			deuda_actual = FALSE
		End If
		rstmp.Close()
		Set rstmp = Nothing	
	End Function
	
	Function correo(codcli)
		Set rstmp = Conn.Execute("select mail from mt_client where codcli = '" + codcli + "';")
		correo = rstmp("mail")
		rstmp.Close()
		Set rstmp = Nothing	
	End Function
	
	Function estacad(codcli)
		Set rstmp = Conn.Execute("select estacad from mt_alumno where rut = '" + codcli + "' and (estacad = 'VIGENTE' OR estacad = 'EGRESADO' OR estacad='TITULADO');")
		'If trim(rstmp("estacad")) = "VIGENTE" Then
		If (not rstmp.eof) Then
			estacad = TRUE
		Else
			estacad = FALSE
		End If
		rstmp.Close()
		Set rstmp = Nothing	
	End Function
	
	Function estacad_txt(codcli)
		Set rstmp = Conn.Execute("select estacad from mt_alumno where rut = '" + codcli + "';")
		estacad_txt = rstmp("estacad")
		rstmp.Close()
		Set rstmp = Nothing	
	End Function
	
	Function estfinan_txt(codcli)
		Set rstmp = Conn.Execute("select estfinan from mt_alumno where rut = '" + codcli + "';")
		estfinan_txt = rstmp("estfinan")
		rstmp.Close()
		Set rstmp = Nothing	
	End Function
	
	
	Function estfinan(codcli)
		Set rstmp = Conn.Execute("select estfinan from mt_alumno where rut = '" + codcli + "' and estfinan = 'AL DIA';")
		'If rstmp("estfinan") = "AL DIA" Then
		If (not rstmp.eof) Then
			estfinan = TRUE
		Else
			estfinan = FALSE
		End If
		rstmp.Close()
		Set rstmp = Nothing	
	End Function
	
	Function nivel(codcli)
		Set rstmp = Conn.Execute("select nivel from mt_alumno where rut = '" + codcli + "';")
		nivel = rstmp("nivel")
		rstmp.Close()
		Set rstmp = Nothing	
	End Function	
	
	Function jornada(codcli)
		Set rstmp = Conn.Execute("select jornada from mt_alumno where rut = '" + codcli + "';")
		jornada = rstmp("jornada")
		rstmp.Close()
		Set rstmp = Nothing	
	End Function
	
	Function ano_mat(codcli)
		Set rstmp = Conn.Execute("select ano_mat from mt_alumno where rut = '" + codcli + "';")
		ano_mat = rstmp("ano_mat")
		rstmp.Close()
		Set rstmp = Nothing	
	End Function
	
	Function numero_ecas(codcli)
		Set rstmp = Conn.Execute("select codcli from mt_alumno where rut = '" + codcli + "';")
		numero_ecas = rstmp("codcli")
		rstmp.Close()
		Set rstmp = Nothing	
	End Function
	
	Function validez(id_tipo)
		Set rstmp = Conn.Execute("select validez from certificados_tipo where id_tipo = '" + id_tipo + "';")
		validez = rstmp("validez")
		rstmp.Close()
		Set rstmp = Nothing	
	End Function
	
	Function periodo_mat(codcli)
		Set rstmp = Conn.Execute("select periodo_mat from mt_alumno where rut = '" + codcli + "';")
		periodo_mat = rstmp("periodo_mat")
		rstmp.Close()
		Set rstmp = Nothing	
	End Function
	
	Function genero(codcli)
		Set rstmp = Conn.Execute("select SEXO from MT_CLIENT where CODCLI = '" + codcli + "';")
		genero = rstmp("sexo")
		rstmp.Close()
		Set rstmp = Nothing	
	End Function
%>