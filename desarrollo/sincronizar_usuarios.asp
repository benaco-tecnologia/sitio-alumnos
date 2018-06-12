<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%On Error Resume Next%>
<%
	Response.Clear()
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
 	Server.ScriptTimeout = 360
	Response.Buffer = true
	Response.Write("Inicio ejecucion sincronizacion usuarios<br>")
	Response.Write(now()&"<br>")
	registros = 0
	servidorMy = "192.168.4.10"
	nombreUsuarioMy = "matricula"
	passwordMy = "dtb01s"
	baseDatosMy = "ecasvirtual"
	
	servidor = "192.168.1.194"
	nombreUsuario = "matricula"
	password = "dtb01s"
	baseDatos = "matricula"
	
	strDsn = "ecasvirtual"
	ANIO = year(now())
	
	If (Session("ConnMySQL") = "") Then
		strConnMySQL = "DSN=" & strDsn &  ";UID="&nombreUsuarioMy&";PWD="&passwordMy&"" 
		Set Session("ConnMySQL") = Server.CreateObject("ADODB.Connection")
		Session("ConnMySQL").Open(strConnMySQL)								
	End If
	
	Set Session("Conn") = Server.CreateObject("ADODB.Connection")
	Session("Conn").Open "DRIVER={SQL Server}; SERVER="& servidor &"; DATABASE="& baseDatos &"; UID="& nombreUsuario &"; PWD="& password &""
	
	Set Session("ConnSQLServer") = Server.CreateObject("ADODB.Connection")
	Session("ConnSQLServer").Open "DRIVER={SQL Server}; SERVER="& servidor &"; DATABASE="& baseDatos &"; UID="& nombreUsuario &"; PWD="& password &""

	strQuery = "SELECT A.RUT, SUBSTRING(A.RUT, 0, 6) AS CLAVE, A.CODCLI AS NUMECAS, C.NOMBRE, C.PATERNO + ' ' + C.MATERNO AS APELLIDOS, C.MAIL, 'ALUMNO' AS TIPO " & _
				"FROM MT_ALUMNO A, MT_CLIENT C " & _
				"where (A.estacad <> 'ELIMINADO' AND A.estacad <> 'SUSPENDIDO') " & _
					"AND C.CODCLI = A.RUT " & _
					"AND A.ANO = " & ANIO & " " & _
				"UNION " & _
				"SELECT RUT, SUBSTRING(RUT, 0, 6) AS CLAVE, RUT AS NUMECAS, NOMBRES AS NOMBRE, AP_PATER + ' ' + AP_MATER AS APELLIDOS, MAIL, 'PROFESOR' AS TIPO " & _
				"FROM RA_PROFES " & _
				"ORDER BY RUT;"
				'"AND (ECASVIRTUAL IS NULL OR ECASVIRTUAL = '' OR ECASVIRTUAL = '') " & _
	'"WHERE (ECASVIRTUAL IS NULL OR ECASVIRTUAL = '' OR ECASVIRTUAL = '') " & _
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open strQuery, Session("Conn"), 1, 1
	Response.Flush
	
	Response.Write(rs.RecordCount & " usuarios encontrados para sincronizar<br>")
	do while not rs.EOF
		
		query = "SELECT username FROM mdl_user WHERE username = '"&rs("RUT")&"';"
		Set rs1 = Server.CreateObject("ADODB.Recordset")
		rs1.Open query, Session("ConnMySQL"), 1, 1
		
		if (rs1.EOF) then
			strQuery = "INSERT INTO mdl_user (username, password, idnumber, firstname, lastname, email, confirmed, mnethostid, tipo_usuario) " & _
				   "VALUES ('"&rs("RUT")&"',MD5('"&rs("CLAVE")&"'),'"&rs("NUMECAS")&"','"&rs("NOMBRE")&"','"&rs("APELLIDOS")&"','"&rs("MAIL")&"',1,3, '"&rs("TIPO")&"');"
			Session("ConnMySQL").Execute(strQuery)
			Session("ConnSQLServer").Execute("UPDATE MT_ALUMNO SET ECASVIRTUAL = '1' WHERE RUT = '"&rs("RUT")&"';")
			Session("ConnSQLServer").Execute("UPDATE RA_PROFES SET ECASVIRTUAL = '1' WHERE RUT = '"&rs("RUT")&"';")
			Response.Write("Usuario " &rs("RUT")& " ingresado<br>")
		Else
			Response.Write("Usuario " &rs1("username")& " ya existe<br>")
		End if
		
		Set rs1 = Nothing
		
		rs.MoveNext
	loop
	
	Response.Write(registros & " usuarios sincronizados<br>")
	'Session("ConnMySQL").Close()
	'Session("Conn").Close()
	Response.Write(now()&"<br>")
	Response.Write("Fin ejecucion sincronizacion usuarios<br>")
%>