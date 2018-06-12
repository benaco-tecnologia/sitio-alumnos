<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%On Error Resume Next%>
<%ScriptTimeOut = 100000%>
<%
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

	If (Session("ConnMySQL") = "") Then
		strConnMySQL = "DSN=" & strDsn &  ";UID="&nombreUsuarioMy&";PWD="&passwordMy&"" 
		Set Session("ConnMySQL") = Server.CreateObject("ADODB.Connection")
		Session("ConnMySQL").Open(strConnMySQL)								
	End If
	
	aConnectionString = "DRIVER={SQL Server}; SERVER="& servidor &"; DATABASE="& baseDatos &"; UID="& nombreUsuario &"; PWD="& password &""
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open aConnectionString

	Set rs = Session("ConnMySQL").Execute("SELECT username FROM mdl_user WHERE idnumber = '' OR idnumber is null;")
	do until (rs.EOF)
		Set rs_p = Conn.Execute("SELECT rut FROM ra_profes WHERE rut = '" & trim(rs("username")) & "';") ' Chequeamos si es profesor
		If (Not rs_p.EOF) Then ' Es profesor
			Session("ConnMySQL").Execute("UPDATE mdl_user set idnumber = '" & trim(rs_p("rut")) & "' WHERE username = '"&rs("username")&"';")
			'registros = registros + 1
			Set rs_p = Nothing
		Else
			Set rs_a = Conn.Execute("SELECT codcli FROM mt_alumno WHERE rut = '" & trim(rs("username")) & "';") ' Buscamos numero ecas
			If (Not rs_a.EOF) Then ' Es alumno
				Session("ConnMySQL").Execute("UPDATE mdl_user set idnumber = '" & trim(rs_a("codcli")) & "' WHERE username = '"&rs("username")&"';")
				registros = registros + 1
				Set rs_a = Nothing
			End If
		End If
		rs.MoveNext
	loop
	Response.Write(registros)
%>