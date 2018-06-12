<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file = "adovbs.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%On Error Resume Next%>
<%
	Response.Clear()
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
 	Server.ScriptTimeout = 360
	Response.Buffer = true
	
	Const servidor = "192.168.1.194"
	Const nombreUsuario = "matricula"
	Const password = "dtb01s"
	Const baseDatos = "matricula"
	
	Dim Conn, StrSql, rs
	
	'"Provider=SQLOLEDB; " _  & "Data Source=source_name; " _  & "Database=db_name; " _  & "UID=userid; " _  & "PWD=password;"
	Dim aConnectionString : aConnectionString = "DRIVER={SQL Server}; "_
  											  & "SERVER="& servidor &"; "_
											  & "DATABASE="& baseDatos &"; "_
											  & "UID="& nombreUsuario &"; "_
											  & "PWD="& password &""
									
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Set Conn2 = Server.CreateObject("ADODB.Connection")
	Conn.Open aConnectionString
	Conn2.Open aConnectionString
	Set rs = Server.CreateObject("ADODB.RecordSet") 								
	Set objConnMySQL = Server.CreateObject("ADODB.Connection")
		objConnMySQL.Open(Session("strConnMySQL"))
		
	StrSql = "SELECT rut FROM mt_webmail WHERE passweb is null or passweb = '';"
		   
	rs.Open strSql,Conn
	usuarios = 0
	
	do until rs.EOF
		cincoDigitos = Left(rs("rut"), 5)
		Conn2.Execute("UPDATE mt_webmail SET passweb = '" & Encripta(cincoDigitos) & "', passchanged = 'S' WHERE rut = '" & rs("rut") & "';")
		'response.write("UPDATE mt_webmail SET passweb = '" & Encripta(Left(rs("rut"), 5)) & "', passchanged = 'S' WHERE rut = '" & rs("rut") & "';<br />")
		If Err.Number <> 0 Then
			Error.Clear
		Else
			usuarios = usuarios + 1
		End If
		
		Set rsm = objConnMySQL.Execute("UPDATE mdl_user SET password = MD5('"&cincoDigitos&"') WHERE username = '"&rut_alumno&"';")
		Set rsm = Nothing
		
		rs.MoveNext	
	loop
	rs.Close
	Conn.Close
	Conn2.Close
	
	Response.Write(CStr(usuarios))
%>