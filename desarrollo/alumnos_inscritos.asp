<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Response.Expires = -1 %>
<%
	Response.Clear()
	Response.AddHeader "Content-Disposition", "inline;filename=alumnos-activos.csv"
	Response.ContentType = "text/csv"

	Response.Write("username;password;firstname;lastname;idnumber" & vbCrLf)
	
	Const servidor = "192.168.1.194"
	Const nombreUsuario = "sa"
	'Const password = "dtb01s"
	Const password = "reflex"
	Const baseDatos = "umasnet"
	
	Dim Conn, StrSql, Actualiza, J
	
	'"Provider=SQLOLEDB; " _  & "Data Source=source_name; " _  & "Database=db_name; " _  & "UID=userid; " _  & "PWD=password;"
	Dim aConnectionString : aConnectionString = "DRIVER={SQL Server}; "_
  											  & "SERVER="& servidor &"; "_
											  & "DATABASE="& baseDatos &"; "_
											  & "UID="& nombreUsuario &"; "_
											  & "PWD="& password &""
									
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open aConnectionString
	Set Actualiza = Server.CreateObject("ADODB.RecordSet") 	
	
	dim periodo
	If (DateDiff("d", now(), "01-08-"&year(now())) > 1) Then
		periodo = "1"
	Else
		periodo = "2"
	End If
	
	'StrSql = "select distinct a.rut, c.nombre, c.paterno + ' '+c.materno, a.codcli, c.mail  from ra_nota n, mt_alumno a, mt_client c where n.ano = DATEPART(YEAR, GETDATE()) and n.periodo = "&periodo&" and n.codcli = a.codcli and a.estacad = 'VIGENTE' and a.matriculado = 'S' and a.rut = c.codcli"
	StrSql = "select distinct a.rut, c.nombre, c.paterno + ' '+c.materno, a.codcli, c.mail  from mt_alumno a, mt_client c where a.ano = DATEPART(YEAR, GETDATE()) and a.periodo = "&periodo&" and a.estacad = 'VIGENTE' and a.matriculado = 'S' and a.rut = c.codcli"
   
	Actualiza.Open strSql,Conn
	do until Actualiza.EOF
		Response.Write(Actualiza.Fields(0)&";"&Left(Actualiza.Fields(0), 5)&";"&Actualiza.Fields(1)&";"&Actualiza.Fields(2)&";"&Actualiza.Fields(3) & vbCrLf)
		Actualiza.MoveNext	
	loop
	Actualiza.Close()
	
	StrSql = "SELECT DISTINCT RUT, NOMBRES, AP_PATER + ' ' + AP_MATER, RUT, mail FROM RA_PROFES"
   
	Actualiza.Open strSql,Conn
	do until Actualiza.EOF
		Response.Write(Actualiza.Fields(0)&";"&Left(Actualiza.Fields(0), 5)&";"&Actualiza.Fields(1)&";"&Actualiza.Fields(2)&";"&Actualiza.Fields(3) & vbCrLf)
		Actualiza.MoveNext	
	loop
	Actualiza.Close()
	Conn.Close()
%>