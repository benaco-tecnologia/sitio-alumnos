<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<% Response.Expires = -1 %>
<%
	Response.Clear()
	Response.AddHeader "Content-Disposition", "inline;filename=alumnos-inscripcion-ramos.csv"
	Response.ContentType = "text/csv"
	
	Response.Write ("username;course1;type1" & vbCrLf)

	Const servidor = "192.168.1.194"
	Const nombreUsuario = "sa"
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
	
	StrSql = "select a.rut, n.ramoequiv from ra_nota n, mt_alumno a where n.ano = DATEPART(YEAR, GETDATE()) and n.periodo = "&periodo&" and n.codcli = a.codcli and a.estacad = 'VIGENTE' and a.matriculado = 'S';"
   
	Actualiza.Open strSql,Conn
	do until Actualiza.EOF
		Response.Write(Actualiza.Fields(0)&";"&Actualiza.Fields(1)&";1" & vbCrLf)
		Actualiza.MoveNext	
	loop
	Actualiza.Close()
	
	StrSql = "SELECT CODPROF, CODRAMO FROM RA_SECCIO WHERE (ANO = DATEPART(YEAR, GETDATE())) AND (PERIODO = "&periodo&")"
   
	Actualiza.Open strSql,Conn
	do until Actualiza.EOF
		Response.Write(Actualiza.Fields(0)&";"&Actualiza.Fields(1)&";2" & vbCrLf)
		Actualiza.MoveNext	
	loop
	Actualiza.Close()
	Conn.Close()
%>