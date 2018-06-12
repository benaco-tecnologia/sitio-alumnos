<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file = "adovbs.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="md5.asp" -->
<%
	Response.Expires = -1
	
	Const servidor = "192.168.1.8"
	Const nombreUsuario = "matricula"
	Const password = "dtb01s"
	Const baseDatos = "matricula"
	
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
	'Conn.Execute("delete from mt_webmail_new;")

	'StrSql = "select mt_webmail.rut,mt_webmail.passweb,SUBSTRING(mt_webmail.rut, 1, 5) from mt_webmail order by mt_webmail.rut ASC"
   	'StrSql = "Select a.rut,SUBSTRING(a.rut, 1, 5), c.nombre as nom, c.paterno as app, a.passweb, a.passchanged, b.codcli, b.codsede,b.codcarpr,b.codpestud, b.ano_mat, b.ano, b.jornada,b.codcarpr From mt_webmail a with (nolock), mt_alumno b with (nolock), mt_client c with (nolock) Where a.rut = b.rut  and c.codcli = a.rut order by b.ano desc;"
   
	'Actualiza.Open strSql,Conn
	'do until Actualiza.EOF
	'	Conn.Execute("INSERT INTO mt_webmail_new VALUES ('"&Actualiza.Fields(0)&"', '"&Encripta(Actualiza.Fields(1))&"');")
	'	Actualiza.MoveNext	
	'loop
	
	'Limpiamos objetos
	'Actualiza.Close
	
	StrSql = "Select rut, clave from mt_webmail_new order by rut desc;"
   
	Actualiza.Open strSql,Conn
	do until Actualiza.EOF
		Response.Write("UPDATE MT_WEBMAIL SET PASSWEB = '"&Actualiza.Fields(1)&"', PASSCHANGED='S' WHERE RUT = '"&Actualiza.Fields(0)&"';")
		Actualiza.MoveNext	
	loop
	Conn.Close
	
%>