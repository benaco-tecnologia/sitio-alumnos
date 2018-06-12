<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file = "adovbs.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%
	Response.ContentType = "text/xml"
	response.write("<?xml version='1.0'?>") 
	Response.Expires = -1
	
	Const servidor = "192.168.1.8"
	Const nombreUsuario = "matricula"
	Const password = "dtb01s"
	Const baseDatos = "matricula"
	
	Dim Conn, StrSql, rs, semestre, anio
	
	semestre = Request.QueryString("semestre")
	anio = Request.QueryString("anio")
	
	'"Provider=SQLOLEDB; " _  & "Data Source=source_name; " _  & "Database=db_name; " _  & "UID=userid; " _  & "PWD=password;"
	Dim aConnectionString : aConnectionString = "DRIVER={SQL Server}; "_
  											  & "SERVER="& servidor &"; "_
											  & "DATABASE="& baseDatos &"; "_
											  & "UID="& nombreUsuario &"; "_
											  & "PWD="& password &""
									
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open aConnectionString
	Set rs = Server.CreateObject("ADODB.RecordSet") 								
	
	StrSql = "select codcli, codramo, ano, periodo, ncert1, ncert2, nej, ne "_
		   & "from ra_nota "_
		   & "where ano = '"&anio&"' AND periodo = '"&semestre&"' " _
		   & "ORDER BY codramo ASC"
		   
	rs.Open strSql,Conn
	
	response.write("<notas>") 
	do until rs.EOF
		Response.Write("<nota>")
		response.write("<ramo>" & rs("codramo") & "</ramo>")
		response.write("<numeroecas>" & rs("codcli") & "</numeroecas>")
		response.write("<primera_departamental>" & rs("ncert1") & "</primera_departamental>")
		response.write("<segunda_departamental>" & rs("ncert2") & "</segunda_departamental>")
		response.write("<ejercicios>" & rs("nej") & "</ejercicios>")
		response.write("<examen>" & rs("ne") & "</examen>")
		Response.Write("</nota>")
		rs.MoveNext	
	loop
	response.write("</notas>") 
	'Limpiamos objetos
	rs.Close
	Conn.Close
%>