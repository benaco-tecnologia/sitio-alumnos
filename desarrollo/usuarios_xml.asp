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
	
	Dim Conn, StrSql, rs
	
	'"Provider=SQLOLEDB; " _  & "Data Source=source_name; " _  & "Database=db_name; " _  & "UID=userid; " _  & "PWD=password;"
	Dim aConnectionString : aConnectionString = "DRIVER={SQL Server}; "_
  											  & "SERVER="& servidor &"; "_
											  & "DATABASE="& baseDatos &"; "_
											  & "UID="& nombreUsuario &"; "_
											  & "PWD="& password &""
									
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open aConnectionString
	Set rs = Server.CreateObject("ADODB.RecordSet") 								
	
	StrSql = "select mt_webmail.rut,mt_client.dig,mt_webmail.passweb,mt_webmail.passchanged,mt_webmail.password, mt_client.codapod from mt_webmail, mt_client "_
		   & "where mt_client.codcli=mt_webmail.rut and mt_webmail.passweb <> '' order by mt_webmail.rut ASC"
		   
	rs.Open strSql,Conn
	
	response.write("<usuarios>") 
	do until rs.EOF
		Response.Write("<usuario>")
		response.write("<rut>" & rs("rut") & "</rut>")
		response.write("<passweb><![CDATA[" & rs("passweb") & "]]></passweb>")
		response.write("<passweb_desentriptado><![CDATA[" & Desencripta(rs("passweb")) & "]]></passweb_desentriptado>")
		Response.Write("</usuario>")
		rs.MoveNext	
	loop
	response.write("</usuarios>") 
	'Limpiamos objetos
	rs.Close
	Conn.Close
%>