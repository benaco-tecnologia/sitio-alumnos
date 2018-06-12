<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!-- #include file = "adovbs.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%
	Response.Expires = -1
	
	Const servidor = "192.168.1.192"
	Const nombreUsuario = "matricula"
	Const password = "dtb01s"
	Const baseDatos = "matricula"
	
	Dim conn, rs, SQL, RecsAffected, qr, bgcolor, rutalumno, currentpassword, newpassword_copy, newpassword
	
	'"Provider=SQLOLEDB; " _  & "Data Source=source_name; " _  & "Database=db_name; " _  & "UID=userid; " _  & "PWD=password;"
	Dim aConnectionString : aConnectionString = "DRIVER={SQL Server}; "_
  											  & "SERVER="& servidor &"; "_
											  & "DATABASE="& baseDatos &"; "_
											  & "UID="& nombreUsuario &"; "_
											  & "PWD="& password &""
									
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open aConnectionString
	Set Actualiza = Server.CreateObject("ADODB.RecordSet") 								
	
	StrSql = "select mt_webmail.rut,mt_client.dig,mt_webmail.passweb,mt_webmail.passchanged,mt_webmail.password, mt_client.codapod from mt_webmail, mt_client "_
		   & "where mt_client.codcli=mt_webmail.rut order by mt_webmail.rut ASC"
		   
	Actualiza.Open strSql,Conn
	
	response.Write("<table><tr>")
	For i = 0 to Actualiza.Fields.Count - 1
       Response.Write("<td>"&Actualiza(i).Name&"</td>")
	   if i = 2 then
		   Response.Write("<td>passweb desentriptado</td>")
	   end if
   Next
   response.Write("</tr>")
	do until Actualiza.EOF
		 Response.Write("<TR>")
		 for I=0 to Actualiza.Fields.Count-1
		 	Response.Write("<TD>"&Actualiza.Fields(I)&"</TD>")
			if I = 2 and Actualiza.Fields(I) <> "" then
				response.Write("<td>"&Desencripta(Actualiza.Fields(I))&"</td>")
			end if 			
		 next
		 Actualiza.MoveNext
	
	loop
	Response.Write("</TABLE>")
	
	'Limpiamos objetos
	Actualiza.Close

	response.Write("</table>")
%>