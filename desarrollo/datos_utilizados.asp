<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Untitled Document</title>
</head>

<body>
<%
	
	Const servidor = "192.168.1.192"
	Const nombreUsuario = "matricula"
	Const password = "dtb01s"
	Const baseDatos = "matricula"									
	
	strSQL = "select name from matricula..sysobjects where xtype = 'U' order by name;"
	sqlStore = "sp_spaceused ''"
		
	Set Conn = Server.CreateObject("ADODB.Connection")
	Set RS = Server.CreateObject("ADODB.Recordset")

	Conn.Open "DRIVER={SQL Server}; SERVER="& servidor &"; DATABASE="& baseDatos &"; UID="& nombreUsuario &"; PWD="& password &""
	
	RS.Open strSQL, Conn, 1, 1
		
  if RS.State = 1 then	'if the recordset has rows		
	if RS.RecordCount > 0 then
		Response.Write(RS.RecordCount & " resultados<br />")
		Response.Write("<table border=1 cellspacing=0 cellpadding=""3"">")
		'show the column names
		Response.Write("<tr bgcolor=LightSteelBlue>")
		response.Write("<td>#</td>")
		for each f in RS.Fields
			Response.Write("<td")
			Response.Write("><b>" & f.Name & "</b></td>")
		next
	
		Response.Write("</tr>")
			
		'show the rows
		row_number = 1
		do while not RS.EOF
				
			Response.Write("<tr bgcolor=White>")
			response.Write("<td align=""right"">" & CStr(row_number) & "</td>")
			row_number = row_number + 1
		  	for each f in RS.Fields
				Response.Write("<td")
				if (f.Type = 3) then
					response.Write(" align=""right""")
				end if
				Response.Write(">" & f.Value & "</td>")
				set rsaux = conn.Execute("sp_spaceused '" & f.Value & "'")
				for each fa in rsaux.Fields
					Response.Write("<td>"&fa.Value&"</td>")
				next
				set rsaux = nothing
		 	next
		  	Response.Write("</tr>")
			RS.MoveNext
		loop
	else
		response.Write("No se encontraron resultados!")
	end if
  else
	'DML was performed
	Response.Write("<tr bgcolor=White><td><b>")
	Response.Write("Command Completed Successfully</b>")
	Response.Write("</td></tr>")
  end if

  Response.Write("</table>")
  Set RS = Nothing
	Conn.Close
	set Conn = Nothing
%>
</body>
</html>
