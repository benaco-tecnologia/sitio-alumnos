<%
	Response.Clear()
	Response.CacheControl = "no-cache"
	Response.AddHeader "Pragma", "no-cache"
	Response.Expires = -1
 	Server.ScriptTimeout = 360
	Response.Buffer = true
		
	'dim exeTotal
	'exeTotal = Timer()
	'dim exeSQL
%>
<%
   '#DEFINE ADEMPTY               0
   '#DEFINE ADTINYINT            16
   '#DEFINE ADSMALLINT            2
   '#DEFINE ADINTEGER            3
   '#DEFINE ADBIGINT            20
   '#DEFINE ADUNSIGNEDTINYINT      17
   '#DEFINE ADUNSIGNEDSMALLINT      18
   '#DEFINE ADUNSIGNEDINT         19
   '#DEFINE ADUNSIGNEDBIGINT      21
   '#DEFINE ADSINGLE            4
   '#DEFINE ADDOUBLE            5
   '#DEFINE ADCURRENCY            6
   '#DEFINE ADDECIMAL            14
   '#DEFINE ADNUMERIC            131
   '#DEFINE ADBOOLEAN            11
   '#DEFINE ADERROR               10
   '#DEFINE ADUSERDEFINED         132
   '#DEFINE ADVARIANT            12
   '#DEFINE ADIDISPATCH            9
   '#DEFINE ADIUNKNOWN            13
   '#DEFINE ADGUID               72
   '#DEFINE ADDATE               7
   '#DEFINE ADDBDATE            133
   '#DEFINE ADDBTIME            134
   '#DEFINE ADDBTIMESTAMP         135
   '#DEFINE ADBSTR               8
   '#DEFINE ADCHAR               129
   '#DEFINE ADVARCHAR            200
   '#DEFINE ADLONGVARCHAR         201
   '#DEFINE ADWCHAR               130
   '#DEFINE ADVARWCHAR            202
   '#DEFINE ADLONGVARWCHAR         203
   '#DEFINE ADBINARY            128
   '#DEFINE ADVARBINARY            204
   '#DEFINE ADLONGVARBINARY         205
   '#DEFINE ADCHAPTER            136
  
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>SQL Query Test</title>
<style type="text/css" media="all">
body
{
	font-size: .9em;
	font-family: Arial, Helvetica, sans-serif;
}
</style>
</head>

<body>
<form action="<%Response.Write(Request.ServerVariables("SCRIPT_NAME"))%>" method="post">
<textarea name="sql_query" rows="10" style="width:100%;"><%
If Request.Form("sql_query") <> "" then
	response.Write(Request.Form("sql_query"))
end if
%></textarea><br />
<input name="sql_password" type="password" />
<input name="sql_query_submit" type="submit" value="Aceptar" />
</form>
<%
If Request.Form("sql_password") = "dtb01s" then
	' Declaracion constantes
	Const servidor = "192.168.1.194"
	Const nombreUsuario = "matricula"
	Dim password
	password = "dtb01s" '"dtb01s"
	Const baseDatos = "Matricula"
	
	'conexion a base datos Estadisticas
	Set conn = Server.CreateObject("ADODB.Connection")
	conn.Open "DRIVER={SQL Server}; SERVER="& servidor &"; DATABASE="& baseDatos &"; UID="& nombreUsuario &"; PWD="& password &""
	
	' Objeto Recorderset -> Obtencion de registro
	Set RS = Server.CreateObject("ADODB.Recordset")
	strSQL = Request.Form("sql_query")
	 
  'for long winded queries, this will write out the response buffers 
  Response.Flush()
if strSQL <> "" then	' execute the SQL if it's not empty
	'exeSQL = Timer()
	RS.Open strSQL, conn, 1, 1
	'exeSQL = Timer() - exeSQL
	'
	'Response.Write("La consulta se ejecutó en " & CStr(exeSQL) & " segundos y generó ")
	
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
		 	next
		  	Response.Write("</tr>")
			Response.Flush()
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
end if
  Set RS = Nothing 
 ' Response.Write("La página tardó " & CStr((Timer() - exeTotal)) & " segundos")
 end if
%>
</body>
</html>
