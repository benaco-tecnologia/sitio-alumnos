<% Option Explicit

Const servidor = "192.168.1.8"
	Const nombreUsuario = "matricula"
	Const password = "dtb01s"
	Const baseDatos = "matricula"
	
	Dim Conn, StrSql, Actualiza, J
	
	'"Provider=SQLOLEDB; " _  & "Data Source=source_name; " _  & "Database=db_name; " _  & "UID=userid; " _  & "PWD=password;"
	Dim strDataConn : strDataConn = "DRIVER={SQL Server}; "_
  											  & "SERVER="& servidor &"; "_
											  & "DATABASE="& baseDatos &"; "_
											  & "UID="& nombreUsuario &"; "_
											  & "PWD="& password &""
											  
Const Filename = "/correos-usuarios-ecas.csv"    ' file to read
Const ForReading = 1, ForWriting = 2, ForAppending = 3
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

' Create a filesystem object
Dim FSO
set FSO = server.createObject("Scripting.FileSystemObject")

' Map the logical path to the physical system path
Dim Filepath
Filepath = Server.MapPath(Filename)
dim dataConn
Set dataConn = Server.CreateObject("ADODB.Connection")
dim sqlDatos
dataConn.Open strDataConn

if FSO.FileExists(Filepath) Then
	Dim TextStream
    Set TextStream = FSO.OpenTextFile(Filepath, ForReading, False, TristateUseDefault)

    ' Read file in one hit
    Dim Contents
	Contents = TextStream.ReadLine
    Contents = TextStream.ReadAll
    'Response.write "<pre>" & Contents & "</pre><hr>"
    TextStream.Close
    Set TextStream = nothing
	
	Dim arrContents
	arrContents = Split(Contents, ";", -1, 1)
	dim i
	For i = 0 to ubound(arrContents)
		'response.write(InStrRev("Test"&arrContents(i+6), "Alumno"))
		'Response.Write i&": " & arrContents(i+2) & ";" & arrContents(i+3) & ";" & arrContents(i+4) & ";" & arrContents(i+5) & ";" & arrContents(i+6) & "<br />"
		
		if (InStrRev("Test"&arrContents(i+6), "Alumno") <> 5) then
			sqlDatos = "update RA_PROFES set mail = '"&arrContents(i+5)&"' where RUT = '"&arrContents(i+2)&"'"
			'dataConn.Execute(sqlDatos)
			'response.write(sqlDatos)
		else
			sqlDatos = "update mt_client set mail = '"&arrContents(i+5)&"' where codcli = '"&arrContents(i+2)&"'"
			'dataConn.Execute(sqlDatos)
			'response.write(sqlDatos)
		end if
		'response.write("<br>")
		
		if (i + 6 >= ubound(arrContents)) then
			i = ubound(arrContents)
		else
			i = i + 5
		end if
	next
    Response.write("<h4>Correos importados correctamente.</h4>")
Else

    Response.Write "<h3><i><font color=red> File " & Filename &_
                       " does not exist</font></i></h3>"

End If

Set FSO = nothing
%>