
<!--#include file="addon/includes/funciones.inc" -->
<!--#include file="addon/includes/libreria.asp" -->
<%

    'Response.Buffer = True
	Dim WSH
	Dim enlace
	Dim objScriptExec,strExeOut

	ID_WEB ="53"

	'enlace="C:\inetpub\wwwroot\umcean\mt0406.exe" & " "& ID_WEB 
	'enlace="C:\mt0406.exe" & " "& ID_WEB 
	
	'response.write(enlace)
	'Set WSH = CreateObject("WScript.Shell")
	'Set objScriptExec = WSH.Exec(enlace) 'Ejecuto el exe pasandole el nombre del archivo con sus parametros
	
	set wshell = CreateObject("WScript.Shell") 
    wshell.run "C:\mt0406.exe" & " "& ID_WEB 
    set wshell = nothing 
	'strExeOut = objScriptExec.StdOut.ReadAll 'Coge la salida del ejecutable en este caso el enlace encriptado
	'response.redirect(strExeOut) 'Me dirijo a la direccion obtenida
	'response.Buffer = false
	'set fso = nothing 
	'Set mySmartUpload = nothing 
	
	response.write("TRANSACCION ok")
	RESPONSE.END()

%>