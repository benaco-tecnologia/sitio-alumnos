
<!--#include file="addon/includes/funciones.inc" -->
<!--#include file="addon/includes/libreria.asp" -->
<%

    'Response.Buffer = True
	Dim mySmartUpload
	Dim i, f, Archivos(2)
	
	ID_WEB = "53"
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set mySmartUpload = Server.CreateObject("Cedentes.AdmCedentes")
	mySmartUpload.mostrar ID_WEB 
	
	'response.Buffer = false
	'set fso = nothing 
	'Set mySmartUpload = nothing 
	
	response.write("TRANSACCION ok")
	RESPONSE.END()

%>