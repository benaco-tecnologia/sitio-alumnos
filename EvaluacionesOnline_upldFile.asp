
<html>  
<head>  
<title></title>  
</head>  
<body>  
<%  
Dim mySmartUpload  
Dim StrSql




  
'*** Create Object ***'
Set fso = CreateObject("Scripting.FileSystemObject")  
Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")  

'*** Upload Files ***'  
mySmartUpload.Upload  

if(mySmartUpload.Files("Docto").FileName = "") then 
%>

<script languaje="javascript"> 
 alert('Debe Seleccionar un documento.');
   document.location=('./EvaluacionOnline.asp');
</script>

<%
end if


Codcli = mySmartUpload.form("CodCli")
Conter =  mySmartUpload.form("Conter")
Atrasado =  mySmartUpload.form("Atrasado")
Raiz1 = Server.MapPath("Doctos")& "\EvaluacionesOnline\"
Raiz2 = Server.MapPath("Doctos")& "\EvaluacionesOnline\" & Conter & "\"

'*** Crea carpeta si no existe ***
if (Not fso.FolderExists(Raiz1)) then
     Set fol = fso.CreateFolder(Raiz1)
   end if

if (Not fso.FolderExists(Raiz2)) then
     Set fol = fso.CreateFolder(Raiz2)
   end if
'*** Obtiene formato del archivo ***
x = split(mySmartUpload.Files("Docto").FileName, ".")
formato = x(UBound(x))
NomArchivo = Conter & "_" & Codcli & "." & formato

strsql = " insert into ca_evaluacionesOnline (ID_CONTER,CODCLI,ARCHIVO,ATRASADO,FECHA_SUBIDA) values ("
			StrSql = StrSql &  Conter  &", '"& Codcli & "', '" & NomArchivo & "', '" & Atrasado & "', '" & Date & "')"

Session("Conn").execute StrSql	


'*** Upload file ***'  
mySmartUpload.Files("Docto").SaveAs(Raiz2 & NomArchivo)

%>

<script languaje="javascript"> 
 alert('Archivo subido con exitosamente.');
   document.location=('./EvaluacionOnline.asp');
</script>

<%
  
Set mySmartUpload = Nothing  
%>  
</body>  
</html>  