<!--#INCLUDE FILE="include/conexion.inc" -->
<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%
   if Session("CodCli") = "" then
       Response.Redirect("saltoinicio.htm")
   end if
   
   Codcli = Session("Codcli")
   Ano = Session("Ano")
   Periodo = Session("Periodo")

   StrSql = "Delete from ra_solici "
   StrSql = StrSql & " where Codcli = '" & Codcli & "' "
   StrSql = StrSql & " And ano = " & ano & " and Periodo = " & Periodo 
   Session("Conn").execute StrSql

   StrSql = "Insert into ra_solici (Codcli, CodRamo, CodSecc, Ano, Periodo, Glosa, CodCarr, Puntaje, RamoEquiv,CodSeccOri ) "
   StrSql = StrSql & "Select Codcli, CodRamo, CodSecc, Ano, Periodo, Glosa, CodCarr, 0, Ramoequiv,CodSecc from tmpsolici "
   StrSql = StrSql & " where Codcli = '" & Codcli & "' "
   StrSql = StrSql & " And ano = " & ano & " and Periodo = " & Periodo 
   
   Session("Conn").execute StrSql

   StrSql = " Update ra_solici set PerInscrip = " & SESSION("PER_ID")
   StrSql = StrSql & " where Codcli = '" & Codcli & "' "
   StrSql = StrSql & " And ano = " & ano & " and Periodo = " & Periodo 
  
   Session("Conn").execute StrSql
   
   'SE COMENTA PARA PERMITIR AL ALUMNO QUE SE PUEDA 
   'INGRESAR A LA SOLICITUD ESPECIAL.
   
   'StrSql = "Insert into sis_reg_solicitud (Codcli, fecha) values ('" & Codcli & "', getdate()) "
  
   'Session("Conn").execute StrSql
   
   'FIN.
   
  
   
%>
<body bgcolor="#FFFFFF" text="#000000">
<script>
  //parent.location = "resultado.htm"
  //parent.parent.frames("leftFrame").location = "f-izq.asp";
  window.location.href='inscrip-asigna.asp';
 	
</script>

</body>
</html>

<!--#INCLUDE file="include/desconexion.inc" -->

