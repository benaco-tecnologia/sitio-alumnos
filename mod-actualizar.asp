<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="include/mail.inc" -->
<!--#INCLUDE FILE="include/audita.inc" -->
<!--#INCLUDE FILE="include/Cupos.inc" -->

<%
  if Session("CodCli") = "" then
  		Response.Redirect("saltoinicio.htm")
  end if
%>

<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%
Dim CodCli,CodSede,CodPestud,Codcarr,ano,Periodo, Confirma

Seccion = Request("S")
Codramo = Request("C")  
PrioridadSol  = request("PR")
Codcli = Session("Codcli")
Ano = Session("Ano")
Periodo = Session("Periodo")
CodSede = Session("CodSede")
CodPestud = Session("CodPestud")

	'response.write(Seccion) + "<br>"
	'response.write(Codramo) + "<br>"
	'response.write(PrioridadSol) + "<br>"
	'response.write(Codcli) + "<br>"
	'response.write(Ano) 
	'response.write(Periodo)
	'response.End()


	StrSql = "Select CodRamo, CodSecc, RamoEquiv from ra_carga_log with (nolock)"
	StrSql = StrSql & " where inscrito = 'S' "
	StrSql = StrSql & " and Codcli = '" & Codcli & "' and codramo = '" & Codramo & "' and Codsecc = '" & Seccion & "'"
	StrSql = StrSql & " and ano = '" & ano & "' and Periodo = '" & Periodo & "' " 

	'response.write(strsql)
	'response.End()
	
	Set Rst = Conn.Execute(StrSql)
	
	Confirma = 0   
	Do while not Rst.Eof
  	  	   
	    StrSql = "Update ra_carga set prioridadsol = '"& PrioridadSol &"'  "
		StrSql = StrSql & " where Codcli = '" & Codcli & "' "
		StrSql = StrSql & " and CodRamo = '" & Rst("CodRamo") & "' and Codsecc = '"& Rst("Codsecc")&"'"
		StrSql = StrSql & " and ano = '" & ano & "' and Periodo = '" & Periodo & "' "
		
        'response.Write(strsql)
		'response.End()
		Conn.Execute StrSql
		
	    StrSql = "Update ra_carga_log set prioridadsol = '"& PrioridadSol &"'  "
		StrSql = StrSql & " where Codcli = '" & Codcli & "' "
		StrSql = StrSql & " and CodRamo = '" & Rst("CodRamo") & "' and Codsecc = '"& Rst("Codsecc")&"'"
		StrSql = StrSql & " and ano = '" & ano & "' and Periodo = '" & Periodo & "' "
		
        'response.Write(strsql)
		'response.End()
		Conn.Execute StrSql		
		
		'StrSql = "Delete from sis_reg_inscripcion "
		'StrSql = StrSql & " where CodCli = '" & CodCli & "' "
		
		Conn.Execute StrSql
	   
	    Confirma = 1
		Rst.movenext
	Loop 
	
	rst.close   

	'Confirma = 0
%>
<body bgcolor="#FFFFFF" text="#000000">


<%if Confirma = 1 then%>
<script>

    //window.top.mainFrame.location.href="solic-resultado.asp"
	parent.top.location = "mod-resultado.asp"
    //parent.location = "resultado.htm"
    //parent.parent.frames("leftFrame").location = "f-izq.asp";

</script>
<%else%>
<script>
  //alert("Esta asignatura ya no tiene cupo, debes revisar y modificar tu inscripcion rte de Seccion en el Ramo <%= session("RamoSinCupo")%>  ya que la vacante, fuè Confirmado por otro alumno ....!!!! ")
  //window.top.mainFrame.location.href="inscrip-asigna.htm"
 parent.location = "modifica-toma-de-ramos.asp"
  //window.top.mainFrame.location.href="asignatura-seccion.asp"
  //parent.parent.frames("leftFrame").location = "f-izq.asp";
</script>
<%end if%>
</body>
</html>
