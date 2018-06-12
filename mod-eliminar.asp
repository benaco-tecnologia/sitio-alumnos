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
Codcli = Session("Codcli")
Ano = Session("Ano")
Periodo = Session("Periodo")

	StrSql = "Select CodRamo, CodSecc, RamoEquiv from ra_carga_log with (nolock)"
	StrSql = StrSql & " Where Codcli = '" & Codcli & "'  "
	StrSql = StrSql & " and codramo = '" & Codramo & "' "
	StrSql = StrSql & " And ano = " & ano & " and Periodo = " & Periodo 
	
	'response.write(strsql)
	'response.End()
	Set Rst = Conn.Execute(StrSql)
	
	Confirma = 0   
	Do while not Rst.Eof
	   
		StrSql = "Delete from ra_carga_log "
		StrSql = StrSql & " where CodCli = '" & CodCli & "' "
		StrSql = StrSql & " and codramo = '" & rst("Codramo")& "' "
		StrSql = StrSql & " and ano = '" & ano & "' and Periodo = '" & Periodo & "' "
  	   
	    'response.write(strsql)
		'response.End()
	    Conn.Execute StrSql
	   
	    StrSql = "Update ra_carga set RamoEquiv = null, RAMOEQUIV_I = null, Codsecc= null, CODSECC_I= null, inscrito = null,  "
		StrSql = StrSql & " preinscrito = null, nombreventana = null, INSCRITOXVENTANA = null, estadosol = null, prioridadsol = null"
		StrSql = StrSql & " where Codcli = '" & Codcli & "' "
		StrSql = StrSql & " and CodRamo = '" & Rst("CodRamo") & "' "
		'StrSql = StrSql & " and Codsecc = '"& Rst("Codsecc")&"' "
		StrSql = StrSql & " and ano = '" & ano & "' and Periodo = '" & Periodo & "' "
		
        'response.Write(strsql)
		'response.End()
		Conn.Execute StrSql
		
		'StrSql = "Delete from sis_reg_inscripcion "
		'StrSql = StrSql & " where CodCli = '" & CodCli & "' "
		
		StrSql = "Update ra_cargaactividad set inscrito = null, PreInscrito = null, ramoequiv = null, ramoequiv_i = null, codsecc = null, "
		Strsql = Strsql & " codsecc_i = null, EstadoSol = null , inscritoxventana = null , nombreventana = null"
		StrSql = StrSql & " where Codcli = '" & Codcli & "' and codramo =  '"& Codramo &"' "
		StrSql = StrSql & " and ano = '" & ano & "' and Periodo = '" & Periodo & "' "
	
		'Conn.Execute StrSql
	    'response.Write(strsql)
		 'response.End()
	    Confirma = 1
		Rst.movenext
	Loop 
	
	rst.close   
    'response.End()
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
