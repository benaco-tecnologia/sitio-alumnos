<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="include/mail.inc" -->
<!--#INCLUDE FILE="include/audita.inc" -->
<!--#INCLUDE FILE="include/Cupos.inc" -->
<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%
dim CodCli,CodSede,CodPestud,Codcarr,ano,Periodo, Confirma

CodCli = Session("CodCli")
CodSede = Session("CodSede")
CodPestud = Session("CodPestud")
Codcarr=session("Codcarr")
ano=session("ano")
Periodo=session("Periodo")

 'Response.Write("Por favor, disculpe las molestias...Estamos trabajando para usted...Bettersoft...")
 'Response.End


 if ValidaCuposPreinscritos(Codcli,Ano,Periodo,Codpestud,Codcarr,Codsede)=1 and ValidaCuposPreinscritosActividad(Codcli,Ano,Periodo,Codpestud,Codcarr,Codsede)=1 and ValidaInscripcionLPT(Codcli,Ano,Periodo,Codpestud,Codcarr,Codsede)=1  and ValidaInscripcionTeorico(Codcli,Ano,Periodo,Codpestud,Codcarr,Codsede)=1 then
     Confirma=1
	  
	  'response.Write("Precaucion.. confirmacion ......")
	  'response.end 
	  
	   Codcli = Session("Codcli")
	   Ano = Session("Ano")
	   Periodo = Session("Periodo")
	
	   StrSql = "Update ra_carga set inscrito = PreInscrito , ramoequiv = ramoequiv_i, codsecc = codsecc_i "
	   StrSql = StrSql & " Where Codcli = '" & Trim(Codcli) & "' "
	   StrSql = StrSql & " And ano = " & ano & " and Periodo = " & Periodo 
		
	   'response.Write(strsql)
	   'response.end 	
	   Session("Conn").execute StrSql

	   StrSql = "Update ra_cargaactividad set inscrito = PreInscrito , ramoequiv = ramoequiv_i, codsecc = codsecc_i "
	   StrSql = StrSql & " Where Codcli = '" & Trim(Codcli) & "' "
	   StrSql = StrSql & " And ano = " & ano & " and Periodo = " & Periodo 
	
	   Session("Conn").execute StrSql
	   
	   StrSql = "Delete from ra_nota where CodCli = '" & CodCli & "' "
	   StrSql = StrSql & " and Ano = " & Ano & " And Periodo = '" & Periodo & "' "
	   StrSql = StrSql & " And ltrim(rtrim(Estado)) = '' "
	   Session("Conn").Execute StrSql

	   StrSql = "Delete from ra_notaactividad where CodCli = '" & CodCli & "' "
	   StrSql = StrSql & " and Ano = " & Ano & " And Periodo = '" & Periodo & "' "
	   StrSql = StrSql & " And ltrim(rtrim(Estado)) = '' "
	   Session("Conn").Execute StrSql
	   
	   ' Inscribiremos en lista definitiva
	   StrSql = "Select CodRamo, CodSecc, RamoEquiv from ra_carga with (nolock)"
	   StrSql = StrSql & " Where inscrito = 'S'"
	   StrSql = StrSql & " And Codcli = '" & Codcli & "' "
	   StrSql = StrSql & " And ano = " & ano & " and Periodo = " & Periodo 
	
	   Set Rst = Session("Conn").Execute(StrSql)
	   
	   Do while not Rst.Eof
		  StrSql = "Select CodCli from ra_nota "
		  StrSql = StrSql & " Where Codcli = '" & Codcli & "' "
		  StrSql = StrSql & " And CodRamo = '" & Rst("CodRamo") & "' "
		  StrSql = StrSql & " And ano = " & ano & " and Periodo = " & Periodo 
		  
		  If BCL_ADO(StrSql, Rst2) then
			 StrSql = " Update ra_nota set RamoEquiv = '" & Rst("RamoEquiv") & "' , CodSecc = '" & Rst("CodSecc") & "' "
			 StrSql = StrSql & " Where Codcli = '" & Codcli & "' "
			 StrSql = StrSql & " And CodRamo = '" & Rst("CodRamo") & "' "
			 StrSql = StrSql & " And ano = " & ano & " and Periodo = " & Periodo 
	  
			 Session("Conn").Execute StrSql 
		  else
			 StrSql = "Insert into Ra_nota (CodCli, CodRamo, CodSecc, Ano, Periodo, Estado, RamoEquiv) Values ('"
			 StrSql = StrSql & Codcli & "','" & Rst("CodRamo") & "', "
			 StrSql = StrSql & " '" & Rst("CodSecc") & "', "
			 StrSql = StrSql & " " & Ano & ", "
			 StrSql = StrSql & " '" & Periodo & "', "
			 StrSql = StrSql & "'', '"& Rst("RamoEquiv") &"')"
			 
			 'response.Write(StrSql)
			 'response.End()
			 Session("Conn").Execute StrSql
			 
		  end if
	
		  RegAuditoria Codcli, "Inscripcion Normal", "Inscripcion:" + Rst("CodRamo") + " " + Rst("RamoEquiv")
		  
		  Rst.movenext
	   Loop
	   
	   rst.close	
	   rst2.close
	   	
	   ' Inscribiremos en lista definitiva de actividad
	   StrSql = "Select CodRamo, CodSecc, RamoEquiv from ra_cargaactividad with (nolock)"
	   StrSql = StrSql & " Where inscrito = 'S'"
	   StrSql = StrSql & " And Codcli = '" & Codcli & "' "
	   StrSql = StrSql & " And ano = " & ano & " and Periodo = " & Periodo 
	
	   Set Rst = Session("Conn").Execute(StrSql)
	   
	   Do while not Rst.Eof

		  StrSql = "Select CodCli from ra_notaactividad "
		  StrSql = StrSql & " Where Codcli = '" & Codcli & "' "
		  StrSql = StrSql & " And CodRamo = '" & Rst("CodRamo") & "' "
		  StrSql = StrSql & " And ano = " & ano & " and Periodo = " & Periodo & " and codsecc = " & Rst("CodSecc") & ""
		  
		  If BCL_ADO(StrSql, Rst2) then
			 StrSql = " Update ra_notaactividad set RamoEquiv = '" & Rst("RamoEquiv") & "' , CodSecc = '" & Rst("CodSecc") & "' "
			 StrSql = StrSql & " Where Codcli = '" & Codcli & "' "
			 StrSql = StrSql & " And CodRamo = '" & Rst("CodRamo") & "' "
			 StrSql = StrSql & " And ano = " & ano & " and Periodo = " & Periodo 
			 strsql = strsql & " and codsecc = " & Rst("CodSecc") & ""
			 Session("Conn").Execute StrSql 
		  else
			 StrSql = "Insert into Ra_notaactividad (CodCli, CodRamo, CodSecc, Ano, Periodo, Estado, RamoEquiv) Values ('"
			 StrSql = StrSql & Codcli & "','" & Rst("CodRamo") & "', "
			 StrSql = StrSql & " '" & Rst("CodSecc") & "', "
			 StrSql = StrSql & " " & Ano & ", "
			 StrSql = StrSql & " '" & Periodo & "', "
			 StrSql = StrSql & " '', '" & Rst("RamoEquiv") & "' )"
			 
			 Session("Conn").Execute StrSql
		  end if
		  RegAuditoria Codcli, "Inscripcion Normal Actividad", "Inscripcion:" + Rst("CodRamo") + " " + Rst("RamoEquiv")
		  Rst.movenext
	   Loop
	
		' Marcaremos la inscripcion
	   StrSql = "Insert into sis_reg_inscripcion (Codcli, fecha) values ('" & Codcli & "', getdate()) "
	   Session("Conn").execute StrSql
	
	   'SendMail Session("RutAlum")
			  
	   'Conn.close()
 else
' 	redireccionar a la Pagina anterior...
	Confirma=0
 end if
%>
<body bgcolor="#FFFFFF" text="#000000">


<%if Confirma = 1 then%>
<script>


	window.top.location.href = "resultado.asp"
    //parent.location = "resultado.htm"
    //parent.parent.frames("leftFrame").location = "f-izq.asp";

</script>
<%else%>
<script>
  //alert("Esta asignatura ya no tiene cupo, debes revisar y modificar tu inscripcion rte de Seccion en el Ramo <%= session("RamoSinCupo")%>  ya que la vacante, fue Confirmado por otro alumno ....!!!! ")
  //window.top.mainFrame.location.href="inscrip-asigna.htm"
// parent.location = "frame-inscrip-asigna.htm"
  window.top.location.href='frame-inscrip-asigna.asp';
  //window.top.mainFrame.location.href="asignatura-seccion.asp"
  //parent.parent.frames("leftFrame").location = "f-izq.asp";
</script>
<%end if%>
</body>
</html>
