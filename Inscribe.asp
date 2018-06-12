<%
  Response.Expires = -1
%>
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->

<%

Dim StrSql
Dim Rst
DIM RST1
  codcarr=session("codcarr") 

  CodCli = Session("CodCli")
  CodSede = Session("CodSede")
  Ano = Session("Ano")
  Periodo = Session("Periodo")


  Curso = Trim(Request("C"))
  Ramo = Trim(Request("R"))
  Seccion = Trim(Request("S"))
  TipoCurso = Request("H")
  CodSeccTeo= trim(request("D"))
  session("RamoSeleccionado")=Ramo
  session("SeccionSeleccionado")=Seccion
  'response.Write(tipocurso)
  'response.End()

 if tipocurso<>"T" then
	strsql="select codcli from ra_carga"
	StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
	StrSql = StrSql & "  And codramo = '" & Curso & "' "
	StrSql = StrSql & "  And Ano = " & Ano
	StrSql = StrSql & "  And Periodo = " & Periodo
	StrSql = StrSql & "  and codsecc_i = " & CodSeccTeo 
	StrSql = StrSql & "  And ramoequiv_i = '" & Ramo & "' " 
	
	if not bcl_ado(strsql, rst) then
		if (clng(Seccion) = clng(0)) then
			strsql = " delete from ra_cargaactividad"
			StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
			StrSql = StrSql & "  And codramo = '" & Curso & "' "
			StrSql = StrSql & "  And Ano = " & Ano
			StrSql = StrSql & "  And Periodo = " & Periodo
			strsql = strsql & "  And TipoCurso = '"& TipoCurso &"'"
			
			Session("Conn").execute StrSql				
			
		end if
		 Response.Redirect("Horario.asp?P=S&T=S&X=1")			   
	else
		if (clng(Seccion) = clng(0)) then				
			strsql = " delete from ra_cargaactividad"
			StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
			StrSql = StrSql & "  And codramo = '" & Curso & "' "
			StrSql = StrSql & "  And Ano = " & Ano
			StrSql = StrSql & "  And Periodo = " & Periodo
			strsql = strsql & "  And TipoCurso = '"& TipoCurso &"'"	
	
			Session("Conn").execute StrSql		
		else

			If VerificaTope(CodCli, Curso, Ramo, Seccion, CodSede, Ano, Periodo) then
			   Response.Redirect("Horario.asp?P=S&T=S&M=S&R=" & Ramo&"&S=" & Seccion & "&C=" & Curso)
			end if			
			strsql = " select codramo from ra_cargaactividad"
			StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
			StrSql = StrSql & "  And codramo = '" & Curso & "' "
			StrSql = StrSql & "  And Ano = " & Ano
			StrSql = StrSql & "  And Periodo = " & Periodo
			strsql = strsql & "  And TipoCurso = '"& TipoCurso &"'"							
			if not bcl_ado(strsql, rst) then
				StrSql = "insert into ra_cargaactividad (codramo,codcli,ano,periodo,reprobado,codcarr,inscrito,estado,ramoequiv,prioridad,ramoequiv_i,codsecc_i,PreInscrito,Tipocurso)"	
				strsql = strsql & " values ('"& curso &"','"& Codcli &"', '"& ano &"','"& periodo &"','N'" 
				strsql = strsql & " ,'"& codcarr &"', 'N'"
				strsql = strsql & " ,'N','"& Ramo &"', '1'"
				strsql = strsql & " ,'"& Ramo &"','"& SECCION &"','S','" & Trim(TipoCurso) & "')"
				Session("Conn").execute StrSql		
			else
				StrSql = "Update ra_cargaactividad Set PreInscrito = 'S', Ramoequiv_i = '" & Ramo & "', CodSecc_i = '" & Seccion & "' "
				StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
				StrSql = StrSql & " And CodRamo = '" & Curso & "' "
				StrSql = StrSql & " And Ano = " & Ano
				StrSql = StrSql & " And Periodo = " & Periodo
				strsql = strsql & " And TipoCurso = '"& TipoCurso &"'"
				Session("Conn").execute StrSql
			end if

			StrSql = "Update ra_cargaactividad Set PreInscrito = 'N', ramoequiv_i = null, codsecc_i = null "
			StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
			StrSql = StrSql & "  And CodRamo <> '" & Curso & "' "
			StrSql = StrSql & "  And RamoEquiv_i = '" & Ramo & "' "
			StrSql = StrSql & "  And Ano = " & Ano
			StrSql = StrSql & "  And Periodo = " & Periodo	
			strsql = strsql & "  And TipoCurso = '"& TipoCurso &"'"
			Session("Conn").execute StrSql
		end if
	end if
  else
 
		if (clng(Seccion) = clng(0)) then
			StrSql = "Update ra_carga Set PreInscrito = 'N',ramoequiv_i = null, codsecc_i = null "', SolicitudEspecial = 'N'"	
			StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
			StrSql = StrSql & "  And codramo = '" & Curso & "' "
			StrSql = StrSql & "  And Ano = " & Ano
			StrSql = StrSql & "  And Periodo = " & Periodo

			Session("Conn").execute StrSql
			  
			'nuevo
			strsql = " delete from ra_cargaactividad"
			StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
			StrSql = StrSql & "  And codramo = '" & Curso & "' "
			StrSql = StrSql & "  And Ano = " & Ano
			StrSql = StrSql & "  And Periodo = " & Periodo
			Session("Conn").execute StrSql
			
		else
			'Mejora validacion combos
			strCombo = "SP_VALIDA_COMBOS_TR_PA '" & session("codcli")&"',"& Ano &"," & Periodo & ",'" & CodSede & "','" & codcarr &"','"& Ramo &"',"& Seccion &""
			'response.Write(strCombo)
			'response.End()
			if bcl_ado(strCombo, rstCombo) then
				if rstCombo("resultado") <> "OK" then
					nombrecombo = rstCombo("nombrecombo")
					Response.Redirect("Horario.asp?P=S&T=S&M=S&NC="& nombrecombo &"&R=" & Ramo&"&S=" & Seccion & "&C=" & Curso)
				end if 
			end if 
			
		
			If VerificaTope(CodCli, Curso, Ramo, Seccion, CodSede, Ano, Periodo) then
			   Response.Redirect("Horario.asp?P=S&T=S&M=S&R=" & Ramo&"&S=" & Seccion & "&C=" & Curso)
			end if
			
			StrSql = "Update ra_carga Set PreInscrito = 'S', Ramoequiv_i = '" & Ramo & "', CodSecc_i = '" & Seccion & "' " ', SolicitudEspecial = 'S' "
			StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
			StrSql = StrSql & " And CodRamo = '" & Curso & "' "
			StrSql = StrSql & " And Ano = " & Ano
			StrSql = StrSql & " And Periodo = " & Periodo
			Session("Conn").execute StrSql
	
			StrSql = "Update ra_carga Set PreInscrito = 'N', ramoequiv_i = null, codsecc_i = null "
			StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
			StrSql = StrSql & "  And CodRamo <> '" & Curso & "' "
			StrSql = StrSql & "  And RamoEquiv_i = '" & Ramo & "' "
			StrSql = StrSql & "  And Ano = " & Ano
			StrSql = StrSql & "  And Periodo = " & Periodo	
			Session("Conn").execute StrSql

		end if
		
 end if



Response.Redirect("Horario.asp?P=S&T=S")
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>
 <script languaje="javascript">


function mensaje_alert1()
{
   alert("hola")
   return;
window.location = "Horario.asp?P=S&T=S";
}
 </script>
<P>&nbsp;</P>

</BODY>
</HTML>
<!--#INCLUDE file="include/desconexion.inc" -->
