<%
  Response.Expires = -1
%>
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->

<%
CodCli = Session("CodCli")
CodSede = Session("codsede")
CodCarr = Session("codcarr")
Ano = Session("Ano")
Periodo = Session("Periodo")

Combo = Trim(Request("C"))
Inscribe = Trim(Request("I"))

session("ComboTR") = Combo

strParame="SELECT coalesce(dbo.Fn_ValorParame('FILTRASECCIONINICIALTRPA'),'')Parame"
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
	FILTRASECCIONINICIALTRPA=rstParame("Parame")
end if

StrSqlc = "select filtrajornada from mt_carrer "
StrSqlc = StrSqlc & " where codcarr='" & trim(CodCarr) & "'"
StrSqlc = StrSqlc & " and filtrajornada = 'S'"
if BCL_ADO(StrSqlc, rstc) then
	FiltraJornada ="S"
else
	FiltraJornada = "N"
end if

if Inscribe = "S" then

	str ="select codramo from ra_Carga where codcli ='" & CodCli & "' and ano = "& Ano &" and periodo =" & Periodo &""
	if bcl_ado(str, rst) then
		While Not rst.Eof
			Ramo = rst("CODRAMO")
			
			if EsPlanificadoXJornada(Ramo, CodSede, Ano, Periodo, CodCarr,FiltraJornada) then
			   StrSeccion = "Select e.CodRamo, e.CodSecc "
			   StrSeccion = StrSeccion & " from ra_seccio e, ra_profes p, ra_ramo r "
			   StrSeccion = StrSeccion & " Where e.CodRamo = '" & Ramo & "' and e.CodSede = '" & CodSede & "' " 
			   StrSeccion = StrSeccion & " And e.Ano = " & ano & " And e.Periodo = " & Periodo & ""
			   StrSeccion = StrSeccion & " and e.CodProf *= p.CodProf "
			   StrSeccion = StrSeccion & " and e.CodCarr='" & trim(Codcarr) & "'" 
			   If FiltraJornada = "S" Then
				   StrSeccion = StrSeccion & " And (e.Jornada = '" & Session("Jornada") & "' OR e.Jornada2 = '" & Session("Jornada") & "')"
			   end if                  
			   StrSeccion = StrSeccion & " and e.codramo = r.codramo" 			   
			   if FILTRASECCIONINICIALTRPA ="SI" then
					StrSeccion = StrSeccion & " and e.codsecc = dbo.FN_CODSECC_INICIAL_ALUMNO('"& session("codcli") &"') "
			   end if			
			else
			   Equivalente = GetEquivalente(Ramo, CodSede, Ano, Periodo, CodCarr)
			   if Trim(Equivalente) <> "" then
					   StrSeccion = "Select e.CodRamo, e.CodSecc "
					   StrSeccion = StrSeccion & " from ra_seccio e, ra_profes p, ra_ramo r "					  
					   StrSeccion = StrSeccion & " Where e.CodRamo in (Select a.ramoequiv from ra_equiv a with (nolock), mt_carrer b with (nolock) Where a.codcarr = b.codcarr And b.Sede = '" & CodSede & "' And a.CodRamo = '" & Ramo & "' and a.codcarr = '" & trim(CodCarr) & "')" 					    
					   StrSeccion = StrSeccion & " and e.CodSede = '" & CodSede & "' " 
					   StrSeccion = StrSeccion & " And e.Ano = " & ano & " And e.Periodo = " & Periodo
					   StrSeccion = StrSeccion & " and e.CodProf *= p.CodProf "
					   If FiltraJornada = "S" Then
						   StrSeccion = StrSeccion & " And (e.Jornada = '" & Session("Jornada") & "' or e.Jornada2 = '" & Session("Jornada") & "')"
					   end if                  
					   StrSeccion = StrSeccion & " and e.codramo = r.codramo" 
					   if FILTRASECCIONINICIALTRPA ="SI" then
							StrSeccion = StrSeccion & " and e.codsecc = dbo.FN_CODSECC_INICIAL_ALUMNO('"& session("codcli") &"') "
					   end if			
			   else
					   StrSeccion = "Select distinct a.codramo,a.codsecc " 
					   StrSeccion = StrSeccion & " from ra_optivo a,ra_ramo b,ra_seccio e, ra_profes p" 
					   StrSeccion = StrSeccion & " where a.codcarr  = '" & CodCarr & "' " 
					   StrSeccion = StrSeccion & " and a.codramo=b.codramo" 
					   StrSeccion = StrSeccion & " and a.ano=" & Ano & "" 
					   StrSeccion = StrSeccion & " and a.periodo=" & Periodo & "" 
					   StrSeccion = StrSeccion & " and a.codramo=e.codramo" 
					   StrSeccion = StrSeccion & " and a.codsecc=e.codsecc" 
					   If FiltraJornada = "S" Then
						   StrSeccion = StrSeccion & " And (e.Jornada = '" & Session("Jornada") & "' or e.Jornada2 = '" & Session("Jornada") & "')"
					   end if                  
					   StrSeccion = StrSeccion & " and a.ano=e.ano" 
					   StrSeccion = StrSeccion & " and a.periodo=e.periodo" 
					   StrSeccion = StrSeccion & " and e.CodSede = '" & CodSede & "' "
					   StrSeccion = StrSeccion & " and a.CodSede = e.CodSede"
					   StrSeccion = StrSeccion & " And a.tipo <> 'OPTATIVO ESPECIAL'"				
					   StrSeccion = StrSeccion & " and a.codopt = '" & Ramo & "' "
					   StrSeccion = StrSeccion & " and e.CodProf *= p.CodProf "					   
					   if FILTRASECCIONINICIALTRPA ="SI" then
							StrSeccion = StrSeccion & " and e.codsecc = dbo.FN_CODSECC_INICIAL_ALUMNO('"& session("codcli") &"') "
					   end if					   
				end if
			end if 
			StrSeccion = StrSeccion & " intersect  SELECT R.CODRAMO,D.CODSECC "
			StrSeccion = StrSeccion & " FROM RA_COMBO_INS_DET  D INNER JOIN RA_RAMO R ON D.CODRAMO = R.CODRAMO "
			StrSeccion = StrSeccion & " WHERE Id_combo = "& Combo &" "
						
			if BCL_ADO(StrSeccion, RstSeccion) then
				While Not RstSeccion.Eof
					
					StrSql = "Update ra_carga Set PreInscrito = 'S', Ramoequiv_i = '" & RstSeccion("codramo") & "', CodSecc_i = '" & RstSeccion("codsecc") & "' " 
					StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
					StrSql = StrSql & " And CodRamo = '" & Ramo & "' "
					StrSql = StrSql & " And Ano = " & Ano
					StrSql = StrSql & " And Periodo = " & Periodo
					Session("Conn").execute StrSql
					
					StrSql = "Update ra_carga Set PreInscrito = 'N', ramoequiv_i = null, codsecc_i = null "
					StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
					StrSql = StrSql & "  And CodRamo <> '" & Ramo & "' "
					StrSql = StrSql & "  And RamoEquiv_i = '" & Ramo & "' "
					StrSql = StrSql & "  And Ano = " & Ano
					StrSql = StrSql & "  And Periodo = " & Periodo	
					Session("Conn").execute StrSql
				
					RstSeccion.MoveNext
				Wend
			end if 
			
			rst.MoveNext
		Wend
	end if
else
	StrSql = "Update ra_carga Set PreInscrito = 'N', Ramoequiv_i = NULL, CodSecc_i = NULL " 
	StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
	StrSql = StrSql & " And Ano = " & Ano
	StrSql = StrSql & " And Periodo = " & Periodo
	Session("Conn").execute StrSql
	
	session("ComboTR") = ""
end if

Response.Redirect("Horario.asp?P=S&T=S&RC=S")
%>

<!--#INCLUDE file="include/desconexion.inc" -->
