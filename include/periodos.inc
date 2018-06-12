<%
function GetPeriodoActivo()
Dim StrSql
Dim Rst
   
  StrSql = "SP_GET_PERIODOACTIVO '" & Session("codcli") & "' "
  
  if BCL_ADO(StrSql, Rst) then  
    SESSION("PERIODO_INS") = Rst("PERIODO_INS")
	SESSION("PER_ID") = Rst("PER_ID")
    SESSION("FECHA_INI_INS") = Valnulo(Rst("FECHA_INI_INS"), DAT_)
    SESSION("FECHA_FIN_INS") = Valnulo(Rst("FECHA_FIN_INS"), DAT_)
	SESSION("ID_TR") = Rst("ID_TR")
  end if 
	
end function

function InscritoNormal()
Dim StrSql
Dim Rst



Sql = "select COALESCE(dbo.Fn_ValorParame('Ano'),0) as Ano,COALESCE(dbo.Fn_ValorParame('periodo'),0) as periodo " 
If bcl_ado(Sql, rst) Then
	Ano = rst("Ano")
	Periodo = rst("periodo")
End If

  if trim(Session("CodCli")) <> "" then
    StrSql = "Select TOP 1 1 from sis_reg_inscripcion a, sis_inscripcion b,mt_alumno al "
    StrSql = StrSql & " Where a.fecha between b.fecha_inicio and b.fecha_final "
    StrSql = StrSql & "  and al.codcli=a.codcli and al.codcarpr=b.codcarr"
    StrSql = StrSql & "  and a.codcli = '" & SESSION("CODCLI") & "' "
	StrSql = StrSql & "  and b.ano=" & Ano & " and b.periodo=" & Periodo & " "
  
    InscritoNormal = BCL_ADO(StrSql, Rst) 
	

	if InscritoNormal = false then    'NUEVO PREGUNTAR X GLOBAL AL IGUAL Q EL CRITERIO DE TOMA DE RAMOS
	
		StrSqlG = "Select TOP 1 1 from sis_reg_inscripcion a, ra_calendario_tr b,mt_alumno al "
		StrSqlG = StrSqlG & "Where a.fecha between b.fecha_ini and b.fecha_fin "
		StrSqlG = StrSqlG & "and al.codcli=a.codcli " 
		StrSqlG = StrSqlG & "and a.codcli = '" & SESSION("CODCLI") & "' "
		StrSqlG = StrSqlG & "and b.ano=" & Ano & " and b.periodo=" & Periodo & " "
		
		InscritoNormal = BCL_ADO(StrSqlG, RstG) 		
	end if 	
	
  else
    InscritoNormal = true
  end if
    
end function

function InscritoEspecial()
Dim StrSql
Dim Rst

  StrSql = "Select 1 from sis_reg_solicitud a, sis_inscripcion b "
  StrSql = StrSql & " Where a.fecha between b.fecha_inicio and b.fecha_final "
  StrSql = StrSql & "  and b.Id = " & SESSION("PER_ID")
  StrSql = StrSql & "  and a.codcli = '" & SESSION("CODCLI") & "' "
  InscritoEspecial = BCL_ADO(StrSql, Rst) 

end function

%>