<%
function GetPeriodoActivo()
Dim StrSql
Dim Rst

  StrSql = "Select * from sis_inscripcion where getdate() between fecha_inicio and fecha_final "
  StrSql = StrSql & " and CodCarr = '" & Session("codCarr") & "' "

  'Response.write(StrSql)
  'Response.end
  if BCL_ADO(StrSql, Rst) then  
    SESSION("PERIODO_INS") = Rst("Nombre")
    SESSION("PER_ID") = Rst("ID")
    SESSION("FECHA_INI_INS") = Valnulo(Rst("fecha_inicio"), DAT_)
    SESSION("FECHA_FIN_INS") = Valnulo(Rst("fecha_final"), DAT_)
  else
    ' no hay periodo
    SESSION("PERIODO_INS") = ""  
    SESSION("PER_ID") = "-1"
    'Response.write("Per_id = -1 ")
  end if
 
end function

function InscritoNormal()
Dim StrSql
Dim Rst

  if trim(Session("CodCli")) <> "" then
    StrSql = "Select 1 from sis_reg_inscripcion a, sis_inscripcion b "
    StrSql = StrSql & " Where a.fecha between b.fecha_inicio and b.fecha_final "
    StrSql = StrSql & "  and b.Id = " & SESSION("PER_ID")
    StrSql = StrSql & "  and a.codcli = '" & SESSION("CODCLI") & "' "
  
    InscritoNormal = BCL_ADO(StrSql, Rst) 
  else
    InscritoNormal = true
  end if
  'Response.write(StrSql)
  'Response.End 
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