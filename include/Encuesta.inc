<%
function GetEncuestado()
Dim StrSql
Dim Rst

  StrSql = "Select * from ra_encuestados where  "
  StrSql = StrSql & " codcli = '" & Session("codcli") & "' "
  StrSql = StrSql & " and ano= " & Session("ano") & ""
  StrSql = StrSql & " and periodo= " & Session("periodo") & ""
  StrSql = StrSql & " and coalesce(evaluado,'')<>'S'"
	
  'Response.write(StrSql)
  'Response.end
  if BCL_ADO(StrSql, Rst) then 
    'Existen ramos sin evaluar 
    SESSION("ENCUETADO") = "N"
  else
    'no tienen ramos sin evaluar
    SESSION("ENCUETADO") = "S"
  end if
 
end function
%>