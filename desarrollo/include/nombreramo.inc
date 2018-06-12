<%
Function GetNombreRamo(CodRamo)
  Dim StrSql
  Dim Rst

  StrSql = "Select Nombre , credito from ra_ramo with (nolock) where codramo = '" & trim(CodRamo) & "' "

  Set Rst = Conn.Execute(StrSql)

  'Rst.open StrSql 

  'Response.write(StrSql)
  'Response.End
  if not rst.eof then
     GetNombreRamo = Rst("Nombre") & "(" & Rst("Credito") & ")"
  else
     GetNombreRamo = ""
  end if
  Rst.close()

End function
%>