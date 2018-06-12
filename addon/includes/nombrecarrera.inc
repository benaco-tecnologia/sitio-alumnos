<%
Function GetNombreCarrera(CodCarr)
  Dim StrSql
  Dim Rst

  StrSql = "Select Nombre_l from mt_carrer where codCarr = '" & CodCarr & "' "

  Set Rst = Session("Conn").Execute(StrSql)

  'Rst.open StrSql 

  if not Rst.eof then
    GetNombreCarrera = Rst("Nombre_l")
  else
    GetNombreCarrera = "Sin Nombre"
  end if
  Rst.close()

End function

%>