<%
Function EstadoRamo(Codcli, CodRamo) 
Dim StrSql
Dim Rst, Est

 '   StrSql = "Select estado from ra_nota Where Codcli = '" & trim(Codcli) & "' "
 '   StrSql = StrSql & " And CodRamo = '" & trim(CodRamo) & "' "
 '   StrSql = StrSql & " Order by Ano desc, periodo desc "
 '   
 '   Set Rst = Conn.execute(StrSql)
 '   if not Rst.eof then
 '     EstadoRamo = ":" & Rst("estado")
 '     Rst.close()    
 '   else
 '     EstadoRamo = ""
 '   end if

  Set Com2 = Server.CreateObject("ADODB.Command")
  Set IPar1= Server.CreateObject("ADODB.Parameter")
  Set IPar2= Server.CreateObject("ADODB.Parameter")
  Set OPar1= Server.CreateObject("ADODB.Parameter")
  With Com2
      .ActiveConnection = Session("Conn")
      .CommandText = "Pr_Ramo_Estado"
      .CommandType = 4
      .Parameters.Refresh
      .CommandTimeout = 0
  End With
  Set IPar1 = Com2.CreateParameter("Codcli", 200, 1, 30, CodCli)
  Com2.Parameters(1).Value = IPar1.Value
  Set IPar2 = Com2.CreateParameter("Ramo", 200, 1, 30, CodRamo)
  Com2.Parameters(2).Value = IPar2.Value
  Set OPar1 = Com2.CreateParameter("Estado", 200, 2, 3, Est)
                        
  Com2.Execute
  EstadoRamo = valnulo(Com2.Parameters(3).Value, STR_)

End Function
%>