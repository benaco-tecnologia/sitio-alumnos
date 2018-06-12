<%

Function PreRequisito(CodRamo, CodMalla)
Dim StrSql, StrResult
Dim Rst, Est
'******12:33 10/01/2007****************
   StrSql = "Select Requisito from ra_secuen WITH (NOLOCK) "
   StrSql = StrSql & " Where CodRamo = '" & trim(CodRamo) & "' "
   StrSql = StrSql & " And CodPestud = '" & trim(CodMalla) & "' "
   
   Set Rst = Session("Conn").execute(StrSql)
   StrResult = "Prerequisitos : \n"
   Do While not Rst.Eof
     'response.write(Rst("Requisito"))
     StrResult = StrResult & "    " & GetnombreRamo(Rst("Requisito")) & "\n"
     Rst.movenext
   loop
   PreRequisito = StrResult
'*************************************  
'  Set Com2 = Server.CreateObject("ADODB.Command")
'  Set IPar1= Server.CreateObject("ADODB.Parameter")
'  Set IPar2= Server.CreateObject("ADODB.Parameter")
'  Set OPar1= Server.CreateObject("ADODB.Parameter")
'  With Com2
'      .ActiveConnection = Conn
'      .CommandText = "Pr_Ramo_Prerreq"
'      .CommandType = 4
'      .Parameters.Refresh
'      .CommandTimeout = 0
'  End With
'  Set IPar1 = Com2.CreateParameter("Codcli", 200, 1, 30, CodCli)
'  Com2.Parameters(1).Value = IPar1.Value
'  Set IPar2 = Com2.CreateParameter("Ramo", 200, 1, 30, CodRamo)
'  Com2.Parameters(2).Value = IPar2.Value
'  Set OPar1 = Com2.CreateParameter("Estado", 200, 2, 3, Est)
                        
'  Com2.Execute
'  PreRequisito = Com2.Parameters(3).Value

End Function
%>
