<%
Function Encripta(sPasswd)
    Dim fc_b1, fc_b2, fc_b3, fc_b4, fc_b5, fc_password
    Dim fn_b1, fn_b2, fn_b3, fn_b4, fn_b5

    If Len(sPasswd) < 5 Then
      sPasswd = Space(5 - Len(sPasswd)) + sPasswd
    End If

    fc_b1 = Mid(sPasswd, 1, 1)
    fc_b2 = Mid(sPasswd, 2, 1)
    fc_b3 = Mid(sPasswd, 3, 1)
    fc_b4 = Mid(sPasswd, 4, 1)
    fc_b5 = Mid(sPasswd, 5, 1)
    fn_b1 = Asc(fc_b1)
    fn_b2 = Asc(fc_b2)
    fn_b3 = Asc(fc_b3)
    fn_b4 = Asc(fc_b4)
    fn_b5 = Asc(fc_b5)
    fn_b1 = fn_b1 + 3
    fn_b2 = fn_b2 + 5
    fn_b3 = fn_b3 + 2
    fn_b4 = fn_b4 + 4
    fn_b5 = fn_b5 + 1
    fc_b1 = Chr(fn_b1)
    fc_b2 = Chr(fn_b2)
    fc_b3 = Chr(fn_b3)
    fc_b4 = Chr(fn_b4)
    fc_b5 = Chr(fn_b5)
    Encripta = fc_b1 & fc_b4 & fc_b5 & fc_b3 & fc_b2

End Function

Function Desencripta(sPasswd)
    Dim fc_b1, fc_b2, fc_b3, fc_b4, fc_b5, fc_password
    Dim fn_b1, fn_b2, fn_b3, fn_b4, fn_b5

    fc_b1 = Mid(sPasswd, 1, 1)
    fc_b4 = Mid(sPasswd, 2, 1)
    fc_b5 = Mid(sPasswd, 3, 1)
    fc_b3 = Mid(sPasswd, 4, 1)
    fc_b2 = Mid(sPasswd, 5, 1)
    fn_b1 = Asc(fc_b1)
    fn_b2 = Asc(fc_b2)
    fn_b3 = Asc(fc_b3)
    fn_b4 = Asc(fc_b4)
    fn_b5 = Asc(fc_b5)
    fn_b1 = fn_b1 - 3
    fn_b2 = fn_b2 - 5
    fn_b3 = fn_b3 - 2
    fn_b4 = fn_b4 - 4
    fn_b5 = fn_b5 - 1
    fc_b1 = Chr(fn_b1)
    fc_b2 = Chr(fn_b2)
    fc_b3 = Chr(fn_b3)
    fc_b4 = Chr(fn_b4)
    fc_b5 = Chr(fn_b5)
    Desencripta = fc_b1 & fc_b2 & fc_b3 & fc_b4 & fc_b5
End Function

Function RutSinDV(RutIngresado)
     Dim RutAux, RutAuxSinDV, i, c
	 
	 RutAux = Mid(RutIngresado,1,Len(RutIngresado) - 1)
	 
	 For i = 1 To Len(RutAux)
	  c = Mid(RutAux,i,1)
	  If c <> "." And c <> "-" Then
	   RutAuxSinDV = RutAuxSinDV & c
	  End If 	
	 Next
	 
	 RutSinDV = RutAuxSinDV
	 
End Function

Function EliminaPalabras(Texto)
    Dim Eliminar, NuevoStr
	
	Eliminar = Array("select","drop",";","delete","insert","update","xp_")
	NuevoStr = Texto
	
	For i = 1 to UBound(Eliminar)
		NuevoStr = Replace(NuevoStr, Eliminar(i), "")
	Next
	
	EliminaPalabras = NuevoStr	
End Function	



Function GetNombreProfe(CodProf)
Dim StrSql
Dim Rst

  StrSql = "Select Ap_Pater, Ap_Mater, Nombres from ra_profes where codprof = '" & CodProf & "' "

  Set Rst = Conn.Execute(StrSql)

  'Rst.open StrSql 

  if not Rst.eof then
    GetNombreProfe = Rst("Ap_Pater") + " " + Rst("Nombres")
  else
    GetNombreProfe = "Sin Profesor"
  end if
  Rst.close()

End function

Function GetHor2(CodRamo, CodSede, CodSecc, Ano, Periodo)
  Dim StrSql
  Dim Rst

    'StrSql = " GetHorario '" & CodRamo & "','" & CodSede & "','" & CodSecc & "'," & Ano & "," & Periodo
    StrSql = "Select Horario from ra_seccio where CodRamo = '" & CodRamo & "' "
    StrSql = StrSql & " and CodSecc = '" & CodSecc & "' "
    StrSql = StrSql & " and ano = " & Ano 
    StrSql = StrSql & " and Periodo = '" & Periodo & "' "
    StrSql = StrSql & " and CodSede = '" & CodSede & "' "
    
    if BCL_ADO(StrSql, Rst) then
       GetHor2 = Rst("Horario")
    else
       GetHor2 = ""
    end if
End Function

Function GetHorario(CodRamo, CodSede, CodSecc, Ano, Periodo)
  Dim StrSql
  Dim Rst
  Dim StrHorario, lastdia

  StrSql = "Select Dia, CodMod from ra_horari where codramo = '" & CodRamo & "' "
  StrSql = StrSql & " And CodSede = '" & CodSede & "' "
  StrSql = StrSql & " And CodSecc = '" & CodSecc & "' "
  StrSql = StrSql & " And Ano = " & Ano 
  StrSql = StrSql & " And Periodo = " & Periodo
  StrSql = StrSql & " Order by case Dia when 'LUNES' then 1 when 'MARTES' then 2 when 'MIERCOLES' then 3 when 'JUEVES' then 4 when 'VIERNES' then 5 else 6 end, " & _
            " convert(numeric,CodMod) "

  'response.write(StrSql)
  'response.end
  Set Rst = Conn.Execute(StrSql)
  
  StrHorario = ""
  lastdia = ""
  lastCodMod = 0
  do while not Rst.eof
    if lastdia <> Rst("dia") then
      if trim(lastdia) <> "" then
        if FirstCodMod <> LastCodMod then
          StrHorario = StrHorario & FirstCodMod & "-" & LastCodMod & " "        
        else
          StrHorario = StrHorario & FirstCodMod & " "        
        end if
      end if
      StrHorario = StrHorario & " " & Mid(Rst("dia"), 1,2)
      lastdia = Rst("dia")
      FirstCodMod = rst("CodMod")
      lastCodMod = FirstCodMod
      'response.write("1-" & FirstCodmod & " : " & lastCodMod & " " & Rst("Dia") & "<br>")
    else
      'Response.write(lastCodmod & " v/s " & Rst("CodMod") & " " & (lastCodmod - Rst("CodMod") + 1))
      if (lastCodMod + 1 )- Rst("CodMod") = 0  then
        lastCodMod = Rst("CodMod")
        'response.write("2-" & FirstCodmod & " : " & lastCodMod & "<br>")
      else
        if FirstCodMod <> LastCodMod then
           StrHorario = StrHorario & FirstCodMod & "-" & LastCodMod & " "
        else
           StrHorario = StrHorario & FirstCodMod & " "
        end if
        FirstCodMod = Rst("CodMod")
        LastCodMod = FirstCodMod
        'response.write("3-" & FirstCodmod & " : " & lastCodMod & "<br>") 
      end if
    end if
    
    Rst.movenext
  loop
  if FirstCodMod <> LastCodMod then
     StrHorario = StrHorario & FirstCodMod & "-" & LastCodMod & " "
  else
     StrHorario = StrHorario & FirstCodMod & " "
  end if
  
  GetHorario = StrHorario  
  Rst.close()

End function

Function GetTipoRamo (CodRamo, CodPestud)
Dim StrSql
Dim Rst

   StrSql = "Select tipo from ra_curric where CodRamo = '" & CodRamo & "' And CodPestud = '" & CodPestud & "' "
   Set Rst = Conn.execute(StrSql)
   if Rst.Eof then
      GetTipoRamo = ""
   else
      GetTipoRamo = Ucase(Rst("Tipo"))
   end if   
   Rst.close()
   
End Function

Function EsPlanificado (CodRamo, CodSede, Ano, Periodo)
Dim StrSql
Dim Rst

   StrSql = "Select 1 from ra_seccio where Codramo = '" & Codramo & "' And Codsede = '" & CodSede & "' " & _
            " And Ano = " & ano & " And Periodo = " & Periodo
            
   Set Rst = Conn.execute(StrSql)
   EsPlanificado = not Rst.Eof 
   Rst.close()
End Function
  
Function GetEquivalente (CodRamo, CodSede, Ano, Periodo)
Dim StrSql, Equivalente
Dim Rst

   StrSql = "Select a.ramoequiv from ra_equiv a, mt_carrer b "
   StrSql = StrSql & " Where a.codcarr = b.codcarr "
   StrSql = StrSql & " And b.Sede = '" & CodSede & "' "
   StrSql = StrSql & " And a.CodRamo = '" & CodRamo & "' "
   StrSql = StrSql & " Order by Prioridad "
            
   Equivalente = ""         
   Set Rst = Conn.execute(StrSql)
   Do while not Rst.Eof 
      if EsPlanificado(Rst("RamoEquiv"), CodSede, Ano, Periodo) then
        GetEquivalente = Rst("RamoEquiv")
        Rst.close()
        exit Function
      else
      end if
      Rst.movenext
   loop
   Rst.close()
   GetEquivalente = Equivalente
End Function

Function EstaInscrito(Codcli, CodRamo, CodSecc, Ano, Periodo) 
Dim StrSql
Dim Rst
 
    'EstaInscrito = false
    'exit function
    
    StrSql = "Select Inscrito from ra_carga Where Codcli = '" & trim(Codcli) & "' "
    StrSql = StrSql & " And RamoEquiv = '" & trim(CodRamo) & "' "
    'StrSql = StrSql & " And CodSecc = '" & trim(CodSecc) & "' "
    StrSql = StrSql & " And ano = " & Ano 
    StrSql = StrSql & " And Periodo = " & Periodo
    
    Set Rst = Conn.execute(StrSql)
    'Response.write(Strsql)
    if not Rst.eof then
       'response.write("Inscrito = " &  Rst("Inscrito"))
       EstaInscrito = (Rst("Inscrito") = "S")
    else
       EstaInscrito = false
    end if 
    
End Function

Function EstaPreInscrito(Codcli, Curso, CodRamo, CodSecc, Ano, Periodo) 
Dim StrSql
Dim Rst

    StrSql = "Select PreInscrito, Inscrito from ra_carga Where Codcli = '" & trim(Codcli) & "' "
    'StrSql = StrSql & " And CodRamo = '" & trim(Curso) & "' "
    StrSql = StrSql & " And RamoEquiv_i = '" & trim(CodRamo) & "' "
    if trim(CodSecc) <> "" then
      StrSql = StrSql & " And CodSecc_i = " & CodSecc
    end if
    StrSql = StrSql & " And ano = '" & Ano  & "' "
    StrSql = StrSql & " And Periodo = '" & Periodo & "' "
    
    'Response.write(Strsql)
    'Response.end
    Set Rst = Conn.execute(StrSql)
    'Response.write(Strsql)
    if not Rst.eof then
       'response.write("Inscrito = " &  Rst("Inscrito"))
       EstaPreInscrito = (Rst("Inscrito") = "S") Or (Rst("PreInscrito") = "S")
    else
       EstaPreInscrito = false
    end if 
    
End Function

Function EstaCursado(Codcli, CodRamo, CodSecc) 
Dim StrSql
Dim Rst

    StrSql = "Select 1 from ra_nota Where Codcli = '" & trim(Codcli) & "' "
    StrSql = StrSql & " And RamoEquiv = '" & trim(CodRamo) & "' "
    StrSql = StrSql & " And estado in ('A', 'E','I') "
    
    Set Rst = Conn.execute(StrSql)
    if not Rst.eof then
      EstaCursado = True
      Rst.close()    
    else
      EstaCursado = false
    end if
End Function

Function ResetInscrip(CodCli, Ano, Periodo)
Dim StrSql
  StrSql = "Update ra_carga set ramoequiv_i = ramoequiv, codsecc_i = codsecc, PreInscrito = 'N' "
  StrSql = StrSql & " Where Codcli = '" & Codcli & "' "
  StrSql = StrSql & " And Ano = " & Ano
  StrSql = StrSql & " And Periodo = " & Periodo

  'response.write(StrSql)
  'response.end
  Conn.execute StrSql

  StrSql = "delete TmpTest "
  StrSql = StrSql & " Where Codcli = '" & Codcli & "' "
  StrSql = StrSql & " And Ano = " & Ano
  StrSql = StrSql & " And Periodo = " & Periodo

  Conn.execute StrSql

End Function


Function PrioridadRamo(Codcli, CodRamo, Ano, Periodo) 
Dim StrSql
Dim Rst

    StrSql = "Select Prioridad from ra_carga Where Codcli = '" & trim(Codcli) & "' "
    StrSql = StrSql & " And RamoEquiv = '" & trim(CodRamo) & "' "
    StrSql = StrSql & " And ano = " & Ano 
    StrSql = StrSql & " And Periodo = " & Periodo
    
    'Response.write(StrSql)
    Set Rst = Conn.execute(StrSql)
    if not Rst.eof then
       PrioridadRamo = Rst("Prioridad") 
    else
       PrioridadRamo = -1
    end if 
    
End Function


Function VerificaTope(CodCli, Curso, CodRamo, CodSecc, CodSede, Ano, Periodo)
Dim StrSql
Dim Rst
   StrSql = " e_TopeHorario '" & CodCli & "','" & Curso & "','" & CodRamo & "','" & CodSecc & "', '" & CodSede & "'," & Ano & "," & Periodo
   Conn.Execute StrSql
  
   'Response.write(Strsql)
   'Response.End
   StrSql = "Select 1 from ra_Topes where Codcli = '" & trim(CodCli) & "' "
   
   VerificaTope = BCL_ADO(StrSql, Rst) 
  
End Function

Function GetTopes(CodCli, Curso, CodRamo, CodSecc, CodSede, Ano, Periodo, Rst)
Dim StrSql
   StrSql = " e_TopeHorario '" & CodCli & "','" & Curso & "','" & CodRamo & "','" & CodSecc & "', '" & CodSede & "'," & Ano & "," & Periodo
   Conn.Execute StrSql
   'Response.Write(StrSql)
   'Response.End

   StrSql = "Select * from ra_Topes where Codcli = '" & CodCli & "' "
   BCL_ADO StrSql, Rst
End Function

Function GetNombreAlumno(codcli)
Dim StrSql
Dim Rst

 'Strsql="select nombre,paterno,materno from mt_client where codcli='" & codcli & "'"
 strsql="select b.nombre,b.paterno,b.materno from mt_alumno a, mt_client b "
 strsql=strsql & " where a.codcli='" & codcli & "' and b.codcli = a.rut "
 Set Rst = Conn.Execute(StrSql)

if not Rst.eof then
    GetNombrealumno = Rst("nombre") & " "& Rst("paterno") & " " &  Rst("materno")
  else
    GetNombrealumno = "Sin Nombre"
  end if
  Rst.close()

end function

Function GetDia(Dia)
Dim j
  select case DIA
      case "LUNES" : j = 1
      case "MARTES" : j = 2
      case "MIERCOLES" : j = 3
      case "JUEVES" : j = 4
      case "VIERNES" : j = 5
      case "SABADO" : j = 6
      case "DOMINGO" : j = 7
  end select
  GetDia = j
End Function


Function EstadoRamoReal(Codcli, CodRamo) 
Dim StrSql
Dim Rst

    StrSql = "Select estado from ra_nota Where Codcli = '" & trim(Codcli) & "' "
    StrSql = StrSql & " And Ramoequiv = '" & trim(CodRamo) & "' "
    StrSql = StrSql & " Order by Ano desc, periodo desc "
    
    Set Rst = Conn.execute(StrSql)
    if not Rst.eof then
      EstadoRamoReal = ":" & Rst("estado")
      Rst.close()    
    else
      EstadoRamoReal = ""
    end if
    'Response.write(StrSql)
    'response.end
End Function


Function GetCreditos(CodPestud, CredMin, CredMax)
Dim StrSql
Dim Rst

   StrSql = "Select CredMin, CredMax from ra_pestud where CodPestud = '" & CodPestud & "' "
   
   if BCL_ADO(Strsql, Rst) then
      CredMin = ValNulo(Rst("CredMin"), NUM_)
      CredMax = ValNulo(Rst("CredMax"), NUM_)
   else
      CredMin = 0
      CredMax = 0   
   end if
End Function


Function AlumnoConDeuda(rut)
Dim StrSql
Dim Rst

    StrSql = "select codcli from mt_ctadoc where codcli = '" & Trim(rut) & "' and estado = 2 "
        'str = str & " and saldo > 0 and fecven < '" & FechaServidor & "'"
    StrSql = StrSql & " and saldo > 0 and fecven < getdate() "
    'Response.write(StrSql)
    'Response.End
    AlumnoConDeuda = BCL_ADO(strsql, rst)
End Function

Function CarreraValidaDeuda(CodCarr)
Dim StrSql
Dim Rst

   StrSql = "Select v_credito from mt_carrer where CodCarr = '" & CodCarr & "' "
   if bcl_ado(StrSql, Rst) then 
      CarreraValidaDeuda = (ucase(trim(Rst("v_Credito"))) = "S" )
   else
      CarreraValidaDeuda = false
   end if
End Function

Function Normaliza(Texto)

  Normaliza = Replace(Valnulo(Texto, STR_), chr(13), "<br>")
   
End Function

Function PoseeConvalidaciones(Codcli)
Dim StrSql
Dim Rst

  StrSql = "select top 1 codcli from ra_nota where estado in ('I', 'E') and codcli = '" & Codcli & "' "
  PoseeConvalidaciones = BCL_ADO(StrSql, Rst)
End Function


Function EsAntiguo()
Dim StrSql
Dim Rst
   
   StrSql = "Select nuevo from mt_Alumno where Codcli = '" & Codcli & "' "
   if BCL_ADO(StrSql, Rst) then
      EsAntiguo = (Rst("nuevo") = "N") 
   else 
      EsAntiguo = false
   end if
End Function

Function GetNombreCarrera(CodCarr)
  Dim StrSql
  Dim Rst

  StrSql = "Select Nombre_l from mt_carrer where codCarr = '" & CodCarr & "' "

  Set Rst = Conn.Execute(StrSql)

  'Rst.open StrSql 

  if not Rst.eof then
    GetNombreCarrera = Rst("Nombre_l")
  else
    GetNombreCarrera = "Sin Nombre"
  end if
  Rst.close()
End function

Function TotalRespuesta()   
Dim Sql
Dim Rst

	'Total Pregunta Encuesta
     sql="Select count(*) from ra_preguntas"
     set rst=conn.execute(sql)
     if not rst.eof then
        TotalRespuesta=valnulo(rst(0),num_)
     else
        TotalRespuesta=0
     end if
end function

Function Encuestado(Codcli,anoed,periodoed)
Dim Sql
Dim Rst

     sql="Select * from mt_Alumno where codcli='" & Codcli & "'"
     sql=sql & " and estacad='VIGENTE' and EncDoc='S' and ano_ed='" & anoed & "' and periodo_ed='" & periodoed & "'"
	 
     set rst=conn.execute(sql)
     if bcl_ado(Sql,rst) then
        Encuestado=1
     else
        Encuestado=0
     end if
	
end Function

function TomadeRamo(codigo,codcli,anoed,periodoed)
Dim Sql
Dim Rst
TomadeRamo=True
	 
	sql="Select coalesce(RequisitoEd,'N') from mt_carrer where codcarr='" & codigo & "'"     
     set rst=conn.execute(sql)
     if bcl_ado(Sql,rst) then
        if rst(0) ="S" then		  
		   if Encuestado(codcli,anoed,periodoed)=0 then
		      TomadeRamo=0
		   else 
		      TomadeRamo=1
		   end if
	else
		TomadeRamo=1
	end if     
     end if 

Rst.close()
end function

function TieneCarga(codcli,anoed,periodoed)
Dim Sql
Dim Rst
       
	sql="Select codcli from ra_nota where codcli='" & codcli & "'" 
    sql=sql & " and ano='" & anoed & "' and periodo='" & periodoed & "'"
	'response.write(sql)
	'response.end
	set rst=conn.execute(sql)
	if not rst.eof then
	   TieneCarga=True
	else
	   TieneCarga=False   
	end if
end function

function TieneNotas(codcli,anoed,periodoed)
Dim Sql
Dim Rst
       
	sql="Select codcli from ra_nota where codcli='" & codcli & "'" 
    sql=sql & " and ano='" & anoed & "' and periodo='" & periodoed & "'"
	sql=sql & " and coalesce(estado,'')= ''"
	set rst=conn.execute(sql)
	if not rst.eof then
	   TieneNotas=False
	else
	   TieneNotas=True   
	end if
end function

%>