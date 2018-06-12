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
     if RutIngresado = "" then     
	 response.redirect("Alumnos.asp")
	 response.end
	 end if 
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
	
	Eliminar = Array("select","drop",";","delete","insert","update","xp_","'")
	NuevoStr = Texto
	
	For i = 1 to UBound(Eliminar)
		NuevoStr = Replace(NuevoStr, Eliminar(i), "")
	Next
	
	EliminaPalabras = NuevoStr	
End Function	

Function GetNombreProfe(CodProf)
Dim StrSql
Dim Rst

  StrSql = "Select Ap_Pater, Ap_Mater, Nombres from ra_profes with (nolock) where codprof = '" & CodProf & "' "

  Set Rst = Session("Conn").Execute(StrSql)

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
	
    'StrSql = "Select Horario from ra_seccio where CodRamo = '" & CodRamo & "' "
    'StrSql = StrSql & " and CodSecc = '" & CodSecc & "' "
    'StrSql = StrSql & " and ano = " & Ano 
    'StrSql = StrSql & " and Periodo = '" & Periodo & "' "
    'StrSql = StrSql & " and CodSede = '" & CodSede & "' and tipocurso='T'"
	
	StrSql = "Select dbo.HORARIO_RAMO_SECCION_orden('" & CodRamo & "','" & CodSecc & "','" &Ano & "','" &Periodo & "') as Horario" 
    
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

  StrSql = "Select Dia, CodMod from ra_horari with (nolock) where codramo = '" & CodRamo & "' "
  StrSql = StrSql & " And CodSede = '" & CodSede & "' "
  StrSql = StrSql & " And CodSecc = '" & CodSecc & "' "
  StrSql = StrSql & " And Ano = " & Ano 
  StrSql = StrSql & " And Periodo = " & Periodo
  StrSql = StrSql & " Order by case Dia when 'LUNES' then 1 when 'MARTES' then 2 when 'MIERCOLES' then 3 when 'JUEVES' then 4 when 'VIERNES' then 5 else 6 end, " & _
  " convert(numeric,CodMod)"
  'response.write(StrSql)
  'response.end
  Set Rst = Session("Conn").Execute(StrSql)
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

   StrSql = "Select tipo from ra_curric with (nolock) where CodRamo = '" & CodRamo & "' And CodPestud = '" & CodPestud & "' "
   Set Rst = Session("Conn").execute(StrSql)
   if Rst.Eof then
      GetTipoRamo = ""
   else
      GetTipoRamo = Ucase(Rst("Tipo"))
   end if   
   Rst.close()
   
End Function

Function EsPlanificado (CodRamo, CodSede, Ano, Periodo, CodCarr)
Dim StrSql
Dim Rst

   StrSql = "Select 1 from ra_seccio with (nolock) where Codramo = '" & Codramo & "' And Codsede = '" & CodSede & "' " & _
            " And Ano = " & ano & " And Periodo = " & Periodo & " and tipocurso='T' and codcarr='" & trim(CodCarr) & "'"
   if codramo="ENFE0002" then 	
       response.write(StrSql)	
   end if
   Set Rst = Session("Conn").execute(StrSql)
   EsPlanificado = not Rst.Eof 
   Rst.close()
End Function

Function EsPlanificadoXJornada (CodRamo, CodSede, Ano, Periodo, CodCarr,FiltraJornada)
Dim StrSql
Dim Rst

   If FiltraJornada = "S" Then
   		StrSql = "Select 1 from ra_seccio with (nolock) where Codramo = '" & Codramo & "' And Codsede = '" & CodSede & "' " & _
        " And Ano = " & ano & " And Periodo = " & Periodo & " and tipocurso='T' and codcarr='" & trim(CodCarr) & "' And (Jornada = '" & Session("Jornada") & "' OR Jornada2 = '" & Session("Jornada") & "')"
   else	
   		StrSql = "Select 1 from ra_seccio with (nolock) where Codramo = '" & Codramo & "' And Codsede = '" & CodSede & "' " & _
        " And Ano = " & ano & " And Periodo = " & Periodo & " and tipocurso='T' and codcarr='" & trim(CodCarr) & "'"
   end if
    
   Set Rst = Session("Conn").execute(StrSql)
   EsPlanificadoXJornada = not Rst.Eof 
   Rst.close()
End Function
  


Function GetEquivalente (CodRamo, CodSede, Ano, Periodo, CodCarr)
Dim StrSql, Equivalente
Dim Rst

   StrSql = "Select a.ramoequiv,a.codcarrequiv from ra_equiv a with (nolock), mt_carrer b with (nolock)"
   StrSql = StrSql & " Where a.codcarr = b.codcarr "
   StrSql = StrSql & " And b.Sede = '" & CodSede & "' "
   StrSql = StrSql & " And a.CodRamo = '" & CodRamo & "' and a.codcarr = '" & trim(CodCarr) & "'"
   'StrSql = StrSql & " Order by Prioridad "

   'response.write(StrSql)         

   Equivalente = ""         
   Set Rst = Session("Conn").execute(StrSql)
   Do while not Rst.Eof 
      if EsPlanificado(Rst("RamoEquiv"), CodSede, Ano, Periodo, Rst("CodCarrEquiv")) then
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
    
    StrSql = "Select Inscrito from ra_carga with (nolock) Where Codcli = '" & trim(Codcli) & "' "
    StrSql = StrSql & " And RamoEquiv = '" & trim(CodRamo) & "' "
    'StrSql = StrSql & " And CodSecc = '" & trim(CodSecc) & "' "
    StrSql = StrSql & " And ano = " & Ano 
    StrSql = StrSql & " And Periodo = " & Periodo
    
    Set Rst = Session("Conn").execute(StrSql)
    'Response.write(Strsql)
    if not Rst.eof then
       'response.write("Inscrito = " &  Rst("Inscrito"))
       EstaInscrito = (Rst("Inscrito") = "S")
    else
       EstaInscrito = false
    end if 
    
End Function

Function EstaPreInscrito(Codcli, Curso, CodRamo, CodSecc, Ano, Periodo, TipoCurso) 
Dim StrSql
Dim Rst
	
	if TipoCurso = "T" then
	    StrSql = "Select PreInscrito, Inscrito from ra_carga with (nolock) Where Codcli = '" & trim(Codcli) & "' "
	else
	    StrSql = "Select PreInscrito, Inscrito from ra_cargaactividad with (nolock) Where Codcli = '" & trim(Codcli) & "' and tipocurso='" & trim(tipocurso) & "'"		
	end if
    StrSql = StrSql & " And RamoEquiv_i = '" & trim(CodRamo) & "' "
	StrSql = StrSql & " And codramo = '" & trim(Curso) & "' "
    if trim(CodSecc) <> "" and tipocurso ="T" then
      StrSql = StrSql & " And CodSecc_i = " & CodSecc
    end if
	if TipoCurso <> "T" then
		StrSql = StrSql & " And CodSecc_i = " & CodSecc & " and tipocurso='" & TipoCurso & "'"
	end if
    StrSql = StrSql & " And ano = '" & Ano  & "' "
    StrSql = StrSql & " And Periodo = '" & Periodo & "' "

    Set Rst = Session("Conn").execute(StrSql)
    if not Rst.eof then
	   EstaPreInscrito = (Rst("Inscrito") = "S") Or (Rst("PreInscrito") = "S")
    else
       EstaPreInscrito = false
    end if 
    
End Function

Function EstaPreSolicitado(Codcli, Curso, CodRamo, CodSecc, Ano, Periodo, TipoCurso) 
Dim StrSql
Dim Rst
	
	if TipoCurso = "T" then
	    StrSql = "Select PreInscrito, Inscrito from ra_carga_log with (nolock) Where Codcli = '" & trim(Codcli) & "' "
'	else
'	    StrSql = "Select PreInscrito, Inscrito from ra_cargaactividad_log with (nolock) Where Codcli = '" & trim(Codcli) & "' and tipocurso='" & trim(tipocurso) & "'"		
	end if
    StrSql = StrSql & " And RamoEquiv_i = '" & trim(CodRamo) & "' "
	StrSql = StrSql & " And codramo = '" & trim(Curso) & "' "
    if trim(CodSecc) <> "" and tipocurso ="T" then
      StrSql = StrSql & " And CodSecc_i = " & CodSecc
    end if
	
	if TipoCurso <> "T" then
		StrSql = StrSql & " And CodSecc_i = " & CodSecc & " and tipocurso='" & TipoCurso & "'"
	end if
    StrSql = StrSql & " And ano = '" & Ano  & "' "
    StrSql = StrSql & " And Periodo = '" & Periodo & "' "
	
	'response.write(strsql)
	'response.end()
    Set Rst = Session("Conn").execute(StrSql)
    if not Rst.eof then
	   EstaPreSolicitado = (Rst("Inscrito") = "S") Or (Rst("PreInscrito") = "S")
    else
      EstaPreSolicitado = false
    end if 
    
End Function

Function EstaCursado(Codcli, CodRamo, CodSecc) 
Dim StrSql
Dim Rst

    StrSql = "Select 1 from ra_nota with (nolock) Where Codcli = '" & trim(Codcli) & "' "
    StrSql = StrSql & " And RamoEquiv = '" & trim(CodRamo) & "' "
    StrSql = StrSql & " And estado in ('A', 'E','I') "
    
    Set Rst = Session("Conn").execute(StrSql)
    if not Rst.eof then
      EstaCursado = True
      Rst.close()    
    else
      EstaCursado = false
    end if
End Function

Function ResetInscrip(CodCli, Ano, Periodo)
Dim StrSql
'if Ano <> 0 or Periodo <> 0 then
  StrSql = "Update ra_carga set ramoequiv_i = ramoequiv, codsecc_i = codsecc, PreInscrito = 'N' "
  StrSql = StrSql & " Where Codcli = '" & Codcli & "' "
  StrSql = StrSql & " And Ano = " & Ano
  StrSql = StrSql & " And Periodo = " & Periodo
  'response.write(StrSql)
  'response.end
  Session("Conn").execute StrSql

  StrSql = "delete TmpTest "
  StrSql = StrSql & " Where Codcli = '" & Codcli & "' "
  StrSql = StrSql & " And Ano = " & Ano
  StrSql = StrSql & " And Periodo = " & Periodo

  Session("Conn").execute StrSql
'end if
End Function


Function PrioridadRamo(Codcli, CodRamo, Ano, Periodo) 
Dim StrSql
Dim Rst

    StrSql = "Select Prioridad from ra_carga with (nolock) Where Codcli = '" & trim(Codcli) & "' "
    StrSql = StrSql & " And RamoEquiv = '" & trim(CodRamo) & "' "
    StrSql = StrSql & " And ano = " & Ano 
    StrSql = StrSql & " And Periodo = " & Periodo
    
    'Response.write(StrSql)
    Set Rst = Session("Conn").execute(StrSql)
    if not Rst.eof then
       PrioridadRamo = Rst("Prioridad") 
    else
       PrioridadRamo = -1
    end if 
    
End Function


Function VerificaTope(CodCli, Curso, CodRamo, CodSecc, CodSede, Ano, Periodo)
Dim StrSql
Dim Rst
   StrSql = " e_topehorario '" & CodCli & "','" & Curso & "','" & CodRamo & "','" & CodSecc & "', '" & CodSede & "'," & Ano & "," & Periodo
   'Response.write(Strsql)
    'Response.End
   Session("Conn").Execute StrSql
   StrSql = "Select 1 from ra_Topes where Codcli = '" & trim(CodCli) & "' "
   
   VerificaTope = BCL_ADO(StrSql, Rst) 
  
End Function

Function GetTopes(CodCli, Curso, CodRamo, CodSecc, CodSede, Ano, Periodo, Rst)
Dim StrSql

   StrSql = " e_topehorario '" & CodCli & "','" & Curso & "','" & CodRamo & "','" & CodSecc & "', '" & CodSede & "'," & Ano & "," & Periodo
   Session("Conn").Execute StrSql
   'Response.Write(StrSql) 
   'Response.End

   StrSql = "Select * from ra_Topes where Codcli = '" & CodCli & "' "
   BCL_ADO StrSql, Rst
End Function

Function GetNombreAlumno(codcli)
Dim StrSql
Dim Rst

 'Strsql="select nombre,paterno,materno from mt_client with (nolock) where codcli='" & codcli & "'"
 strsql="select b.nombre,b.paterno,b.materno from mt_alumno a with (nolock), mt_client b with (nolock)"
 strsql=strsql & " where a.codcli='" & codcli & "' and b.codcli = a.rut "
 Set Rst = Session("Conn").Execute(StrSql)

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


Function EstadoRamoReal(Codcli, CodRamo, TipoCurso) 
Dim StrSql
Dim Rst
	
	if TipoCurso = "T" then
	    StrSql = "Select estado from ra_nota with (nolock) Where Codcli = '" & trim(Codcli) & "' "
	else
		StrSql = "Select estado from ra_notaactividad with (nolock) Where Codcli = '" & trim(Codcli) & "' "
	end if 
    StrSql = StrSql & " And Ramoequiv = '" & trim(CodRamo) & "' "
    StrSql = StrSql & " Order by Ano desc, periodo desc "
    
    Set Rst = Session("Conn").execute(StrSql)
    if not Rst.eof then
      'EstadoRamoReal = ":" & Rst("estado")
      EstadoRamoReal = Rst("estado")
      Rst.close()    
    else
      EstadoRamoReal = ""
    end if
    'Response.write(StrSql)
    'response.end
End Function



Function GetCreditos(CodPestud, CredMin, CredMax, CredMinOtros, CredMaxOtros)
Dim StrSql
Dim Rst
Dim StrSql2
Dim Rst2

   StrSql = "Select CredMin, CredMax, CredMinOtros, CredMaxOtros from ra_pestud where CodPestud = '" & CodPestud & "' "
   'response.write(StrSql)
   'response.end()
   if BCL_ADO(Strsql, Rst) then
      CredMin = ValNulo(Rst("CredMin"), NUM_)
      CredMax = ValNulo(Rst("CredMax"), NUM_)
	  CredMinOtros = ValNulo(Rst("CredMinOtros"), NUM_)
      CredMaxOtros = ValNulo(Rst("CredMaxOtros"), NUM_)
   else
      CredMin = 0
      CredMax = 0   
      CredMinOtros = 0
      CredMaxOtros = 0   
   end if
   
   'StrSql2 = "Select creditos_max from RA_PESTUD_NIVEL where CodPestud = '" & CodPestud & "'  and nivel =(SELECT dbo.fn_Obtiene_Nivel_alumno('"& session("codcli") &"')) "
   
   StrSql2 = "Select creditos_max from RA_PESTUD_NIVEL where CodPestud = '" & CodPestud & "'  and nivel =(Select CONVERT(VARCHAR, MIN(CONVERT(INT, COALESCE(c.nivel,0))  ) )FROM ra_curric c , mt_alumno a WHERE c.codpestud = a.CODPESTUD AND a.codcli = '"& session("codcli") &"' AND c.codramo NOT IN ( SELECT codramo FROM ra_nota WHERE codcli = '"& session("codcli") &"' AND estado IN ('A', 'E', 'I') ) ) "
    'response.write(Strsql2)
	'response.end
   if BCL_ADO(Strsql2, Rst2) then
      CredMax = ValNulo(Rst2("creditos_max"), NUM_)   
   end if
    
End Function



'Function AlumnoConDeuda(rut)
'Dim StrSql
'Dim Rst

'    StrSql = "select codcli from mt_ctadoc where codcli = '" & Trim(rut) & "' and estado = 2 "
'        'str = str & " and saldo > 0 and fecven < '" & FechaServidor & "'"
'    StrSql = StrSql & " and saldo > 0 and fecven < getdate() "

'    AlumnoConDeuda = BCL_ADO(strsql, rst)
	
'End Function

Function ValidaDeuda()
    Dim StrSql
    Dim rst
    
	strsql = "select validadeuda from mt_parame "
    If BCL_ADO(strsql, rst) Then
		if ValNulo(rst("validadeuda"), STR_) = "S" then
			ValidaDeuda = 1
		else
			ValidaDeuda = 0
		end if
    Else
        ValidaDeuda = 0
    End If
End Function

Function TipoBloqueo 
Dim StrSql
Dim Rst

   StrSql = "Select tipovalidacionbloqueo from mt_parame"
   if bcl_ado(StrSql, Rst) then 
      TipoBloqueo = trim(Rst("tipovalidacionbloqueo"))
   else
      TipoBloqueo = ""
   end if
End Function

Function AlumnoConDeuda(rut)
Dim StrSql
Dim Rst

    AlumnoConDeuda = 0
    G_VALIDADEUDA = ValidaDeuda
	
    If G_VALIDADEUDA=1 Then
	
        If TipoBloqueo = "INTERNA" Then
			'Analiza si tiene bloqueo financiero
			StrSql = "select codcli from mt_ctadoc where codcli = '" & Trim(Rut) & "'"
			'StrSql = StrSql & " and saldo > 0 and fecven < getdate() and estado=2"
			'if BCL_ADO(strsql, rst) then
            ' 	AlumnoConDeuda = 1
			'else
			'	AlumnoConDeuda = 0
			'end if
			'response.write("Paso por aquÏ")
			'Response.end
			if GetDeudaProtesto(Trim(Rut))=true then
				AlumnoConDeuda = 1
			else
				AlumnoConDeuda = 0
			end if
			
			if GetDeudaMora(Trim(Rut))=true then
				AlumnoConDeuda = 1
			else
				AlumnoConDeuda = 0
			end if
			
			if GetDeudaCuotas(Trim(Rut))=true then
				AlumnoConDeuda = 1
			else
				AlumnoConDeuda = 0
			end if
			'analiza si tiene bloqueo de bliblioteca
            StrSql = "select biblioteca from mt_alumno where rut = '" & Trim(Rut) & "'"
			if BCL_ADO(strsql, rst) then
				if valnulo(rst("biblioteca"),str_)="S" then
					if AlumnoConDeuda = 1 then
	            		AlumnoConDeuda = 3
					else
	            		AlumnoConDeuda = 2					
					end if
				else
					if AlumnoConDeuda = 1 then
	            		AlumnoConDeuda = 1
					else
	            		AlumnoConDeuda = 0					
					end if
				end if
			else
				if AlumnoConDeuda = 1 then
	          		AlumnoConDeuda = 1
				else
	            	AlumnoConDeuda = 0					
				end if
			end if

        Else 

            StrSql = "SELECT bl1,bl2,resolucion FROM MT_BLOQUEOS "
            strsql = strsql & " WHERE COD_ALUMNO = '" & Trim(rut) & "' "
			'response.write(strsql)
            If BCL_ADO(strsql, rst) Then
			     response.write("Analisis de Bloqueos")
                If Trim(ValNulo(rst("resolucion"), STR_)) <> "S" Then
                    If ValNulo(rst("bl2"), NUM_) Then
					    'Alumno con deuda Financiera 
                        AlumnoConDeuda = 1
                    End If
                    If ValNulo(rst("bl1"), NUM_) Then
					    'Bloqueo de Biblioteca 
                        AlumnoConDeuda = 2
                    End If
                    If ValNulo(rst("bl1"), NUM_) And ValNulo(rst("bl2"), NUM_) Then
					    'Ambos Bloqueos
                        AlumnoConDeuda = 3
                    End If
                Else
				    'Tiene Resolucion
                    AlumnoConDeuda = 0
                End If
            Else
			    'No Existe en maestros Bloqueos No Tiene Bloqueos
                AlumnoConDeuda = 0
            End If
        End If
    Else

        AlumnoConDeuda = 0
		'No se valida Bloqueos
    End If
End Function
'------------------------------------------
Function GetDeudaProtesto ( Rut )
Dim Str
dim Rst

str=" select * from mt_ctadoc a, mt_docum b "
str= str & " where a.ctadoc=b.tipodoc and "
str= str & " a.codcli='"& rut &"' and "
str= str & " a.saldo > 0 and b.protesto='S' "
if bcl_ado(Str, Rst) then 
	GetDeudaProtesto=true
else
	GetDeudaProtesto=false
end if
End Function

	Function GetDiasMora
	dim str 
	dim rst
	str="select DIAS_MORA from mt_parame"
	if bcl_ado(Str, Rst) then 
		GetDiasMora=rst("DIAS_MORA")
	else
		GetDiasMora=0
	end if
	End Function

	Function GetCantidadCuotasMora
	dim str 
	dim rst
	str="select CUOTAS_MORA from mt_parame"
	if bcl_ado(Str, Rst) then 
		GetCantidadCuotasMora=rst("CUOTAS_MORA")
	else
		GetCantidadCuotasMora=0
	end if
	End Function

	Function GetDeudaMora (Rut)
	dim Str 
	dim Rst
	
	Str=" select * from mt_ctadoc where datediff(d,mt_ctadoc.fecven, getdate()) > "& GetDiasMora &" and "
	Str= Str & " mt_ctadoc.codcli='"& Rut &"' and mt_ctadoc.saldo > 0 "

	if bcl_ado(Str, Rst) then 
		GetDeudaMora=True
	else
		GetDeudaMora=false	
	end if

	End Function
	Function GetDeudaCuotas( Rut )
	dim str 
	dim rst
	dim Cuotas
	str=" select count(*) as Cantidad "
	Str= Str & " from mt_ctadoc "
	Str= Str & " where mt_ctadoc.fecven > getdate () "
	Str= Str & " and codcli='"& trim(Rut) &"' and saldo > 0 "
	if bcl_ado(Str, Rst) then 
		Cuotas=rst("Cantidad")
	else
		Cuotas=0
	end if
	if Cuotas > GetCantidadCuotasMora then 
		GetDeudaCuotas=true
	else
		GetDeudaCuotas=false
	end if	
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
	   
	   if isnull(Rst("v_Credito")) then	   
	       StrSqlP = "SELECT dbo.Fn_ValorParame('VALIDADEUDA')Parame"
		   if bcl_ado(StrSqlP, RstP) then 
			  if ucase(trim(RstP("Parame"))) = "S" then
			      CarreraValidaDeuda = true	
			  else
		  		  CarreraValidaDeuda = false
	  		  end if
		   end if		   
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

  StrSql = "Select Nombre_l from mt_carrer with (nolock) where codCarr = '" & CodCarr & "' "


  'Set Rst = Conn.Execute(StrSql)
  Set Rst = Session("Conn").execute(strsql)

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
     if bcl_ado(Sql,rst) then
        TotalRespuesta=valnulo(rst(0),num_)
     else
        TotalRespuesta=0
     end if
	 
end function


Function TotalRespuestasEncuesta(Encuesta)   
Dim Sql
Dim Rst

	'Total Pregunta Encuesta
     Sql="SELECT COUNT(*) FROM RA_ENCUESTA_DETALLE WHERE NRO_ENCUESTA = '" & Encuesta & "'"
'	 response.write(Sql)
'	 response.end()
     if bcl_ado(Sql,rst) then
        TotalRespuestasEncuesta=valnulo(rst(0),num_)
     else
        TotalRespuestasEncuesta=0
     end if
end function



'Function Encuestado(Codcli,anoed,periodoed)
'Dim Sql
'Dim Rst
'      'ModifiquË esta funcion debido a que discrimina a los egresados 
'		 sql="Select * from mt_Alumno where codcli='" & Codcli & "'"
'		 sql=sql & " and estacad in('VIGENTE','EGRESADO') and EncDoc='S' "
'		 sql=sql & " and ano_ed='" & anoed & "' and periodo_ed='" & periodoed & "'"
'		 'set rst=conn.execute(sql)
'		 if bcl_ado(Sql,rst) then
'			Encuestado=1
'		 else
'			Encuestado=0
'		 end if
'end Function

Function Encuestado(Codcli,anoed,periodoed)
Dim Sql
Dim Rst
    'ModifiquË esta funcion debido a que discrimina a los egresados 
	If TieneCarga(Codcli, AnoEd, PeriodoEd) = 0 Then
    	Encuestado = 2		
	Else				
		'Obtengo numero de encuesta docente vigente		
		sqlenc="select coalesce(numero,0)numero from RA_ENCUESTA Where CONVERT(DATE,getDate())>=FecIni And CONVERT(DATE,getDate()) <=FecFin AND TIPO_ENCUESTA=1"	
		
		if bcl_ado(sqlenc,rstenc) then
		
			numeroEnc=rstenc("numero")
			
			'Valido que tenga registros de encuesta, si no tiene y hay encuesta vigente es porque no entro a la opcion Y debe encuesta		
			sqlenc2="select top 1 1 from RA_ENCUESTADOS WHERE encuesta="& numeroEnc &" AND codcli='" & Codcli & "'"
			
			if bcl_ado(sqlenc2,rstenc2) then
				
				sqlenc3="SP_VERIFICAASIGNATURASCONTESTADAS '" & Codcli & "',"& numeroEnc &""
				if bcl_ado(sqlenc3,rstenc3) then
					'no ha contestado todas las encuestas
					Encuestado=0
					session("debeEncDoc")="SI"
				else
					Encuestado=1
					session("debeEncDoc")="NO"
				end if				
			else
				'para determinar si debe encuesta, primero validamos si tiene ramos/profes para encuestar, si no tiene encuestado=1
				Strsql ="sp_listaRamosEncDoc '" & session("codcli") & "','" & session("Codsede") & "',"& session("Ano_Ed") &","& session("Periodo_Ed") &","& session("NumeroMaximoPrueba") &",'" & session("Cerrada") & "'" 
				'response.write(Strsql)
				'response.end
				if bcl_ado(strsql,rs) then
					'no ha contestado encuesta
					Encuestado=0
					session("debeEncDoc")="SI"
				else
					Encuestado=4
					session("SinRamosEncDoc")="SI" 
					session("debeEncDoc")="NO"
				end if
			end if
		else
			'no hay encuesta
			Encuestado=3
		end if
		
		

		if Encuestado <>0 then
			'Nuevo - Obtengo si el alumno es nuevo
			StrSqlN = "Select 1 from mt_Alumno where Codcli = '" & Session("Codcli") & "' And Ano = Ano_mat and ano = dbo.Fn_ValorParame('anoadmision') " 
			if BCL_ADO(StrSqlN, Rstn) then
			  	TipoAlumno="Nuevo"
			else 
			  	TipoAlumno="Antiguo"
			end if
			
			'Obtengo numero de encuesta de satisfaccion	
			sqlenc = "select coalesce(numero,0)numero from RA_ENCUESTA Where CONVERT(DATE,getDate()) >=FecIni And CONVERT(DATE,getDate()) <=FecFin AND TIPO_ENCUESTA=7"	
			sqlenc = sqlenc & "AND (TIPO_ALUMNO ='"& TipoAlumno &"' OR COALESCE(TIPO_ALUMNO,'Todos') = 'Todos')"
			
			if bcl_ado(sqlenc,rstenc) then
			
				numeroEnc=rstenc("numero")
				sqlenc3="select top 1 1 from RA_ENCUESTADOS WHERE encuesta="& numeroEnc &" AND codcli='" & Codcli & "' AND Evaluado='SI'"
				if bcl_ado(sqlenc3,rstenc3) then
					Encuestado=1
				else
					Encuestado=0
					session("debeEncSat")="SI"
				end if	
			'else
				'no hay encuesta
			'	Encuestado=3
			end if			
		end if
		
		'nuevo, para encuesta de alumnos nuevos, tipo 11
		if Encuestado <>0 then
			strEncAn = "SELECT COALESCE(dbo.Fn_ValorParame('VALIDA_ENCUESTA_AN_EN_PORTAL'),'NO')Parame, "
			strEncAn = strEncAn & "coalesce(dbo.Fn_ValorParame('ENCUESTA_AN'),'0')ENCUESTA_AN, "
			strEncAn = strEncAn & "coalesce(dbo.Fn_ValorParame('ENCUESTA_AN2'),'0')ENCUESTA_AN2, "
			strEncAn = strEncAn & "coalesce(dbo.Fn_ValorParame('ENCUESTA_AN3'),'0')ENCUESTA_AN3 "
			
			if bcl_ado(strEncAn,rstEncAn) then
				USAENCUESTAAN = rstEncAn("parame")
				ENCUESTA_AN  = rstEncAn("ENCUESTA_AN")
				ENCUESTA_AN2 = rstEncAn("ENCUESTA_AN2")
				ENCUESTA_AN3 = rstEncAn("ENCUESTA_AN3")
			end if
			
			StrSqlN2 = "Select 1 from mt_Alumno where Codcli = '" & Session("Codcli") & "' And Ano = Ano_mat and ano = dbo.Fn_ValorParame('anoadmision') and nuevo ='S'" 
			if BCL_ADO(StrSqlN2, Rstn2) then
			  	EsAlumnoNuevo_EAN = "S"
			else 
			  	EsAlumnoNuevo_EAN = "N"
			end if 
			
			'Si esta activada la validacion y ademas el alumno es nuevo
			if USAENCUESTAAN ="SI" and EsAlumnoNuevo_EAN = "S" then
				
				if ENCUESTA_AN <> "0" then
					'Obtengo numero de encuesta de alumnos nuevos 1
					sqlenc_an = "select coalesce(numero,0)numero from RA_ENCUESTA Where CONVERT(DATE,getDate()) >=FecIni And CONVERT(DATE,getDate()) <=FecFin AND TIPO_ENCUESTA=11 and numero = "& ENCUESTA_AN &" "
					if bcl_ado(sqlenc_an,rstenc_an) then
						numeroEnc_an=rstenc_an("numero")
						
						sqlenc4="select top 1 1 from RA_ENCUESTADOS WHERE encuesta="& numeroEnc_an &" AND codcli='" & Codcli & "' AND Evaluado='SI'"
						if bcl_ado(sqlenc4,rstenc4) then
							Encuestado=1
						else
							Encuestado=0
						end if
						
					end if	
				end if
				
				if Encuestado <> 0 and ENCUESTA_AN2 <> "0" then
					'Obtengo numero de encuesta de alumnos nuevos 2
					sqlenc_an2 = "select coalesce(numero,0)numero from RA_ENCUESTA Where CONVERT(DATE,getDate()) >=FecIni And CONVERT(DATE,getDate()) <=FecFin AND TIPO_ENCUESTA=11 and numero = "& ENCUESTA_AN2 &" "
					if bcl_ado(sqlenc_an2,rstenc_an2) then
						numeroEnc_an2=rstenc_an2("numero")
						
						sqlenc42="select top 1 1 from RA_ENCUESTADOS WHERE encuesta="& numeroEnc_an2 &" AND codcli='" & Codcli & "' AND Evaluado='SI'"
						if bcl_ado(sqlenc42,rstenc42) then
							Encuestado=1
						else
							Encuestado=0
						end if
						
					end if
				end if 
				
				if Encuestado <> 0 and ENCUESTA_AN3 <> "0" then
					'Obtengo numero de encuesta de alumnos nuevos 3
					sqlenc_an3 = "select coalesce(numero,0)numero from RA_ENCUESTA Where CONVERT(DATE,getDate()) >=FecIni And CONVERT(DATE,getDate()) <=FecFin AND TIPO_ENCUESTA=11 and numero = "& ENCUESTA_AN3 &" "
					if bcl_ado(sqlenc_an3,rstenc_an3) then
						numeroEnc_an3=rstenc_an3("numero")
						
						sqlenc43="select top 1 1 from RA_ENCUESTADOS WHERE encuesta="& numeroEnc_an3 &" AND codcli='" & Codcli & "' AND Evaluado='SI'"
						if bcl_ado(sqlenc43,rstenc43) then
							Encuestado=1
						else
							Encuestado=0
						end if
						
					end if
				end if 
						
			end if 
		
		end if
		 	
	end if
	
end Function

Function EncuestadoED(Codcli,anoed,periodoed)
Dim Sql
Dim Rst
    'ModifiquË esta funcion debido a que discrimina a los egresados 
	If TieneCarga(Codcli, AnoEd, PeriodoEd) = 0 Then
    	EncuestadoED = 2
	Else				
		'Obtengo numero de encuesta docente vigente		
		sqlenc="select numero from RA_ENCUESTA Where CONVERT(DATE,getDate())>=FecIni And CONVERT(DATE,getDate()) <=FecFin AND TIPO_ENCUESTA=1"	
			
		if bcl_ado(sqlenc,rstenc) then
		
			numeroEnc=valnulo(rstenc("numero"),num_)
			
			'Valido que tenga registros de encuesta, si no tiene y hay encuesta vigente es porque no entro a la opcion Y debe encuesta		
			sqlenc2="select top 1 1 from RA_ENCUESTADOS WHERE encuesta="& numeroEnc &" AND codcli='" & Codcli & "'"
			
			if bcl_ado(sqlenc2,rstenc2) then
				
				'sqlenc3="select top 1 1 from RA_ENCUESTADOS WHERE encuesta="& numeroEnc &" AND codcli='" & Codcli & "' AND Evaluado='NO'"
				sqlenc3="SP_VERIFICAASIGNATURASCONTESTADAS '" & Codcli & "',"& numeroEnc &""
				if bcl_ado(sqlenc3,rstenc3) then
					'no ha contestado todas las encuestas
					EncuestadoED=0
				else
					EncuestadoED=1
				end if				
			else
				'para determinar si debe encuesta, primero validamos si tiene ramos/profes para encuestar, si no tiene encuestado=1
				Strsql ="sp_listaRamosEncDoc '" & session("codcli") & "','" & session("Codsede") & "',"& session("Ano_Ed") &","& session("Periodo_Ed") &","& session("NumeroMaximoPrueba") &",'" & session("Cerrada") & "'" 
				'response.write(Strsql)
				'response.end
				if bcl_ado(strsql,rs) then
					'no ha contestado encuesta
					EncuestadoED=0
				else
					EncuestadoED=1
				end if
			end if
		else
			'no hay encuesta
			EncuestadoED=1
		end if
		
	end if
	
end Function 


'function TomadeRamo(codigo,codcli,anoed,periodoed)
'Dim Sql
'Dim Rst
	 
'	sql="Select coalesce(RequisitoEd,'N') from mt_carrer where codcarr='" & codigo & "'" 
    
'     'set rst=conn.execute(sql)
'     if bcl_ado(Sql,rst) then
'        if rst(0) ="S" then		  
'		   if Encuestado(codcli,anoed,periodoed)=0 then
'		      TomadeRamo=0
'		   else 
'		      TomadeRamo=1
'		   end if
'		else
'			TomadeRamo=0
'		end if
'	 else
'		TomadeRamo=0     
 '    end if 

'Rst.close()
'end function

function TomadeRamo(codigo,codcli,anoed,periodoed)
Dim Sql
Dim Rst
	 
    if Encuestado(codcli,anoed,periodoed)=0 then
		TomadeRamo=0
	else 
		TomadeRamo=1
	end if

end function

function TomadeRamoED(codigo,codcli,anoed,periodoed)
Dim Sql
Dim Rst
	 
    sql="Select coalesce(RequisitoEd,'N') from mt_carrer where codcarr='" & codigo & "'" 
	if bcl_ado(Sql,rst) then
		if rst(0) ="S" then		  
			if EncuestadoED(codcli,anoed,periodoed)=0 then
				TomadeRamoED=0
			else 
				TomadeRamoED=1
			end if
		else
			TomadeRamoED=1
		end if
	else
		TomadeRamoED=0     
	end if 

Rst.close()
end function

function TieneCarga(codcli,anoed,periodoed)
Dim Sql
Dim Rst
    if anoed="" then
		anoed=0
	end if  
	 
	sql="Select codcli from ra_nota where codcli='" & codcli & "'" 
    'sql=sql & " and ano='" & anoed & "' and periodo='" & periodoed & "'"
	sql=sql & " and ano='" & anoed & "' "
	
	'response.write(Sql)
	'response.end
	if bcl_ado(Sql,rst) then
	   TieneCarga=1
	else
	   TieneCarga=0   
	end if

end function

function TieneNotas(codcli,anoed,periodoed)
Dim Sql
Dim Rst
'dim Carrera

Carrera=GetCarreraAlumno(Codcli)
if GetHabilitaEncuesta(Carrera)=1 then
  	 TieneNotas=1   
else
 		sql = "Select a.codcli from ra_nota a, ra_seccio b where a.codcli='" & codcli & "'" 
		sql = sql & " and a.ano='" & anoed & "'"
		'and a.periodo='" & periodoed & "'"
		sql = sql & " and coalesce(a.estado,'')= ''"
		sql = sql & " and a.ramoequiv = b.codramo"
		sql = sql & " and a.codsecc = b.codsecc"
		sql = sql & " and a.ano = b.ano"
		sql = sql & " and a.periodo = b.periodo"
		'sql = sql & " and b.codsecc = '" & trim(session("codsede")) & "'"
		sql = sql & " and b.codsede = '" & trim(session("codsede")) & "'"
		sql = sql & " and b.cerrada = '" & trim(session("cerrada")) & "'"
		'response.write(sql)
		if bcl_ado(Sql,rst) then
		   TieneNotas=0
		else
		   TieneNotas=1   
		end if
		'response.write(tienenotas)
		'response.end()
end if	
end function


Function GetPermiso(carrera, estado, bloqueoF,bloqueoB,codigo,nivel,EstadoMatriculado, bloqueoA)
Dim Str 
Dim Rst

Str = "select permiso from ra_perfilinternet where codcarr='"& carrera &"' and estacad ='"& estado &"' and "
Str = Str & " estafinan='"& bloqueoF &"' and  estabiblio='"& BloqueoB &"'" 
str = str & " and codigo='"& Codigo &"' and nivel='"& nivel &"' "
str = str & " and estaadmin='"& bloqueoA &"'"
str = str & " and (matriculado='"& EstadoMatriculado &"' or matriculado = 'TODOS LOS ESTADOS') "

'response.write(str)
'response.end()
if codigo="INT_001" then
	'response.write(str)
end if
if bcl_ado(Str,Rst) then
	GetPermiso=valnulo(Rst("Permiso"), num_)
else
   	GetPermiso=0
end if

end function


Function GetPerfil(carrera, estado, bloqueoF,bloqueoB,EstadoMatriculado, bloqueoA)
Dim Str 
Dim Rst

Str = "select perfil from ra_perfilinternet where codcarr='"& carrera &"' and (estacad ='"& estado &"' or estacad = 'TODOS LOS ESTADOS')  and "
Str = Str & " (estafinan='"& bloqueoF &"' or estafinan = 'TODOS LOS ESTADOS') and  (estabiblio='"& BloqueoB &"' or estabiblio = 'TODOS LOS ESTADOS') " 
str = str & " and (matriculado='"& EstadoMatriculado &"' or matriculado = 'TODOS LOS ESTADOS') "
str = str & " and (estaadmin='"& bloqueoA &"' or estaadmin = 'TODOS LOS ESTADOS') "

'response.write(str)
'response.end()

if bcl_ado(Str,Rst) then
	GetPerfil=valnulo(Rst("perfil"), num_)
	
 	'Actualiza perfil del alumno
	 ActP=" exec Pr_ActualizaPerfil "& Session("id_usuario") &" ,"& Rst("perfil") &""
		'response.write(ActP)
		'response.end()
	Session("Conn").Execute ActP
	
else
   	GetPerfil=0
end if

end function



Function GetEstado(Rut, Carrera)
dim Str
dim Rst
dim Str2
dim Rst2

Str ="Select top 1 estacad from mt_alumno where rut='"& Rut &"' and codcarpr='"& Carrera &"' and estacad='VIGENTE' order by fec_mat desc"
if bcl_ado(Str,rst) then
	GetEstado=Valnulo(rst(0), Str_)
else
	Str2 ="Select top 1 estacad from mt_alumno where rut='"& Rut &"' and codcarpr='"& Carrera &"' order by fec_mat desc"
	if bcl_ado(Str2,rst2) then
		GetEstado=Valnulo(rst2(0), Str_)
	end if
end if

end Function

Function GetEstado2(Rut)
DIM Str
dim Rst

strParameNM="SELECT coalesce(dbo.Fn_ValorParame('FILTRAESTADOVIGENTEPA'),'NO')Parame"
set rstParameNM= Session("Conn").Execute(strParameNM)	
	
if not rstParameNM.eof then
	FILTRAESTADOVIGENTEPA=rstParameNM("Parame")
end if

Str ="Select top 1 estacad from mt_alumno where rut='"& Rut &"'"
if FILTRAESTADOVIGENTEPA = "SI" then
	Str = Str & " and estacad='VIGENTE' "
end if 
Str = Str & " order by ano desc"


if bcl_ado(Str,rst) then
	GetEstado2=Valnulo(rst(0), Str_)
end if

end Function

Function GetMensaje(carrera, estado, bloqueoF, bloqueoB,codigo,nivel)
Dim Str 
Dim Rst
Str="select mensaje from ra_perfilinternet where codcarr='"& carrera &"' and estacad ='"& estado &"' and "
Str=Str & "estafinan='"& bloqueoF &"' and  estabiblio='"& BloqueoB &"' and codigo='"& Codigo &"' and nivel='"& nivel &"' "

'response.write(str)
'response.end()
if bcl_ado(Str,Rst) then
	GetMensaje=valnulo(Rst("mensaje"), Str_)
end if
end function

Function GetVariasCarreras (RutCli)
dim str 
dim rst

strParameNM="SELECT coalesce(dbo.Fn_ValorParame('FILTRAESTADOVIGENTEPA'),'NO')Parame"
set rstParameNM= Session("Conn").Execute(strParameNM)	
	
if not rstParameNM.eof then
	FILTRAESTADOVIGENTEPA=rstParameNM("Parame")
end if
		
str="select rut, count(rut)as Cantidad "
str=str & " from mt_alumno "
str=str & " where rut='"& RutCli &"' "

if FILTRAESTADOVIGENTEPA ="SI" then
	str=str & " and estacad='VIGENTE' "
end if 

str=str & " group by rut  "
str=str & " having count(rut) >1 "


'response.write(str)
'response.end()

if bcl_ado(Str,Rst) then
	GetVariasCarreras="Verdadero"
else
	GetVariasCarreras ="Falso"
end if
end Function

Function GetHabilitaEncuesta (Codcarr)
dim Str 
dim Rst 

	Str="select encuestasinNota from mt_carrer where codcarr='"& Codcarr &"'"
	'response.write(str)
	if bcl_ado(Str,Rst) then
		if Rst("encuestasinNota")= null or Rst("encuestasinNota")="N" or  Rst("encuestasinNota")="" then
			GetHabilitaEncuesta =0
		else
			GetHabilitaEncuesta =1
		end if
	end if
	'response.write(GetHabilitaEncuesta)
	'response.end()
	
End Function

Function GetCarreraAlumno(codcli)
dim Str
dim Rst 
Str="select top 1 c.codcarr from " 
Str=Str & " mt_alumno a, mt_carrer c "
Str=Str & " where a.codcarpr=c.codcarr and "
Str=Str & " a.codcli='"& codcli &"'"
'response.write(str)
'response.end()
	if bcl_ado(Str,Rst) then
		GetCarreraAlumno=rst("Codcarr")
	end if
End function

Function BloqueoAdministrativoCarrera(CodCarr)

Dim StrSql
Dim Rst

	StrSql = "Select coalesce(validadocpend,'N') as bloqueoAdm from mt_carrer  where CodCarr = '" & CodCarr & "'"	
	'response.write(StrSql)
	'response.end
	if bcl_ado(StrSql, Rst) then 	
	   BloqueoAdministrativoCarrera = trim(rst("bloqueoAdm"))
	   rst.close
	end if
	
End Function

Function ValidaDocPend(Codigo)
Dim Rst
Dim Str

	
ValidaDocPend = 0

str = "select VALIDADOCPEND from mt_carrer a, mt_alumno b where "
str = str & " a.codcarr=b.codcarpr and b.rut='" & Trim(Codigo) & "'"
'response.write(str)
'response.end
If bcl_ado(str, rst) Then
    If Trim(ValNulo(rst("VALIDADOCPEND"), STR_)) = "S" Then
        ValidaDocPend = 1
    End If
End If

End Function

Function SCase(strPalabra)
    Dim ix
	Dim iy
    
    ix = Len(strPalabra)
    If ix < 2 Then
        SCase = LCase(strPalabra)
        Exit Function
    End If
    ix = 1
    Do
        iy = InStr(ix, strPalabra, " ", vbTextCompare)
        If iy = 0 Then iy = Len(strPalabra) + 1
        If iy = ix + 1 Then
            strPalabra = Left(strPalabra, ix - 1) & _
                         LCase(Mid(strPalabra, ix, 1)) & _
                         Mid(strPalabra, ix + 1)
        Else
            strPalabra = Left(strPalabra, ix - 1) & _
                         UCase(Mid(strPalabra, ix, 1)) & _
                         LCase(Mid(strPalabra, ix + 1, iy - ix)) & _
                         Mid(strPalabra, iy + 1)
        End If
        ix = iy + 1
    Loop Until ix >= Len(strPalabra)
    SCase = strPalabra
End Function

Function BloqueoAdministrativoAlumno(Rut)
Dim StrSql, Str
Dim Rst
dim DocumentosRequisito

   DocumentosRequisito=""
 '  response.write("BloqueoAdministrativoAlumno")
   
   If ValidaDocPend(Rut) = 1 Then    
        'str = "select b.descripcion from mt_certifalu a, mt_certif b, mt_alumno c"
        'str = str & " where a.codcli=c.rut and c.rut='" & Rut & "'"
        'str = str & " and a.codcertif = b.codcertif and entregado='N'"
		'str = str & " and a.codcertif in ('1','2','4') "
		'str = str & " and c.estacad  in ('egresado','vigente')"
				
		str = "select b.descripcion from mt_certifalu a, mt_certif b, mt_alumno c,mt_client cli "
		str = str & "where a.codcli=c.rut and c.rut='" & IdRut & "' "
		str = str & "and a.codcertif = b.codcertif and entregado='N' "
		str = str & "and a.codcertif in (select CODCERTIF from MT_CERTVIA  WHERE COD_VIA=cli.VIADMISION) "
		str = str & "and c.estacad  in ('egresado','vigente') "
		str = str & "AND cli.codcli=c.rut "
	
		'response.write(str)
        'response.end()
		If bcl_ado(str, rst) Then
            DocumentosRequisito = "Tiene documentos pendientes,"
            While Not rst.EOF
                DocumentosRequisito = DocumentosRequisito & " " & UCase(ValNulo(rst("descripcion"), STR_))
                If rst.recordcount > 1 Then
                    Coma = "S"
                End If
                rst.MoveNext
                If Coma = "S" And Not rst.EOF Then
                    DocumentosRequisito = DocumentosRequisito & ","
                End If
            Wend			
        Else
            DocumentosRequisito = ""
        End If
    Else
        DocumentosRequisito = ""
    End If
    BloqueoAdministrativoAlumno = DocumentosRequisito
	
End Function

Function GetTipoRegimen(Rut, Carrera )
Dim StrSql
Dim Rst
StrSql="select b.tipo_regimen from mt_alumno a, mt_carrer b  "
StrSql=StrSql & " where a.codcarpr=b.codcarr and a.rut='"& Rut &"' "
StrSql=StrSql & " and a.codcarpr='"& Carrera &"'"
if bcl_ado(StrSql, Rst) then 
	GetTipoRegimen=rst("tipo_regimen")
	else
	GetTipoRegimen=0
end if
End Function

Function BloqueoBibliotecaCarrera(CodCarr)
Dim StrSql
Dim Rst

StrSql = "Select coalesce(v_bloqueo,'N') as v_bloqueo from mt_carrer  where CodCarr = '" & CodCarr & "' "

if bcl_ado(StrSql, Rst) then 
   BloqueoBibliotecaCarrera = trim(rst(0))
end if
End Function

Function BloqueoBibliotecaAlumno(Rut)
Dim StrSql
Dim Rst

StrSql = "Select TOP 1 coalesce(biblioteca,'N') as Biblioteca from mt_alumno  where Rut = '" & Rut & "' order by ano desc"
if bcl_ado(StrSql, Rst) then 
   BloqueoBibliotecaAlumno = trim(rst(0))
end if
End Function

Function GetTipoRegimen(Rut, Carrera )
Dim StrSql
Dim Rst
StrSql="select b.tipo_regimen from mt_alumno a, mt_carrer b  "
StrSql=StrSql & " where a.codcarpr=b.codcarr and a.rut='"& Rut &"' "
StrSql=StrSql & " and a.codcarpr='"& Carrera &"'"
if bcl_ado(StrSql, Rst) then 
	GetTipoRegimen=rst("tipo_regimen")
	else
	GetTipoRegimen=0
end if
End Function

Function GetPeriodoEd(Rut, Carrera)
Dim StrSql
Dim Rst
Dim Ano
Dim Periodo
Dim NumeroMaximoPeriodo
Dim AnoEd
Dim PeriodoEd
'Query que trae el Periodo Apterurado, Numero Maximo de Periodo tanto del Plan como de la Carrera
StrSql=" select b.ano,b.periodo,a.NumMaxPer, "
StrSql=StrSql & " CONVERT(INTEGER,c.num_max_periodo) as NMPLAN,CONVERT(INTEGER,d.num_max_periodo) as NMCARRERA "
StrSql=StrSql & " from mt_alumno a ,mt_carreraperiodo b, ra_pestud c, mt_carrer d "
StrSql=StrSql & " where a.codcarpr= b.codcarr  "
StrSql=StrSql & " and a.codpestud=b.codpestud "
StrSql=StrSql & " and a.codpestud=c.codpestud "
StrSql=StrSql & " and a.codcarpr=c.codcarr "
StrSql=StrSql & " and b.codcarr=d.codcarr "
StrSql=StrSql & " and c.codcarr=d.codcarr AND a.ESTACAD<>'ELIMINADO'"
StrSql=StrSql & " and a.rut='"& Rut &"' "
StrSql=StrSql & " and a.codcarpr='"& Carrera &"' "

if bcl_ado(StrSql, Rst) then 
	'Si es Tipo Regimen=1 tomamos el periodo del Plan de lo contrario de la carrera
	if GetTipoRegimen(Rut, Carrera ) =1 then
		NumeroMaximoPeriodo=Rst("NMPLAN")
	else
		NumeroMaximoPeriodo=Rst("NMCARRERA")		
	end if
	'response.write(NumeroMaximoPeriodo)
	'response.end
	Ano=Rst("ano")
	Periodo=rst("periodo")
	Session("ano")=Rst("ano")
	Session("Periodo")=rst("periodo")
	
else
	StrSql=" select CONVERT(INTEGER,c.num_max_periodo) as NMPLAN,CONVERT(INTEGER,d.num_max_periodo) as NMCARRERA "
	StrSql=StrSql & " from mt_alumno a , ra_pestud c, mt_carrer d "
	StrSql=StrSql & " where a.codpestud=c.codpestud "
	StrSql=StrSql & " and a.codcarpr=c.codcarr "
	StrSql=StrSql & " and c.codcarr=d.codcarr "
	StrSql=StrSql & " and a.rut='"& Rut &"' "
	StrSql=StrSql & " and a.codcarpr='"& Carrera &"' "
	'response.write(StrSql)
	'response.end()
	
	if bcl_ado(StrSql, Rst) then 
		'Si es Tipo Regimen=1 tomamos el periodo del Plan de lo contrario de la carrera
		if GetTipoRegimen(Rut, Carrera ) =1 then
			NumeroMaximoPeriodo=Rst("NMPLAN")
		else
			NumeroMaximoPeriodo=Rst("NMCARRERA")		
		end if
	end if
end if
'analisis del Periodo Ed 
'response.write("este es el valor" & NumeroMaximoPeriodo)

if NumeroMaximoPeriodo <> 1 then
    if valnulo(Periodo, num_)=1 then
		Session("Periodo_Ed")= NumeroMaximoPeriodo
		Session("Ano_Ed")=valnulo(Ano, num_) - 1 
	else
		Session("Periodo_Ed")= valnulo(Periodo, num_) - 1
		Session("Ano_Ed")=valnulo(Ano, num_) 
    end if
else
	Session("Ano_Ed")=valnulo(Ano, num_) - 1
	Session("Periodo_Ed")= NumeroMaximoPeriodo
end if
	
	if GetEncuentaActual (Rut , Carrera )=1 then
		Session("Ano_Ed")=Session("ano")
		Session("Periodo_Ed")= Session("Periodo")
	end if	
		
	if Session("Ano_Ed") ="" then
		Session("Ano_Ed")=0
	end if 
	if Session("Periodo_Ed") ="" then
		Session("Periodo_Ed") =0
	end if
	if Session("ano") ="" then
		Session("ano")=0
	end if 
	if Session("Periodo") ="" then
		Session("Periodo") =0
	end if  
	
	Session("NumeroMaximoPrueba")=NumeroMaximoPeriodo
	
End Function


Function GetBuscaRamosInasistentes(Rut, Carrera )
Dim Strsql 
dim Rst 

StrSql=" select coalesce(b.contestasiginasistente,'N') as Campo  "
StrSql=StrSql & " from mt_alumno a, mt_carrer b "
StrSql=StrSql & " where a.codcarpr=b.codcarr and a.rut='"& Rut &"'"
StrSql=StrSql & " and a.codcarpr='"& Carrera &"'"
if bcl_ado(StrSql, Rst) then 
	if rst("Campo")="S" then
		GetBuscaRamosInasistentes=1
	else
		GetBuscaRamosInasistentes=0
	end if
else
	GetBuscaRamosInasistentes=0
end if
	Session("RamosInasistentes")=valnulo(GetBuscaRamosInasistentes, num_)
End Function

Function GetHabilitaBotonEncuenta(Rut, Carrera )
Dim Strsql 
dim Rst 

StrSql=" select coalesce(b.OMITEENCUESTA,'N') as Campo  "
StrSql=StrSql & " from mt_alumno a, mt_carrer b "
StrSql=StrSql & " where a.codcarpr=b.codcarr and a.rut='"& Rut &"'"
StrSql=StrSql & " and a.codcarpr='"& Carrera &"'"
if bcl_ado(StrSql, Rst) then 
	if rst("Campo")="S" then
		GetHabilitaBotonEncuenta=1
	else
		GetHabilitaBotonEncuenta=0
	end if
else
	GetHabilitaBotonEncuenta=0
end if
Session("HABILITAENCUESTA")=valnulo(GetHabilitaBotonEncuenta, num_)
End Function

Function GetEncuentaActual(Rut, Carrera)
Dim Strsql 
dim Rst 

StrSql=" select coalesce(b.ENCUESTAANTES,'N') as Campo  "
StrSql=StrSql & " from mt_alumno a, mt_carrer b "
StrSql=StrSql & " where a.codcarpr=b.codcarr and a.rut='"& Rut &"'"
StrSql=StrSql & " and a.codcarpr='"& Carrera &"'"
 
if bcl_ado(StrSql, Rst) then 
	if rst("Campo")="S" then
		GetEncuentaActual=1
	else
		GetEncuentaActual=0
	end if
else
	GetEncuentaActual=0
end if
End Function

Function MiMensajeBloqueo(CodBloqueo)
Dim StrSql
Dim Rst

StrSql = "select codbloqueo,mensaje,departamento  "
StrSql =StrSql & " from ra_mensajebloqueo "
StrSql =StrSql & " where codbloqueo="& VALNULO(CodBloqueo, NUM_) &""

if bcl_ado(StrSql, Rst) then 
    MiMensajeBloqueo = trim(rst("mensaje")) + " " + trim(rst("departamento")) 
else
	MiMensajeBloqueo="No podr&aacute; acceder a esta opci&oacute;n ."    
end if

End Function

Function EstaHabilitada(CodAplicacion)
Dim StrSql
Dim Rst

StrSql = "select coalesce(Habilitada,'N') as Habilitada "
StrSql =StrSql & " from ra_opinternet "
StrSql =StrSql & " where CODIGO='" & VALNULO(CodAplicacion, STR_) & "'"

if bcl_ado(StrSql, Rst) then 
	if trim(rst("Habilitada")) ="N" then
			EstaHabilitada="N"
	else
			EstaHabilitada="S"
	end if
else
	EstaHabilitada="N" 	   
end if

End Function


Function EstaHabilitadaNW(CodAplicacion)
Dim StrSql
Dim Rst

'StrSql = "SELECT COALESCE(men_activo,0) as men_activo FROM ca_menu a , ca_perfilesopciones b WHERE id_perfil = '"& Session("perfilNW") &"' and "
'StrSql =StrSql & " a.id_opcion = b.id_opcion AND a.id_menu ='" & VALNULO(CodAplicacion, STR_) & "'"

StrSql = "SELECT COALESCE(men_activo,0) as men_activo FROM ca_menu WHERE id_menu ='" & CodAplicacion & "'"



if bcl_ado(StrSql, Rst) then 
	if Rst("men_activo") = 1  then
		EstaHabilitadaNW="S"
	else
		EstaHabilitadaNW="N"
	end if
else
	EstaHabilitadaNW="N"
end if


End Function



Function GetPermisoNW(CodAplicacion)
Dim StrSql
Dim Rst

StrSql = "SELECT COALESCE(men_url,'N') as url FROM ca_menu a , ca_perfilesopciones b WHERE id_perfil = '"& Session("perfilNW") &"' and "
StrSql =StrSql & " a.id_opcion = b.id_opcion AND a.id_menu ='" & CodAplicacion & "'"

'response.write(strsql)
'response.end()
'Set Rst = Session("Conn").execute(StrSql)
'response.write(Rst("url"))
if bcl_ado(StrSql, Rst) then 
	if Rst("url") = N or Rst("url") = "" then
		GetPermisoNW="N"
	else
		GetPermisoNW=Rst("url")
	end if
else
	GetPermisoNW="N"
end if


End Function



Function GetFormatoEspecial (CODCLI)
Dim Str
Dim RstSq

str=" SELECT COALESCE(B.FORMATOESPECIALP,'NO') AS FORMATOESPECIALP  "
str = str & " FROM MT_ALUMNO A, MT_CARRER B "
str = str & " WHERE A.CODCARPR=B.CODCARR AND  "
str = str & " A.CODCLI='"& CODCLI &"' "

if bcl_ado(Str, RstSq) then 
	if trim(RstSq("FORMATOESPECIALP")) ="NO" then
			GetFormatoEspecial="N0"
	else
			GetFormatoEspecial="SI"
	end if
else
	GetFormatoEspecial="N0" 	   
end if

End Function

Function GetTieneRamosAnuales (Codcli)
dim str 
dim Rstsq

str= "	select distinct b.ramoequiv  "
str = str & "	from mt_alumno a, ra_nota b , ra_pestud c, ra_curric d, ra_ramo e "
str = str & "	where  a.codpestud=c.codpestud and "
str = str & "	a.codcarpr=c.codcarr and "
str = str & "	c.codpestud=d.codpestud and "
str = str & "	a.codpestud=c.codpestud and "
str = str & "	b.codramo=d.codramo and "
str = str & "	b.codramo=e.codramo and "
str = str & "	e.codramo=d.codramo and "
str = str & "	e.num_max_per=1 and "
str = str & "   a.codcli=b.codcli and "
str = str & "	a.codcli='"& codcli &"'   "
if bcl_ado(Str, RstSq) then 
	GetTieneRamosAnuales=1
else
	GetTieneRamosAnuales=0
end if
RstSq.close

End Function

Function GetPlanOtrosRamos(codpestud,codramo)
dim Str
dim Rst 
	Str = "select codpestudadic from ra_planesadic "
	Str = Str & " where codpestud='" & trim(codpestud) & "'"
	Str = Str & " and  codpestudadic in "
	str = str & " (select a.codpestud " 
	str = str & " from ra_pestud a, ra_curric b"
	str = str & " where a.codpestud=b.codpestud and b.codramo='" & trim(codramo) & "')"
	if bcl_ado(Str,Rst) then
		GetPlanOtrosRamos = rst("Codpestudadic")
	else
		GetPlanOtrosRamos = ""
	end if

End function

function ValidaRamosPendientes(Codcli,ano,periodo)     
dim sql 
dim Rst

	 	  sql="Select * from ra_encuestados where Ano=" & Ano & ""
		  sql=sql & " and Periodo='" & periodo & "' and CodCli='" & Codcli & "' and encuesta=(select COALESCE(NUMERO,0) from RA_ENCUESTA WHERE convert(date,GETDATE()) BETWEEN fecini AND fecfin AND TIPO_ENCUESTA=1 )"
		  sql=sql & " and Evaluado='NO'" 
		  
		  if bcl_ado(sql,rst) then
			 ValidaRamosPendientes=0			  
		  else
			 ValidaRamosPendientes=1
		  end if  	 
end function
  
  
Function EsAlumnoAntiguo(Rut,CodCarr,AnoAdmi,PeriodoAdmi)
Dim StrSql
Dim Rst

'esto query lee si el alumno es nuevo y sino retorna  EsAlumnoAntiguo = 1


'StrSql = "Select rut from mt_alumno where rut = '" & trim(Rut) & "'" 
'StrSql = StrSql & " and CodCarpr = '" & CodCarr & "'"
'StrSql = StrSql & " and ano = " & AnoAdmi & ""
'StrSql = StrSql & " and periodo = '" & trim(PeriodoAdmi) & "' and coalesce(nuevo,'')<>'N'"

StrSql ="Select 1 from mt_Alumno where Codcli = '" & Session("Codcli") & "' And Ano = Ano_mat and ano = dbo.Fn_ValorParame('anoadmision') and nuevo='S'"

'response.write(StrSql)
'response.end()

if bcl_ado(StrSql, Rst) then 
   EsAlumnoAntiguo = 0 'Es nuevo
else
   EsAlumnoAntiguo = 1 'Es antiguo
end if

End Function		

Function PeriodoAcad()
Dim Sql
Dim Rst 
 
Sql = "select COALESCE(dbo.Fn_ValorParame('Periodo'),0) as Periodo "
If bcl_ado(Sql, rst) Then
	PeriodoAcad = rst("Periodo")
End If

End Function

Function AnoAcad()
Dim Sql
Dim Rst 

Sql = "select COALESCE(dbo.Fn_ValorParame('Ano'),0) as Ano " 'from mt_parame "
If bcl_ado(Sql, rst) Then
	AnoAcad = rst("Ano")
End If

End Function	

Function PeriodoAdmision()
Dim Sql
Dim Rst 
 
Sql = "select COALESCE(dbo.Fn_ValorParame('periodoadmision'),0) as periodoadmi "
If bcl_ado(Sql, rst) Then
	PeriodoAdmision = rst("periodoadmi")
End If

End Function

Function AnoAdmision()
Dim Sql
Dim Rst 

Sql = "select COALESCE(dbo.Fn_ValorParame('anoadmision'),0) as anoadmision " 'from mt_parame "
If bcl_ado(Sql, rst) Then
	AnoAdmision = rst("anoadmision")
End If

End Function		


Function EncuestasVigentes()
Dim StrSql
Dim Rst

  StrSql = "Select * From RA_ENCUESTA_EI Where getDate()>=FecIni And getDate()<=FecFin"

  Set Rst = Session("Conn").Execute(StrSql)

  'Rst.open StrSql 

  if not Rst.eof then
    EncuestasVigentes = 1
  else
    EncuestasVigentes = 0
  end if
  Rst.close()
End function

Function getNroEncuestaVigentes()
Dim StrSql
Dim Rst

  StrSql = "Select Nroencuesta From RA_ENCUESTA_EI Where getDate()>=FecIni And getDate()<=FecFin"

  Set Rst = Session("Conn").Execute(StrSql)

  'Rst.open StrSql 

  if not Rst.eof then
    EncuestasVigentes = Rst("Nroencuesta")
  else
    EncuestasVigentes = 0
  end if
  Rst.close()
End function

Function EstaPendiente(thisCodCli, Nro)
Dim StrSql
Dim Rst

  StrSql = "Select Count(CodCli) as hay From ra_encuestados_ei Where CodCli='" & thisCodCli & "' And NroEncuesta='" & Nro & "'"

  Set Rst = Session("Conn").Execute(StrSql)

  'Rst.open StrSql 

  if not Rst.eof then
    EstaPendiente = cInt(Rst("hay"))
  else
    EstaPendiente = 0
  end if
  Rst.close()
End function

Function Ceiling(intNumber)
Dim dblNumber
dblNumber = CDbl(intNumber)

If Int(dblNumber * 10) MOD 10 > 0 Then
Ceiling = Int(dblNumber) + 1
Else
Ceiling = Int(dblNumber)
End If
End Function

Public Sub Audita (opcion,transaccion)
dim StrSql 
dim username
	username=SacaUsuario
	
	StrSql ="INSERT INTO RA_AUDITA( USERNAME,CODCLI,OPCION,FECHA,TRANSACCION) " 
	StrSql =StrSql & "VALUES  ("& username &",'"& session("codcli") &"',"& opcion &",GETDATE(),'"& transaccion &"')"

Session("Conn").Execute(StrSql)
end Sub

Function SacaUsuario
dim str 
dim Rst

	str ="select id_usuario from dbo.ca_usuarios WHERE us_consuser='"& Session("usrNW") &"'"
	Set Rst = Session("Conn").Execute(str)
	if not Rst.eof then
		SacaUsuario = Rst("id_usuario")
	else
		SacaUsuario = 0
	end if
	Rst.close()
End Function

function ExisteArchivo(archivo)

	dim fs
	set fs=Server.CreateObject("Scripting.FileSystemObject")
	if fs.FileExists(archivo) then
	  ExisteArchivo = true  'response.write("File c:\asp\introduction.asp exists!")
	else
	  ExisteArchivo = false 'response.write("File c:\asp\introduction.asp does not exist!")
	end if
	set fs=nothing 
	
end function

Function Decode(sIn)

dim x, y, abfrom, abto
Decode="": ABFrom = ""
For x = 0 To 25: ABFrom = ABFrom & Chr(65 + x): Next
For x = 0 To 25: ABFrom = ABFrom & Chr(97 + x): Next
For x = 0 To 9: ABFrom = ABFrom & CStr(x): Next
abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
For x=1 to Len(sin): y=InStr(abto, Mid(sin, x, 1))
	If y = 0 then
		Decode = Decode & Mid(sin, x, 1)
	Else
		Decode = Decode & Mid(abfrom, y, 1)
	End If
Next

End Function

Function Encode(sIn)

dim x, y, abfrom, abto
Encode="": ABFrom = ""
For x = 0 To 25: ABFrom = ABFrom & Chr(65 + x): Next
For x = 0 To 25: ABFrom = ABFrom & Chr(97 + x): Next
For x = 0 To 9: ABFrom = ABFrom & CStr(x): Next
abto = Mid(abfrom, 14, Len(abfrom) - 13) & Left(abfrom, 13)
For x=1 to Len(sin): y = InStr(abfrom, Mid(sin, x, 1))
	If y = 0 Then
		Encode = Encode & Mid(sin, x, 1)
	Else
		Encode = Encode & Mid(abto, y, 1)
	End If
Next

End Function 

Function MuestraOpcion(Permiso) 

if Permiso <> "S" then 
	MuestraOpcion = "style='display:none'"
else
	MuestraOpcion = ""
end if 

End Function 

Function encripta_portal(valor)
dim str 
dim Rst

	str ="SELECT dbo.fn_encripta_portal('"& valor &"')resultado"
	Set Rst = Session("Conn").Execute(str)
	if not Rst.eof then
		encripta_portal = Rst("resultado")
	else
		encripta_portal = ""
	end if
	Rst.close()

End Function 

Function desencripta_portal(valor)
dim str 
dim Rst

	str ="SELECT dbo.fn_desencripta_portal('"& valor &"')resultado"
	Set Rst = Session("Conn").Execute(str)
	if not Rst.eof then
		desencripta_portal = Rst("resultado")
	else
		desencripta_portal = ""
	end if
	Rst.close()

End Function 



Function GetPaginaLink(id)
dim str 
dim Rst

	str ="select link_url from fw_links WHERE id='"& id &"'"
	Set Rst = Session("Conn").Execute(str)
	if not Rst.eof then
		GetPaginaLink = Rst("link_url") 
	else
		GetPaginaLink = ""
	end if
	Rst.close()

End Function 

Function GetEsArchivoPaginaLink(id)
dim str 
dim Rst

	str ="select coalesce(esArchivo,'NO')esArchivo from fw_links WHERE id='"& id &"'"
	Set Rst = Session("Conn").Execute(str)
	if not Rst.eof then
		GetEsArchivoPaginaLink = Rst("esArchivo") 
	else
		GetEsArchivoPaginaLink = ""
	end if
	Rst.close()

End Function 

Function GetPaginaCertifOnline()
dim str 
dim Rst

	str="SELECT coalesce(dbo.Fn_ValorParame('PaginaCertificadosOnline'),'')Parame"
	set Rst= Session("Conn").Execute(str)
	
	if not Rst.eof then
		GetPaginaCertifOnline=Rst("Parame")
	else
		GetPaginaCertifOnline=""
	end if

	Rst.close()

End Function 

Function LimpiarTexto(ByVal texto)
    
    Dim objRegExp
    Set objRegExp = New Regexp
    
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    
    objRegExp.Pattern = "\s+"
    texto = objRegExp.Replace(texto, " ")
    
    objRegExp.Pattern = "[()?*"",\\<>&#~%{}+:\/!;']+"
    texto = objRegExp.Replace(texto, "")
    
    Dim i, s1, s2
    s1 = "¡¿…»Õœ”“⁄‹·‡ËÈÌÔÛÚ˙¸ÒÁ"
    s2 = "AAEEIIOOUUaaeeiioouunc-"
    If Len(texto) <> 0 Then
        For i = 1 To Len(s1)
            texto = Replace(texto, Mid(s1,i,1), Mid(s2,i,1))
        Next
    End If


	sql="SELECT dbo.fnQuitarCaracteresRaros('"+ texto +"')x"	
	set rTQ=Session("Conn").Execute (sql)
	if not rTQ.EOF then
		texto= rTQ(0)
	end if

    LimpiarTexto = LCase(texto)
End Function

function ObtienekeyMoodleUCV() 

str ="sp_muestra_interfaz_usuario '"& Session("usrNW") &"'"

set rst= Session("Conn").Execute(str)		
if not rst.eof then
	user=rst("usuario_id")
	pass=rst("usuarios_password")
else
	user=""
	pass=""
end if

envio ="http://virtual.ucevalpo.cl/webservice/rest/server.php?wstoken=27f881516b9cd0dcc298b6dbf1d1103d&username="& user &"&password="& pass &"&wsfunction=local_ssostandar_peticion"
 
set req = Server.CreateObject("MSXML2.ServerXMLHTTP")
            req.open "GET", envio, false
            req.send ""
            dim myJSON 
            			
			Set myXmlDoc = Server.CreateObject("MSXML2.DOMDocument")
            myXmlDoc.loadXML (req.responseXml.xml)
			
dim estado, key			
			
For Each oNode In myXmlDoc.SelectNodes("/RESPONSE/SINGLE/KEY")
  sKey = oNode.GetAttribute("name")
  sValue = oNode.Text
	
  Select Case sKey
    Case "estado"
		estado = sValue
    Case "key"
    	key = sValue 
  End Select
  
Next

ObtienekeyMoodleUCV = key
	
end function

function ObtieneEstadoMoodleUCV()

str ="sp_muestra_interfaz_usuario '"& Session("usrNW") &"'"

set rst= Session("Conn").Execute(str)		
if not rst.eof then
	user=rst("usuario_id")
	pass=rst("usuarios_password")
else
	user=""
	pass=""
end if

envio ="http://virtual.ucevalpo.cl/webservice/rest/server.php?wstoken=27f881516b9cd0dcc298b6dbf1d1103d&username="& user &"&password="& pass &"&wsfunction=local_ssostandar_peticion" 
 
set req = Server.CreateObject("MSXML2.ServerXMLHTTP")
            req.open "GET", envio, false
            req.send ""
            dim myJSON 
            			
			Set myXmlDoc = Server.CreateObject("MSXML2.DOMDocument")
            myXmlDoc.loadXML (req.responseXml.xml)
			
dim estado, key			
			
For Each oNode In myXmlDoc.SelectNodes("/RESPONSE/SINGLE/KEY") 
  sKey = oNode.GetAttribute("name")
  sValue = oNode.Text
	
  Select Case sKey
    Case "estado"
		estado = sValue
    Case "key"
    	key = sValue 
  End Select
  
Next

ObtieneEstadoMoodleUCV = estado
	
end function

function ObjetosLocalizacion(URL)
Dim StrSql
Dim Rst
dim i 

  i = 0 
  StrSql = "sp_trae_localizacion '" & URL & "',NULL"
  set Rst = Session("Conn").execute(StrSql)

  do while not Rst.eof
  
  	if i = 0 then
		response.write("<script language='javascript' type='text/javascript'>"& Chr(10))	
		i = 1
	end if 
	
	if Rst("INDICE_OBJ") = 1 then
		response.write("var elementos = document.getElementsByName('" & Rst("Nombre_Objeto") & "');"& Chr(10))
		response.write("for (x=0;x<elementos.length;x++)"& Chr(10))
		response.write("document.getElementsByName('" & Rst("Nombre_Objeto") & "').item(x).innerHTML='" & Rst("valor") & "';"& Chr(10))
	else
		response.write("document.getElementById('" & Rst("Nombre_Objeto") & "').innerHTML='" & Rst("valor") & "';"& Chr(10))
	end if 
	
    Rst.movenext
	
	if Rst.eof = true then
		response.write("</script>")
	end if 
	
  loop
	
end function
%>