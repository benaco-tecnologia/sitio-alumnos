<%
Function BuscaProximoCodigo()
	Dim strCodigo
	Dim rstCodigo

	If RDBMS = "ORACLE" Then
	    strCodigo = "select idunico.nextval codigo from dual"
	    If BuscaCodigoLectura(strCodigo, rstCodigo) Then
      	  BuscaProximoCodigo = ValNulo(rstCodigo("codigo"), STR_)
	    Else
      	  BuscaProximoCodigo = "Fallo"
	    End If
    	ElseIf RDBMS = "SQLSERVER" Then
	    BuscaProximoCodigo = BuscaCorrel("idunico")
	Else
	   'Msgbox "Debe Indicar La Base De Datos A Utilizar", VbCritical
	   'Unload Me
	End If
End Function

Function BuscaCorrel(Seq)
	Dim strCorrel
	Dim strSession
	Dim rstSession
	Dim rstCorrel

	strSession = "SET TRANSACTION ISOLATION LEVEL SERIALIZABLE"

	con.Execute strSession
	strSession = "BEGIN TRANSACTION"
	Conn.Execute strSession

	strCorrel = "select max(" & Seq & ") " & Seq & " from " & Seq & " (TABLOCKX)"
	If BuscaCodigoEscritura(strCorrel, rstCorrel) Then
	    BuscaCorrel = rstCorrel.rdoColumns(0)
	    rstCorrel.Edit
	    rstCorrel.rdoColumns(0) = Val(rstCorrel.rdoColumns(0) + 1)
	    rstCorrel.Update
	    rstCorrel.Close
	Else
	    MsgBox "No Existe La Secuencia", vbCritical
	    BuscaCorrel = " "
	End If

	strSession = "COMMIT"
	Conn.Execute strSession
End Function

Function Insertar(var_alumno,var_pestud,Codcarr)
Dim strPlan, StrNota
Dim rstPlan
Dim rdoIdunico
Dim StrCurric
Dim rstCurric
Dim strResol
Dim rstResol
Dim v_nivel
Dim Ramo
Dim Cred, Duracion

    v_idunico = ValNulo(BuscaProximoCodigo, NUM_)
    If primera Then
        v_idunico_Ini = v_idunico
        primera = False
    End If
    v_idunico_Fin = v_idunico

    PromedioGeneral = 0
    CalculoPromFinal = True
    CuentaRamos = 0
    Suma = 0
    NivelAnterior = 0
    SumaCred = 0
    SumaCredAnt = 0

    StrCurric = "Select Regimen from mt_carrer where CodCarr = '" & Trim(Codcarr) & "' "

    Set RstCurric = Conn.Execute(strcurric)

    If not rstcurric.eof Then
	if (ValNulo(rstCurric("Regimen"), STR_ )= "ANUAL") then
  	      Duracion = 2
        else
	      Duracion = 1
		end if
    Else
       Duracion = 1
    End If

    rstCurric.Close

    StrCurric = " select ra_curric.codramo, ra_curric.Ramoequiv, ra_curric.nivel, ra_curric.Orden,ra_ramo.nombre" _
    & " from ra_curric,ra_ramo " _
    & " where ra_curric.codramo = ra_ramo.codramo" _
    & " and ra_curric.codpestud = '" & Trim(var_pestud) & "'" _
    & " order by convert(int,nivel) asc,orden asc"

    Set RstCurric = Conn.Execute(strcurric)

    If not rstcurric.eof Then
      While Not rstCurric.EOF
        'Response.write(rstCurric("CodRamo") & "<br>")
        AnalizaRamo var_alumno, rstCurric("Codramo"), rstCurric("RamoEquiv"), rstCurric("Nivel"), rstCurric("Nombre"), ValNulo(rstCurric("Orden"), NUM_), Duracion
        rstcurric.MoveNext
        'DoEvents
      Wend
    End If
  
    'Response.end
    StrCurric = " select distinct n.codramo, r.nombre " _
    & " from ra_nota n, ra_ramo r " _
    & " where n.codramo = r.codramo" _
    & "   and n.codcli = '" & var_alumno & "' " _
    & "   and n.codramo not in " _
    & "    (select c.codramo from ra_curric c " _
    & "      Where c.Codpestud = '" & var_pestud & "' ) "
'    & "      Where Codpestud = '" & Trim(var_pestud) & "' ) "
'    & " and n.codramo in (select distinct b.codramo from ra_curric b)" _

    Set RstCurric = Conn.Execute(strcurric)

    If not rstcurric.eof Then
        Do While Not rstCurric.EOF
            AnalizaRamo var_alumno, rstCurric("Codramo"), rstCurric("Codramo"), 99, rstCurric("Nombre"), 1, Duracion
        rstCurric.MoveNext
        Loop
    End If

    strResol = "select * from ra_resolu where codcli = '" & Trim(var_alumno) & "'"

    Set rstResol = Conn.Execute(strResol)

    If not rstResol.eof Then
        StrNota = "insert into liscarhis  (rut,idunico,nivel,codramo,institucion) values ('" & Trim(var_alumno) & _
        "','" & v_idunico & "',79,'" & Trim(Ramo) & "','" & Institucion & "'," &  _
        Cred & ")"
        Conn.Execute StrNota		
        StrNota = "insert into liscarhis  (rut,idunico,nivel,curnocur,codramo,institucion) values ('" & _
                   Trim(var_alumno) & "','" & v_idunico & "',80,'RESOLUCIONES OTORGADAS','" & _
                   Trim(Ramo) & "','" & Institucion & "'," & Cred & ")"
        Conn.Execute StrNota
        v_nivel = 81
        While Not rstResol.EOF
            StrNota = " insert into liscarhis (rut,oport_2,curnocur,oport_1,idunico,nivel,codramo,institucion) values" _
                    & " ('" & Trim(var_alumno) & "','" & rstResol("codres") & "','" & rstResol("tipo") & "','" & rstResol("fecha") & "','" & v_idunico & _
                    "'," & v_nivel &    ",'" & Trim(Ramo) & "','" & Institucion & "'," & Cred & ")"
            v_nivel = v_nivel + 1
            Conn.Execute StrNota
            rstResol.MoveNext
        Wend
    End If

    'If CalculoPromFinal = True And SumaGeneral > 0 Then
    '    PromedioGeneral = SumaGeneral / ContadorGeneral
    'End If

    'Conn.CommitTrans
    rstCurric.Close
    'Rstnota.Close
End Function


Function AnalizaRamo(var_alumno,Codramo,CodRamoReal,Nivel,NombreRamo,Orden,Duracion)
Dim Var1, Var2, Var3, Var4
Dim Var5
Dim RamoReal
Dim RamoMalla
Dim Cred
Dim Obser
Dim StrNota
Dim Rstnota, Rst2
Dim contador
Dim CantidadRamosNivel
Dim Strsalida
Dim Ramo
Dim AnoR
Dim Promedio
Dim SumaGeneral
Dim ContadorGeneral

    ContadorGeneral = 0: SumaGeneral = 0
    CalculaPromedio = True

    If Duracion > 1 Then
       Nivel = Int((Nivel - 1) / Duracion) * Duracion + 1
    End If

    Sw = True: Var1 = "": Var2 = "": Var3 = "": Var4 = "": Var5 = ""
    Nota = 0: Obser = ""
    'Do Events
    If NivelAnterior = 0 Then NivelAnterior = Nivel

    If NivelAnterior <> Nivel Then
            'Suma = 0
       CuentaRamos = 0
            'SumaCred = 0
    End If

    Cred = SacaCreditosRamo(Trim(Codramo))
	'-- el campo concepto no esta en la tabla ra_nota no supe que hacer si sacarla


    StrNota = " select nf,estado,ano,periodo,ramoequiv,codramo,codsecc,concepto "
'    StrNota = " select nf,estado,ano,periodo,ramoequiv,codramo,codsecc"
    StrNota = StrNota & " FROM ra_nota " _
    & " where codcli = '" & Trim(var_alumno) & "' and " _
    & " codramo = '" & Codramo & "'" _
    & " order by ano asc, periodo asc"


	Set Rstnota = Conn.execute(StrNota)

	if not Rstnota.eof then

            RamoMalla = ValNulo(Rstnota("codramo"), STR_)
            RamoReal = ValNulo(Rstnota("ramoequiv"), STR_)

            contador = 1
            While Not Rstnota.EOF
                'DoEvents
                If Trim(RamoReal) <> Trim(ValNulo(Rstnota("ramoequiv"), STR_)) Then
                    RamoReal = ValNulo(Rstnota("ramoequiv"), STR_)
                End If
                If contador <= 5 Then 'Cuenta las veces que un alumno a hecho una ramo
                    If contador = 1 Then
                       'Var1 = Trim(ValNulo(Rstnota("codsecc"), STR_))) & " " & ValNulo(Rstnota("ano"), STR_) & "/" & ValNulo(Rstnota("periodo"), STR_) & " " & ValNulo(Rstnota("nf"), NUM_) & " " & RstNota("Concepto") & " " & ValNulo(Rstnota("estado"), STR_)
                       Var1 = Trim(ValNulo(Rstnota("codsecc"), STR_)) & " " & ValNulo(Rstnota("ano"), STR_) & "/" & ValNulo(Rstnota("periodo"), STR_) & " " & ValNulo(Rstnota("nf"), NUM_) & " " & " " & ValNulo(Rstnota("estado"), STR_)
                       
                    ElseIf contador = 2 Then
                           'Var2 = Trim(ValNulo(Rstnota("codsecc"), STR_)) & " " & ValNulo(Rstnota("ano"), STR_) & "/" & ValNulo(Rstnota("periodo"), STR_) & " " & ValNulo(Rstnota("nf"), NUM_), & " " & ValNulo(Rstnota("estado"), STR_)
                           Var2 = Trim(ValNulo(Rstnota("codsecc"), STR_)) & " " & ValNulo(Rstnota("ano"), STR_) & "/" & ValNulo(Rstnota("periodo"), STR_) & " " & ValNulo(Rstnota("nf"), NUM_) & " " &  ValNulo(Rstnota("concepto"), STR_) & " " &  ValNulo(Rstnota("estado"), STR_)
                    ElseIf contador = 3 Then
    'Var3 = Trim(ValNulo(Rstnota("codsecc"), STR_)) & " " & ValNulo(Rstnota("ano"), STR_) & "/" & ValNulo(Rstnota("periodo"), STR_) & " " & Format(ValNulo(Rstnota("nf"), NUM_), "0.0") & " " & ValNulo(Rstnota("estado"), STR_)
	
                           Var3 = Trim(ValNulo(Rstnota("codsecc"), STR_)) & " " & ValNulo(Rstnota("ano"), STR_) & "/" & ValNulo(Rstnota("periodo"), STR_) & " " & ValNulo(Rstnota("nf"), NUM_) & " " & ValNulo(Rstnota("concepto"), STR_) & " " & ValNulo(Rstnota("estado"), STR_)
                    ElseIf contador = 4 Then
                           'Var4 = Trim(ValNulo(Rstnota("codsecc"), STR_)) & " " & ValNulo(Rstnota("ano"), STR_) & "/" & ValNulo(Rstnota("periodo"), STR_) & " " & ValNulo(Rstnota("nf"), NUM_) & " " & ValNulo(Rstnota("estado"), STR_)
                           Var4 = Trim(ValNulo(Rstnota("codsecc"), STR_)) & " " & ValNulo(Rstnota("ano"), STR_) & "/" & ValNulo(Rstnota("periodo"), STR_) & " " & ValNulo(Rstnota("nf"), NUM_) & " " & ValNulo(Rstnota("concepto"), STR_)   & " " & ValNulo(Rstnota("estado"), STR_)
                    ElseIf contador = 5 Then
                           'Var5 = Trim(ValNulo(Rstnota("codsecc"), STR_)) & " " & ValNulo(Rstnota("ano"), STR_) & "/" & ValNulo(Rstnota("periodo"), STR_) & " " & ValNulo(Rstnota("nf"), NUM_) & " " & ValNulo(Rstnota("estado"), STR_)
                           Var5 = Trim(ValNulo(Rstnota("codsecc"), STR_)) & " " & ValNulo(Rstnota("ano"), STR_) & "/" & ValNulo(Rstnota("periodo"), STR_) & " " & ValNulo(Rstnota("nf"), NUM_) & " " &  ValNulo(Rstnota("concepto"), STR_) & " " & ValNulo(Rstnota("estado"), STR_)
                    End If
                End If

                Nota = ValNulo(Rstnota("nf"), NUM_)
                If Trim(ValNulo(Rstnota("estado"), STR_)) = "" Then
                    Sw = False ' Si no tiene estado

            'Habiliar Promedios Inicio
                    CalculaPromedio = False
                    'Habiliar Promedios Fin
                ElseIf Trim(ValNulo(Rstnota("estado"), STR_)) = "R" Then
                    Sw = True 'Si  tiene estado o esta reprobado
                    CuentaRamos = CuentaRamos + 1
                    Suma = Suma + (Nota * Cred)
                    SumaCred = SumaCred + Cred
                    CalculaPromedio = True
                ElseIf Trim(ValNulo(Rstnota("estado"), STR_)) = "A" Then
                    Sw = True 'Si  tiene estado o esta reprobado
                    CuentaRamos = CuentaRamos + 1
                    Suma = Suma + (Nota * Cred)
                    SumaCred = SumaCred + Cred

                    CalculaPromedio = True
                ElseIf Trim(ValNulo(Rstnota("estado"), STR_)) = "I" Or Trim(ValNulo(Rstnota("estado"), STR_)) = "E" Then
                    Sw = True 'Si  tiene estado o esta reprobado
                    CuentaRamos = CuentaRamos + 1
                End If

                Rstnota.MoveNext
                contador = contador + 1
            Wend

            If contador > 1 And Sw = True Then 'Si aprobo el ramo en una segunda instancia
                CalculaPromedio = True
                CalculoPromFinal = True
            End If
        Else
            RamoReal = ""
            RamoMalla = ""
            Nota = 0
            Sw = False
        End If

        Strsalida = Codramo & " " & _
                    NombreRamo
        Ramo = Codramo
        AnoR = Mid(Var1, 1, 4)
        Obser = ""
        'If RamoMalla = "BEC212" Then MsgBox "Lo encontre"
        If Trim(RamoMalla) <> Trim(RamoReal) And Trim(RamoReal) <> "" Then
		    'response.write(ramoreal)
			'response.end
            Obser = Trim(RamoReal) & " " & Trim(BuscaDescripcion("Select nombre from ra_ramo where codramo='" & Trim(RamoReal) & "'"))
            'Obser = ObserRamo(Trim(Ramo), AnoR, Trim(var_alumno))
            'Obser = Trim(RamoReal)
			
        End If

 If NivelAnterior <> Nivel Then
            CantidadRamosNivel = SacaCantidadRamos(NivelAnterior, Nivel)
            If CuentaRamosAnt >= CantidadRamosNivel And CalculaPromedioAnt = True And SwAnt = True Then
                Promedio = SumaAnt / IIf(SumaCredAnt = 0, 1, SumaCredAnt)
                If Aproximacion = "N" Then
                    Promedio = Mid(Promedio, 1, 4)
                Else
                    Promedio = Format(Promedio, "0.0")
                End If

                'Buscaremos donde insertarlo
                StrNota = "Select max(orden) as orden from liscarhis"
                StrNota = StrNota & " Where rut = '" & var_alumno & "' "
                StrNota = StrNota & "     and Nivel=" & NivelAnterior & ""

                If BCL_ADO(StrNota, Rst2) Then
                   OrdenAnt = ValNulo(Rst2("Orden"), NUM_)
                End If
                If OrdenAnt <> 0 Then
                    StrNota = "Update liscarhis set promedio='" & Trim(Promedio) & "'" _
                    & " where rut='" & var_alumno & "'" _
                    & "     and orden=" & OrdenAnt & "" _
                    & "     and Nivel=" & NivelAnterior & "" _
                    & "     and idunico='" & v_idunico & "'"
                    '& "     and codramo='" & Trim(RamoAnt) & "'"
                    Conn.Execute StrNota, rdExecDirect
                End If
                SumaGeneral = SumaGeneral + Promedio
                ContadorGeneral = ContadorGeneral + 1

                StrNota = "insert into liscarhis " _
                & " values ('" & Trim(var_alumno) & "','" & Strsalida & "'," _
                & " '" & Var1 & "','" & Var2 & "','" & Var3 & "','" & Var4 & "'," _
                & " '" & Var5 & " ','" & v_idunico & "'," & Val(Nivel) & "," _
                & " '','" & Trim(Ramo) & "'," _
                & " '" & Institucion & "'," & Cred & ",'" & Obser & "'," & Orden & ")"
            Else
                StrNota = "insert into liscarhis" _
                & " values ('" & Trim(var_alumno) & "','" & Strsalida & "'," _
                & " '" & Var1 & "','" & Var2 & "','" & Var3 & "','" & Var4 & "'," _
                & " '" & Var5 & " ','" & v_idunico & "'," & Val(Nivel) & "," _
                & " '','" & Trim(Ramo) & "'," _
                & " '" & Institucion & "'," & Cred & ",'" & Obser & "'," & Orden & ")"
            End If

        Else
            if not IsNumeric(Nivel) then
              Nivel = 0
            end if
            CantidadRamosNivel = SacaCantidadRamos(NivelAnterior, Nivel, CodCarr, CodPestud)
            StrNota = "insert into liscarhis (rut ,curnocur,oport_1 ,oport_2,oport_3,oport_4,oport_5, idunico, nivel, codramo,institucion,creditoramo,obser,orden)"_ 
            & " values ('" & Trim(var_alumno) & "','" & trim(Strsalida) & "'," _
            & " '" & trim(Var1) & "','" & trim(Var2) & "','" & trim(Var3) & "','" & trim(Var4) &  "'," _
            & " '" & trim(Var5) & " ','" & trim(v_idunico) & "'," & trim(Nivel) & "," _
            & " '" & Trim(Ramo) & "','" & trim(Institucion) & "'," & trim(Cred) & ",'" & trim(Obser) & "'," & trim(Orden) & ")"
			'response.write(strnota)
			'response.end
        End If
        NivelAnterior = Nivel
        StrsalidaAnt = Strsalida
        Var1Ant = Var1
        Var2Ant = Var2
        Var3Ant = Var3
        Var4Ant = Var4
        Var5ant = Var5
        PromedioAnt = Promedio
        CredAnt = Cred
        ObserAnt = Obser
        OrdenAnt = Orden
        RamoAnt = Ramo
        SumaAnt = Suma
        SumaCredAnt = SumaCred
        CalculaPromedioAnt = CalculaPromedio
        CuentaRamosAnt = CuentaRamos
        SwAnt = Sw
        'Response.write(StrNota)
        'Response.End
        conn.Execute StrNota
End Function

Function SacaCreditosRamo(Ramo)
Dim str
Dim rst
    str = "select credito from ra_ramo where codramo='" & Trim(Ramo) & "' "
    Set Rst = Conn.execute(Str)

    if not Rst.eof then
        SacaCreditosRamo = ValNulo(rst("credito"), NUM_)
    End If
End Function
Private Function SacaCantidadRamos(Nivel, NivelPos, CodCarr, CodPestud)
Dim str 
Dim rst 
     str = " select count(codramo) from ra_curric,ra_pestud" _
       & " where ra_curric.codpestud=ra_pestud.codpestud" _
       & " and ra_pestud.codcarr = '" & CodCarr & "'" _
       & " and ra_curric.nivel >= " & Nivel & "" _
       & " and ra_curric.nivel < " & NivelPos & "" _
       & " and ra_curric.codpestud = '" & CodPestud & "' "
    If BCL_ADO(str, rst) Then
        SacaCantidadRamos = ValNulo(rst(0), NUM_)
    End If
End Function

Public Function BuscaDescripcion(str)
Dim rst 
  If BCL_ADO(str, rst) Then
    BuscaDescripcion = rst(0)
  Else
    BuscaDescripcion = ""
  End If
  rst.Close
End Function



%>
