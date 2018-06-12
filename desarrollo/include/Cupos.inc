<%
dim TipoCupo
Function EsRepitente(Codramo, Codcli) 
Dim StrSql 
Dim Rst 

   StrSql = "Select estado from ra_nota" _
		  & " where (codramo='" & Trim(Codramo) & "' or ramoequiv='" & Trim(Codramo) & "')" _
          & " and codcli='" & Trim(Codcli) & "'" _
          & " and estado='R'"
   'response.write(strsql)
   If BCL_ADO(strSql, Rst) Then       
	     EsRepitente = True
   Else
         EsRepitente = False
   End If
End Function


Function SacaCupoAntiguo(Ramo, Seccion, Ano, Periodo, CodCarr, CodSede, EsRepit, TipoCupo, RamoReal)
Dim str, Strsql 
Dim Rst, Rst2, Rst3, RstEquiv
Dim cupo
Dim Reserva
Dim lc_CodCarrSecc 
Dim CupoOtrasCarreras 

	CupoOtrasCarreras = 0
    session("RamoObliPlanificadoOp") = "N"
    TipoCupo = "CT"	
	'Si el ramo está planificado en su carrera
	
    StrSql = "Select cupo, cuponuevo,cupototal,reserva from ra_seccio where " _
    & " codramo = '" & Trim(Ramo) & "'" _
    & " and codsecc='" & Trim(Seccion) & "'" _
    & " and ano=" & Ano & "" _
    & " and periodo=" & Periodo & "" _
    & " and codSede ='" & Trim(CodSede) & "'" _
    & " and Codcarr ='" & Trim(Codcarr) & "'"


    If BCL_ADO(StrSql, Rst) Then
	
        If ValNulo(Rst("cupo"), NUM_) = 0 And ValNulo(Rst("cuponuevo"), NUM_) = 0 Then
            cupo = ValNulo(rst("cupototal"), NUM_)
            TipoCupo = "CT"
        Else
            If EsRepit Then
                cupo = ValNulo(rst("cupo"), NUM_) 'CupoRepitente
                TipoCupo = "CR"
            Else
                cupo = ValNulo(rst("cuponuevo"), NUM_) 'CupoNoRepitente
                TipoCupo = "CN"
            End If
        End If
        
		Reserva = ValNulo(rst("reserva"), NUM_) 
		
        ' sacaremos el total de cupos asignados
        Strsql = "Select sum(cupo) as cupo from ra_optivo"
        Strsql = Strsql & " Where CodSede = '" & CodSede & "' "
        Strsql = Strsql & " and CodRamo = '" & Trim(Ramo) & "' "
        Strsql = Strsql & " and CodSecc = '" & Trim(Seccion) & "' "
        Strsql = Strsql & " and ano = " & Ano & " "
        Strsql = Strsql & " and Periodo = " & Periodo & " "

        If BCL_ADO(Strsql, Rst2) Then
			If Reserva = 0 Then
	            SacaCupoAntiguo = cupo - ValNulo(Rst2("cupo"), NUM_)
            Else
                SacaCupoAntiguo = Reserva
				TipoCupo = ""
            End If
           Rst2.close()
		else
			SacaCupoAntiguo = cupo
        End If
        Rst.close()
    Else

        StrSql = "Select RamoEquiv,CodCarrequiv from ra_equiv "
        strSql = strSql & " WHERE CodRamo = '" & RamoReal & "' "
        strSql = strSql & " AND CodCarr = '" & CodCarr & "' "
        If BCL_ADO(StrSql, Rstequiv) Then
           Do While Not rstequiv.EOF
		   	cupo = SacaCupoAntiguo(ValNulo(RstEquiv("RamoEquiv"), STR_), Seccion, Ano, Periodo, RstEquiv("CodCarrequiv"),codsede, EsRepit, TipoCupo,ValNulo(RstEquiv("RamoEquiv"), STR_))
      
             If cupo <> 0 Then
               Exit Do
             End If
             rstequiv.MoveNext
           Loop
		   
		   if Cupo = 0 then
		   		' Si no está planificado en la carrera y si no tiene equivalencias de malla, 
				' entonces ve si está planificado en otras carreras
				' aquí la variable ramo corresponde al ramo real que es dictado
				' Solo en este else se debe ocupar el ramoreal 
	
				str = "Select codcarrsec,sum(cupo) as cupo from ra_optivo where" 
				if session("OtrosRamos") = "N" then
					str = str & " codopt='" & Trim(Ramoreal) & "'" 
					str = str & " and codramo='" & Trim(Ramo) & "'"
				else
					str = str & " codopt=''" 
					str = str & " and codramo='" & Trim(Ramoreal) & "'"
				end if			 
				str = str & " and codsecc='" & Trim(Seccion) & "'" 
				str = str & " and ano=" & Ano & "" 
				str = str & " and periodo=" & Periodo & "" 
				str = str & " and codcarr='" & Trim(CodCarr) & "'" 
				str = str & " and codsede='" & Trim(codsede) & "' group by codcarrsec"
	
				'Marcamos está variable para diferenciar que es un ramo obligatorio, pero planificado en otra carrera
				session("RamoObliPlanificadoOp") = "S"
				if Ramo="EDUC0007" then
					'response.write(str) & "</br>"
				end if		
				If BCL_ADO(str, Rst) Then 'Si hay cupos asignados para mi carrera
					lc_CodCarrSecc = valnulo(Rst("codcarrsec"),str_)
					If Valnulo(Rst("Cupo"),num_) > 0 Then 'Si los cupos son administrados por carrera
						Rst.Close
						str = "Select cupo from ra_optivo where" 
						if session("OtrosRamos") = "N" then
							str = str & " codopt='" & Trim(Ramoreal) & "'" 
							str = str & " and codramo='" & Trim(Ramo) & "'"
						else
							str = str & " codopt=''" 
							str = str & " and codramo='" & Trim(Ramoreal) & "'"
						end if			
						str = str & " and codsecc='" & Trim(Seccion) & "'" 
						str = str & " and ano=" & Ano & "" 
						str = str & " and periodo=" & Periodo & "" 
						str = str & " and codcarr='" & Trim(CodCarr) & "'" 
						str = str & "  and codsede='" & Trim(codsede) & "'"
						If BCL_ADO(str, Rst) Then
							SacaCupoAntiguo = ValNulo(Rst("cupo"), NUM_)
							TipoCupo = "CT"
							Rst.Close
						End If
						session("CupoCero") = "N"	
					Else 'Si los cupos son administrados como cupo total
						session("CupoCero") = "S"
						Rst.Close
	
						'Esto está mal, se debe corregir, los campos no corresponden a la tabla
						str = "Select cupo,cupototal,cuponuevo,reserva from ra_seccio where" _
						& " codsecc='" & Trim(Seccion) & "'" _
						& " and codramo='" & Trim(Ramo) & "'" _
						& " and ano=" & Ano & "" _
						& " and periodo=" & Periodo & "" _
						& " and codcarr='" & Trim(lc_CodCarrSecc) & "'" _
						& " and codsede='" & Trim(codsede) & "'"
						if Ramo="COMP0003" then
			'				response.write(str) & "</br>"
						end if
						If BCL_ADO(str, Rst) Then
							If ValNulo(Rst("cupo"), NUM_) = 0 And ValNulo(Rst("cuponuevo"), NUM_) = 0 Then
								cupo = ValNulo(Rst("cupototal"), NUM_)
								TipoCupo = "CT"
							Else
								If session("EsRepit") Then
									cupo = ValNulo(Rst("cupo"), NUM_) 'CupoRepitente
									TipoCupo = "CR"
								Else
									cupo = ValNulo(Rst("cuponuevo"), NUM_) 'CupoNoRepitente
									TipoCupo = "CN"
								End If
							End If
							Strsql = "Select sum(cupo) as cupo from ra_optivo"
							Strsql = Strsql & " Where CodSede = '" & CodSede & "' "
							Strsql = Strsql & " and CodRamo = '" & Trim(Ramo) & "' "
							Strsql = Strsql & " and CodSecc = '" & Trim(Seccion) & "' "
							Strsql = Strsql & " and ano = " & Ano & " "
							Strsql = Strsql & " and Periodo = " & Periodo & " "
							Strsql = Strsql & " and codcarrsec='" & Trim(lc_CodCarrSecc) & "'"
							If BCL_ADO(Strsql, Rst3) Then
								CupoOtrasCarreras = ValNulo(Rst3("cupo"), NUM_)
							end if
							if Ramo="COMP0003" then
								'response.write(Strsql)   
							end if
							SacaCupoAntiguo = cupo - (ValNulo(Rst("reserva"), NUM_) + CupoOtrasCarreras)
							'response.write(SacaCupoAntiguo)
							'response.end
							Rst.Close
						End If
					End If
					'rst.Close
				else
					SacaCupoAntiguo = 0
				End If		
		   else
				SacaCupoAntiguo = cupo
		   End If		
           Rstequiv.close()
        Else 	
		
			' Si no está planificado en la carrera y si no tiene equivalencias de malla, 
			' entonces ve si está planificado en otras carreras
			' aquí la variable ramo corresponde al ramo real que es dictado
			' Solo en este else se debe ocupar el ramoreal 

			str = "Select codcarrsec,sum(cupo) as cupo from ra_optivo where" 
			if session("OtrosRamos") = "N" then
				str = str & " codopt='" & Trim(Ramoreal) & "'" 
				str = str & " and codramo='" & Trim(Ramo) & "'"
			else
				str = str & " codopt=''" 
				str = str & " and codramo='" & Trim(Ramoreal) & "'"
			end if			 
			str = str & " and codsecc='" & Trim(Seccion) & "'" 
			str = str & " and ano=" & Ano & "" 
			str = str & " and periodo=" & Periodo & "" 
			str = str & " and codcarr='" & Trim(CodCarr) & "'" 
			str = str & " and codsede='" & Trim(codsede) & "' group by codcarrsec"

			'Marcamos está variable para diferenciar que es un ramo obligatorio, pero planificado en otra carrera
		    session("RamoObliPlanificadoOp") = "S"
			if Ramo="EDUC0007" then
				'response.write(str) & "</br>"
			end if		
			If BCL_ADO(str, Rst) Then 'Si hay cupos asignados para mi carrera
				lc_CodCarrSecc = valnulo(Rst("codcarrsec"),str_)
				If Valnulo(Rst("Cupo"),num_) > 0 Then 'Si los cupos son administrados por carrera
					Rst.Close
					str = "Select cupo from ra_optivo where" 
					if session("OtrosRamos") = "N" then
						str = str & " codopt='" & Trim(Ramoreal) & "'" 
						str = str & " and codramo='" & Trim(Ramo) & "'"
					else
						str = str & " codopt=''" 
						str = str & " and codramo='" & Trim(Ramoreal) & "'"
					end if			
					str = str & " and codsecc='" & Trim(Seccion) & "'" 
					str = str & " and ano=" & Ano & "" 
					str = str & " and periodo=" & Periodo & "" 
					str = str & " and codcarr='" & Trim(CodCarr) & "'" 
					str = str & "  and codsede='" & Trim(codsede) & "'"
					If BCL_ADO(str, Rst) Then
						SacaCupoAntiguo = ValNulo(Rst("cupo"), NUM_)
						TipoCupo = "CT"
						Rst.Close
					End If
					'Falta ver que valor se deja en está variable o en que variable "CupoCero = "N"", 
   					'para determinar la cantidad de alumnos inscritos porteriormente
                    session("CupoCero") = "N"

				Else 'Si los cupos son administrados como cupo total
                    session("CupoCero") = "S"
					Rst.Close

					'Esto está mal, se debe corregir, los campos no corresponden a la tabla
					str = "Select cupo,cupototal,cuponuevo,reserva from ra_seccio where" _
					& " codsecc='" & Trim(Seccion) & "'" _
					& " and codramo='" & Trim(Ramo) & "'" _
					& " and ano=" & Ano & "" _
					& " and periodo=" & Periodo & "" _
					& " and codcarr='" & Trim(lc_CodCarrSecc) & "'" _
					& " and codsede='" & Trim(codsede) & "'"
					if Ramo="COMP0003" then
		'				response.write(str) & "</br>"
					end if
					If BCL_ADO(str, Rst) Then
						If ValNulo(Rst("cupo"), NUM_) = 0 And ValNulo(Rst("cuponuevo"), NUM_) = 0 Then
							cupo = ValNulo(Rst("cupototal"), NUM_)
							TipoCupo = "CT"
						Else
							If session("EsRepit") Then
								cupo = ValNulo(Rst("cupo"), NUM_) 'CupoRepitente
								TipoCupo = "CR"
							Else
								cupo = ValNulo(Rst("cuponuevo"), NUM_) 'CupoNoRepitente
								TipoCupo = "CN"
							End If
						End If
				        Strsql = "Select sum(cupo) as cupo from ra_optivo"
						Strsql = Strsql & " Where CodSede = '" & CodSede & "' "
						Strsql = Strsql & " and CodRamo = '" & Trim(Ramo) & "' "
						Strsql = Strsql & " and CodSecc = '" & Trim(Seccion) & "' "
						Strsql = Strsql & " and ano = " & Ano & " "
						Strsql = Strsql & " and Periodo = " & Periodo & " "
						Strsql = Strsql & " and codcarrsec='" & Trim(lc_CodCarrSecc) & "'"
						If BCL_ADO(Strsql, Rst3) Then
							CupoOtrasCarreras = ValNulo(Rst3("cupo"), NUM_)
						end if
						if Ramo="COMP0003" then
							'response.write(Strsql)   
						end if
						SacaCupoAntiguo = cupo - (ValNulo(Rst("reserva"), NUM_) + CupoOtrasCarreras)
						'response.write(SacaCupoAntiguo)
						'response.end
						Rst.Close
					End If
				End If
				'rst.Close
			else
				SacaCupoAntiguo=0
			End If
        End If
    End If
	if Ramo="COMP0003" then
		'response.write(SacaCupoAntiguo)   
	end if
   'response.write(SacaCupoAntiguo)
End Function

Function SacaCupoOptativo(Ramo, RamoReal, Seccion, Ano, Periodo, CodCarr,CodSede) 
Dim Rst, Rst3
dim Str 
dim lc_CodCarrSecc 
dim Cupo
dim CupoOtrasCarreras 
dim Strsql

		TipoCupo = "CT"
		'aquí la variable ramo corresponde al ramo real que es dictado

        str = "Select codcarrsec,sum(cupo) as cupo from ra_optivo where" _
        & " codramo='" & Trim(RamoReal) & "'" _
        & " and codsecc='" & Trim(Seccion) & "'" _
        & " and ano=" & Ano & "" _
        & " and periodo=" & Periodo & " and" _
        & " codcarr='" & Trim(CodCarr) & "' and" _
        & " codsede='" & Trim(CodSede) & "' and tipo<>'OPTATIVO ESPECIAL' group by codcarrsec"
		
        'response.write(str)
        If BCL_ADO(str, Rst) Then 'Si hay cupos asignados para mi carrera
            lc_CodCarrSecc = Rst("codcarrsec")
            If valnulo(Rst("Cupo"),num_) > 0 Then 'Si los cupos son administrados por carrera
                Rst.Close
                str = "Select cupo from ra_optivo where" _
                & " codramo='" & Trim(RamoReal) & "'" _
                & " and codsecc='" & Trim(Seccion) & "'" _
                & " and ano=" & Ano & "" _
                & " and periodo=" & Periodo & "" _
                & " and codopt='" & Ramo & "'" _
				& " and codcarr='" & Trim(Codcarr) & "'" _
				& " and codsede='" & Trim(CodSede) & "' and tipo<>'OPTATIVO ESPECIAL'"
                If BCL_ADO(str, Rst) Then
					SacaCupoOptativo  = ValNulo(Rst("cupo"), NUM_)
					TipoCupo = "CT"
                    Rst.Close
                End If
                session("CupoCero") = "N"
            Else 'Si los cupos son administrados como cupo total
                session("CupoCero") = "S"
                Rst.Close
                str = "Select cupo,cupototal,cuponuevo,reserva from ra_seccio where" _
                & " codramo='" & Trim(RamoReal) & "'" _
                & " and codsecc='" & Trim(Seccion) & "'" _
                & " and ano=" & Ano & "" _
                & " and periodo=" & Periodo & ""_
				& " and codcarr='" & Trim(lc_CodCarrSecc) & "'" _
				& " and codsede='" & Trim(CodSede) & "'"
				'response.write(str)
                If BCL_ADO(str, Rst) Then
                    If ValNulo(Rst("cupo"), NUM_) = 0 And ValNulo(Rst("cuponuevo"), NUM_) = 0 Then
                        cupo = ValNulo(Rst("cupototal"), NUM_)
                        TipoCupo = "CT"
                    Else
                        If EsRepit Then
                            cupo = ValNulo(Rst("cupo"), NUM_) 'CupoRepitente
                            TipoCupo = "CR"
                        Else
                            cupo = ValNulo(Rst("cuponuevo"), NUM_) 'CupoNoRepitente
                            TipoCupo = "CN"
                        End If
                    End If
					Strsql = "Select sum(cupo) as cupo from ra_optivo"
					Strsql = Strsql & " Where CodSede = '" & CodSede & "' "
					Strsql = Strsql & " and CodRamo = '" & Trim(Ramoreal) & "' "
					Strsql = Strsql & " and CodSecc = '" & Trim(Seccion) & "' "
					Strsql = Strsql & " and ano = " & Ano & " "
					Strsql = Strsql & " and Periodo = " & Periodo & " "
					Strsql = Strsql & " and codcarrsec='" & Trim(lc_CodCarrSecc) & "'"
					If BCL_ADO(Strsql, Rst3) Then
						CupoOtrasCarreras = ValNulo(Rst3("cupo"), NUM_)
					end if
					SacaCupoOptativo = cupo - (ValNulo(Rst("reserva"), NUM_) + CupoOtrasCarreras)
                    'SacaCupoOptativo = cupo - ValNulo(Rst("reserva"), NUM_)
                    Rst.Close
				else
					SacaCupoOptativo = 0
                End If
            End If
        End If

End Function

Function SacaCupoOptativoEspecial(RamoReal, Seccion, Ano, Periodo, CodCarr, RamoMalla) 
Dim Str
Dim Rst
Dim Rst2
Dim Cupo

    session("CupoCero") = "N"

	str = "Select cupo,codcarrsec from ra_optivo where"
	str = str & " codramo='" & Trim(RamoReal) & "'"
	str = str & " and tipo='OPTATIVO ESPECIAL'"
	str = str & " and codsecc='" & Trim(Seccion) & "'"
	str = str & " and ano=" & Ano & ""
	str = str & " and periodo=" & Periodo & " "
	str = str & " and codsede='" & Trim(CodSede) & "'"
	str = str & " and codcarr='" & Trim(Codcarr) & "'"
	str = str & " and codopt='" & Trim(RamoMalla) & "'"
	
	if RamoMalla = "OPT0006" and RamoReal = "ARQT0067" then
		'response.write(str)
	end if
	If BCL_ADO(str, Rst2) Then
		If ValNulo(Rst2("cupo"), NUM_) = 0 Then 'Preguntamos si para mi carrera el cupo es cero
		    session("CupoCero") = "S"
			'Si el cupo es cero, debo indentificar el cupo asignado para mi carrera
			str = "Select sum(cupo) as cupo from ra_optivo where"
			str = str & " codramo='" & Trim(RamoReal) & "'"
			str = str & " and codsecc='" & Trim(Seccion) & "'"
			str = str & " and ano=" & Ano & ""
			str = str & " and periodo=" & Periodo & " "
			str = str & " and codcarrsec='" & Trim(Rst2("codcarrsec")) & "'" 'Carrera que dicta el ramo
			str = str & " and codsede='" & Trim(CodSede) & "' "
			str = str & " and codcarr+codopt<>'" & Trim(Codcarr + RamoMalla) & "' "
			if RamoMalla = "OPT0006" and RamoReal = "ARQT0067" then
				'response.write(str)
			end if
			If BCL_ADO(str, Rst) Then 'Si hay cupos asignados para otras carreras
				CupoOtrasCarreras = ValNulo(Rst("cupo"), NUM_)
			Else
				CupoOtrasCarreras = 0
			End If

			str = "Select cupo,cupototal,cuponuevo,reserva " _
			& " from ra_seccio where" _
			& " codramo='" & Trim(RamoReal) & "'" _
			& " and codsecc='" & Trim(Seccion) & "'" _
			& " and ano=" & Ano & "" _
			& " and periodo=" & Periodo & ""_
			& " and codcarr='" & Trim(Rst2("codcarrsec")) & "'" _
			& " and codsede='" & Trim(CodSede) & "'"
			if RamoMalla = "OPT0006" and RamoReal = "ARQT0067" then
				'response.write(str)
			end if
			If BCL_ADO(str, Rst) Then
				If ValNulo(Rst("cupo"), NUM_) = 0 And ValNulo(Rst("cuponuevo"), NUM_) = 0 Then
					cupo = ValNulo(Rst("cupototal"), NUM_)
					TipoCupo = "CT"
				Else
					If EsRepit Then
						cupo = ValNulo(Rst("cupo"), NUM_) 'CupoRepitente
						TipoCupo = "CR"
					Else
						cupo = ValNulo(Rst("cuponuevo"), NUM_) 'CupoNoRepitente
						TipoCupo = "CN"
					End If
				End If
				cupo = cupo - (CupoOtrasCarreras + ValNulo(Rst("reserva"), NUM_))
				Rst.Close
			End If
		Else 'Si el cupo está asignado a mi carrera
			Cupo = ValNulo(Rst2("cupo"), NUM_)
			TipoCupo = "CT"
		End If
	else
	    Cupo = 0	
	End If

	SacaCupoOptativoEspecial = Cupo
    'response.write(SacaCupoOptativoEspecial)
'	response.write(TipoCupo)    
End Function

Function CantidadInscritosOptativos(Curso, Ramo, Seccion, CodSede, Ano, Periodo, CodCarr, TipoCupo, TipoCurso)
Dim StrSql
Dim Rst
dim str 
  	
	If session("CupoCero")  = "N" Then ' hay cupos asignados a carreras 	
		If TipoCupo="CT" Then
			if TipoCurso = "T" then
				str = "Select count(codcli) from ra_carga where "
			else
				str = "Select count(codcli) from ra_cargaactividad where "
			end if
			str = str & " ramoequiv='" & Trim(Ramo) & "'"
			str = str & " and codsecc='" & Trim(Seccion) & "'"
			str = str & " and ano=" & Ano & ""
			str = str & " and periodo=" & Periodo & ""
			str = str & " and codcarr='" & Trim(CodCarr) & "'"
			str = str & " and inscrito = 'S'"
			str = str & " and codramo='" & Curso & "' "
			str = str & " and codcarr in (select codcarr from mt_carrer "
			str = str & " where sede='" & Trim(CodSede) & "')"
		ElseIf TipoCupo = "CR" Then
			if TipoCurso = "T" then
				str = "Select count(codcli) from ra_carga where"
			else
				str = "Select count(codcli) from ra_cargaactividad where"
			end if
			str = str & " ramoequiv='" & Trim(Ramo) & "'"
			str = str & " and codsecc='" & Trim(Seccion) & "'"
			str = str & " and ano=" & Ano & ""
			str = str & " and periodo=" & Periodo & ""
			str = str & " and codcarr='" & Trim(CodCarr) & "'"
			str = str & " and inscrito = 'S'"
			str = str & " and codramo='" & Curso & "' "
			str = str & " and codcarr in (select codcarr from mt_carrer "
			str = str & " where sede='" & Trim(CodSede) & "')"
			
			if TipoCurso = "T" then
				str = str & " and ra_carga.codcli in"
				str = str & " (Select ra_nota.codcli from ra_nota"
				str = str & " where (ra_nota.codramo='" & Trim(Ramo) & "' or"
				str = str & " ra_nota.ramoequiv='" & Trim(Ramo) & "')"
				str = str & " and estado in ('R'))"
			else
				str = str & " and ra_cargaactividad.codcli in"
				str = str & " (Select ra_notaactividad.codcli from ra_notaactividad"
				str = str & " where (ra_notaactividad.codramo='" & Trim(Ramo) & "' or"
				str = str & " ra_notaactividad.ramoequiv='" & Trim(Ramo) & "')"
				str = str & " and estado in ('R'))"
			end if
			
		ElseIf TipoCupo = "CN" Then
			if TipoCurso = "T" then
				str = "Select count(codcli) from ra_carga where"
			else
				str = "Select count(codcli) from ra_cargaactividad where"
			end if
			str = str & " ramoequiv='" & Trim(Ramo) & "'"
			str = str & " and codsecc='" & Trim(Seccion) & "'"
			str = str & " and ano=" & Ano & ""
			str = str & " and periodo=" & Periodo & ""
			str = str & " and codcarr='" & Trim(CodCarr) & "'"
			str = str & " and inscrito = 'S'"
			str = str & " and codramo='" & Curso & "' "
			str = str & " and codcarr in (select codcarr from mt_carrer "
			str = str & " where sede='" & Trim(CodSede) & "')"
			
			if TipoCurso = "T" then			
				str = str & " and ra_carga.codcli not in"
				str = str & " (Select ra_nota.codcli from ra_nota"
				str = str & " where (ra_nota.codramo='" & Trim(Ramo) & "' or"
				str = str & " ra_nota.ramoequiv='" & Trim(Ramo) & "')"
				str = str & " and estado in ('R'))"
			else
				str = str & " and ra_cargaactividad.codcli not in"
				str = str & " (Select ra_notaactividad.codcli from ra_notaactividad"
				str = str & " where (ra_notaactividad.codramo='" & Trim(Ramo) & "' or"
				str = str & " ra_notaactividad.ramoequiv='" & Trim(Ramo) & "')"
				str = str & " and estado in ('R'))"			
			end if
		End If
		
	Else 'no hay cupos asignados a carreras, pero pueden haber cupos cero
		If TipoCupo="CT" Then
			if TipoCurso = "T" then
				str = "Select count(codcli) from ra_carga where"
			else
				str = "Select count(codcli) from ra_cargaactividad where"
			end if
			str = str & " ramoequiv='" & Trim(Ramo) & "'"
			str = str & " and codsecc='" & Trim(Seccion) & "'"
			str = str & " and ano=" & Ano & ""
			str = str & " and periodo=" & Periodo & ""
			str = str & " and inscrito = 'S' and codcarr in "
			str = str & " (Select codcarr from ra_optivo where"
			str = str & " codramo='" & Trim(Ramo) & "'"
			str = str & " and codsecc='" & Trim(Seccion) & "'"
			str = str & " and ano=" & Ano & ""
			str = str & " and periodo=" & Periodo & ""
			str = str & " and tipo<>'OPTATIVO ESPECIAL' "
			str = str & " and codsede='" & Trim(CodSede) & "' and cupo=0)"			
		ElseIf TipoCupo = "CR" Then
			if TipoCurso = "T" then
				str = "Select count(codcli) from ra_carga where"
			else
				str = "Select count(codcli) from ra_cargaactividad where"
			end if
			str = str & " ramoequiv='" & Trim(Ramo) & "'"
			str = str & " and codsecc='" & Trim(Seccion) & "'"
			str = str & " and ano=" & Ano & ""
			str = str & " and periodo=" & Periodo & ""
			str = str & " and inscrito = 'S' and codcarr in "
			str = str & " (Select codcarr from ra_optivo where"
			str = str & " codramo='" & Trim(Ramo) & "'"
			str = str & " and codsecc='" & Trim(Seccion) & "'"
			str = str & " and ano=" & Ano & ""
			str = str & " and periodo=" & Periodo & ""
			str = str & " and tipo<>'OPTATIVO ESPECIAL' "
			str = str & " and codsede='" & Trim(CodSede) & "' and cupo=0)"		
			if TipoCurso = "T" then
				str = str & " and ra_carga.codcli in"
				str = str & " (Select ra_nota.codcli from ra_nota"
				str = str & " where (ra_nota.codramo='" & Trim(Ramo) & "' or"
				str = str & " ra_nota.ramoequiv='" & Trim(Ramo) & "')"
				str = str & " and estado in ('R'))"
			else
				str = str & " and ra_cargaactividad.codcli in"
				str = str & " (Select ra_notaactividad.codcli from ra_notaactividad"
				str = str & " where (ra_notaactividad.codramo='" & Trim(Ramo) & "' or"
				str = str & " ra_notaactividad.ramoequiv='" & Trim(Ramo) & "')"
				str = str & " and estado in ('R'))"
			end if			
		ElseIf TipoCupo = "CN" Then
			if TipoCurso = "T" then
				str = "Select count(codcli) from ra_carga where"
			else
				str = "Select count(codcli) from ra_cargaactividad where"
			end if
			str = str & " ramoequiv='" & Trim(Ramo) & "'"
			str = str & " and codsecc='" & Trim(Seccion) & "'"
			str = str & " and ano=" & Ano & ""
			str = str & " and periodo=" & Periodo & ""
			str = str & " and inscrito = 'S' and codcarr in "
			str = str & " (Select codcarr from ra_optivo where"
			str = str & " codramo='" & Trim(Ramo) & "'"
			str = str & " and codsecc='" & Trim(Seccion) & "'"
			str = str & " and ano=" & Ano & ""
			str = str & " and periodo=" & Periodo & ""
			str = str & " and tipo<>'OPTATIVO ESPECIAL' "
			str = str & " and codsede='" & Trim(CodSede) & "' and cupo=0)"			
			if TipoCurso = "T" then
				str = str & " and ra_carga.codcli not in"
				str = str & " (Select ra_nota.codcli from ra_nota"
				str = str & " where (ra_nota.codramo='" & Trim(Ramo) & "' or"
				str = str & " ra_nota.ramoequiv='" & Trim(Ramo) & "')"
				str = str & " and estado in ('R'))"
			else
				str = str & " and ra_cargaactividad.codcli not in"
				str = str & " (Select ra_notaactividad.codcli from ra_notaactividad"
				str = str & " where (ra_notaactividad.codramo='" & Trim(Ramo) & "' or"
				str = str & " ra_notaactividad.ramoequiv='" & Trim(Ramo) & "')"
				str = str & " and estado in ('R'))"
			end if			
		End If
	End If
		
	If BCL_ADO(str, rst) Then
		CantidadInscritosOptativos = ValNulo(rst(0), NUM_)
	End If	
End Function

Function CantidadInscritosOptativoEspecial(Ramo, Seccion, CodSede, Ano, Periodo, CodCarr, TipoCupo, TipoCurso) 
Dim StrSql
Dim Rst
    
	If session("CupoCero") = "N" Then
		If TipoCupo="CT" Then
			if TipoCurso = "T" then
				str = "Select count(codcli) from ra_carga where"
			else
				str = "Select count(codcli) from ra_cargaactividad where"
			end if
			str = str & " ramoequiv='" & Trim(Ramo) & "'"
			str = str & " and codsecc='" & Trim(Seccion) & "'"
			str = str & " and ano=" & Ano & ""
			str = str & " and periodo=" & Periodo & ""
			str = str & " and codcarr='" & Trim(CodCarr) & "'"
			str = str & " and inscrito = 'S'"
		ElseIf TipoCupo="CR" Then
			if TipoCurso = "T" then
				str = "Select count(codcli) from ra_carga where"
			else
				str = "Select count(codcli) from ra_cargaactividad where"
			end if
			str = str & " ramoequiv='" & Trim(Ramo) & "'"
			str = str & " and codsecc='" & Trim(Seccion) & "'"
			str = str & " and ano=" & Ano & ""
			str = str & " and periodo=" & Periodo & ""
			str = str & " and codcarr='" & Trim(CodCarr) & "'"
			str = str & " and inscrito = 'S'"
			if TipoCurso = "T" then
				str = str & " and ra_carga.codcli in"
				str = str & " (Select ra_nota.codcli from ra_nota"
				str = str & " where (ra_nota.codramo='" & Trim(Ramo) & "' or"
				str = str & " ra_nota.ramoequiv='" & Trim(Ramo) & "')"
				str = str & " and estado in ('R'))"
			else
				str = str & " and ra_cargaactividad.codcli in"
				str = str & " (Select ra_notaactividad.codcli from ra_notaactividad"
				str = str & " where (ra_notaactividad.codramo='" & Trim(Ramo) & "' or"
				str = str & " ra_notaactividad.ramoequiv='" & Trim(Ramo) & "')"
				str = str & " and estado in ('R'))"
			end if		
		ElseIf TipoCupo="CN" Then
			if TipoCurso = "T" then
				str = "Select count(codcli) from ra_carga where"
			else
				str = "Select count(codcli) from ra_cargaactividad where"
			end if
			str = str & " ramoequiv='" & Trim(Ramo) & "'"
			str = str & " and codsecc='" & Trim(Seccion) & "'"
			str = str & " and ano=" & Ano & ""
			str = str & " and periodo=" & Periodo & ""
			str = str & " and codcarr='" & Trim(CodCarr) & "'"
			str = str & " and inscrito = 'S'"
			if TipoCurso = "T" then
				str = str & " and ra_carga.codcli not in"
				str = str & " (Select ra_nota.codcli from ra_nota"
				str = str & " where (ra_nota.codramo='" & Trim(Ramo) & "' or"
				str = str & " ra_nota.ramoequiv='" & Trim(Ramo) & "')"
				str = str & " and estado in ('R'))"
			else
				str = str & " and ra_cargaactividad.codcli not in"
				str = str & " (Select ra_notaactividad.codcli from ra_notaactividad"
				str = str & " where (ra_notaactividad.codramo='" & Trim(Ramo) & "' or"
				str = str & " ra_notaactividad.ramoequiv='" & Trim(Ramo) & "')"
				str = str & " and estado in ('R'))"
			end if
		End If
		'response.write(TipoCupo)
		'response.write(str)
		If BCL_ADO(str, rst) Then
			CantidadInscritosOptativoEspecial = ValNulo(rst(0), NUM_)
		else
			CantidadInscritosOptativoEspecial = 0
		End If
	Else
		If TipoCupo="CT" Then
			if TipoCurso = "T" then
				str = "Select count(codcli) from ra_carga where"
			else
				str = "Select count(codcli) from ra_cargaactividad where"
			end if
			str = str & " ramoequiv='" & Trim(Ramo) & "'"
			str = str & " and codsecc='" & Trim(Seccion) & "'"
			str = str & " and ano=" & Ano & ""
			str = str & " and periodo=" & Periodo & ""
			str = str & " and inscrito = 'S' and codcarr in "
			str = str & " (Select ra_optivo.codcarr from ra_optivo where"
			str = str & " ra_optivo.codramo='" & Trim(Ramo) & "'"
			str = str & " and ra_optivo.codsecc='" & Trim(Seccion) & "'"
			str = str & " and ra_optivo.ano=" & Ano & ""
			str = str & " and ra_optivo.periodo=" & Periodo & ""
			str = str & " and ra_optivo.tipo='OPTATIVO ESPECIAL' "
			str = str & " and ra_optivo.codsede='" & Trim(CodSede) & "'"
			str = str & " and ra_optivo.cupo=0)"
		ElseIf TipoCupo="CR" Then
			if TipoCurso = "T" then
				str = "Select count(codcli) from ra_carga where"
			else
				str = "Select count(codcli) from ra_cargaactividad where"
			end if
			str = str & " ramoequiv='" & Trim(Ramo) & "'"
			str = str & " and codsecc='" & Trim(Seccion) & "'"
			str = str & " and ano=" & Ano & ""
			str = str & " and periodo=" & Periodo & ""
			str = str & " and inscrito = 'S' and codcarr in "
			str = str & " (Select ra_optivo.codcarr from ra_optivo where"
			str = str & " ra_optivo.codramo='" & Trim(Ramo) & "'"
			str = str & " and ra_optivo.codsecc='" & Trim(Seccion) & "'"
			str = str & " and ra_optivo.ano=" & Ano & ""
			str = str & " and ra_optivo.periodo=" & Periodo & ""
			str = str & " and ra_optivo.tipo='OPTATIVO ESPECIAL' "
			str = str & " and ra_optivo.codsede='" & Trim(CodSede) & "'"
			str = str & " and ra_optivo.cupo=0)"
			if TipoCurso = "T" then
				str = str & " and ra_carga.codcli in"
				str = str & " (Select ra_nota.codcli from ra_nota"
				str = str & " where (ra_nota.codramo='" & Trim(Ramo) & "' or"
				str = str & " ra_nota.ramoequiv='" & Trim(Ramo) & "')"
				str = str & " and estado in ('R'))"
			else
				str = str & " and ra_cargaactividad.codcli in"
				str = str & " (Select ra_notaactividad.codcli from ra_notaactividad"
				str = str & " where (ra_notaactividad.codramo='" & Trim(Ramo) & "' or"
				str = str & " ra_notaactividad.ramoequiv='" & Trim(Ramo) & "')"
				str = str & " and estado in ('R'))"
			end if
		ElseIf TipoCupo="CN" Then
			if TipoCurso = "T" then
				str = "Select count(codcli) from ra_carga where"
			else
				str = "Select count(codcli) from ra_cargaactividad where"
			end if
			str = str & " ramoequiv='" & Trim(Ramo) & "'"
			str = str & " and codsecc='" & Trim(Seccion) & "'"
			str = str & " and ano=" & Ano & ""
			str = str & " and periodo=" & Periodo & ""
			str = str & " and inscrito = 'S' and codcarr in "
			str = str & " (Select ra_optivo.codcarr from ra_optivo where"
			str = str & " ra_optivo.codramo='" & Trim(Ramo) & "'"
			str = str & " and ra_optivo.codsecc='" & Trim(Seccion) & "'"
			str = str & " and ra_optivo.ano=" & Ano & ""
			str = str & " and ra_optivo.periodo=" & Periodo & ""
			str = str & " and ra_optivo.tipo='OPTATIVO ESPECIAL' "
			str = str & " and ra_optivo.codsede='" & Trim(CodSede) & "'"
			str = str & " and ra_optivo.cupo=0)"
			if TipoCurso = "T" then
				str = str & " and ra_carga.codcli not in"
				str = str & " (Select ra_nota.codcli from ra_nota"
				str = str & " where (ra_nota.codramo='" & Trim(Ramo) & "' or"
				str = str & " ra_nota.ramoequiv='" & Trim(Ramo) & "')"
				str = str & " and estado in ('R'))"
			else
				str = str & " and ra_cargaactividad.codcli not in"
				str = str & " (Select ra_notaactividad.codcli from ra_notaactividad"
				str = str & " where (ra_notaactividad.codramo='" & Trim(Ramo) & "' or"
				str = str & " ra_notaactividad.ramoequiv='" & Trim(Ramo) & "')"
				str = str & " and estado in ('R'))"
			end if
		End If
		if Ramo = "ARQT0067" then
			'response.write(str)
		end if
		If BCL_ADO(str, rst) Then
	        CantidadInscritosOptativoEspecial = ValNulo(rst(0), NUM_)
		else
			CantidadInscritosOptativoEspecial = 0
		End If
	End If

End Function

Function CantidadInscritos(Ramo, Seccion, CodSede, Ano, Periodo, TipoCupo, TipoCurso)
Dim StrSql
Dim Rst

    'Response.Write("TipoCupo = " & TipoCupo & " Ramo = " & Ramo)

    If TipoCupo = "CT" Then 'Cupo Total
		if TipoCurso = "T" then
	        StrSql = "Select count(s.codcli) from ra_carga s, mt_carrer c where"
		else
			StrSql = "Select count(s.codcli) from ra_cargaactividad s, mt_carrer c where"
		end if
        StrSql = StrSql & " s.ramoequiv='" & Trim(Ramo) & "'"
		StrSql = StrSql & " and s.codsecc='" & Trim(Seccion) & "'"
		StrSql = StrSql & " and s.ano=" & Ano & ""
		StrSql = StrSql & " and s.periodo=" & Periodo & ""
		StrSql = StrSql & " and s.inscrito='S'"
		StrSql = StrSql & " and s.codcarr = c.codcarr "
		StrSql = StrSql & " and c.sede = '" & CodSede & "' "
        If BCL_ADO(StrSql, Rst) Then
            CantidadInscritos = ValNulo(rst(0), NUM_)
            Rst.close()
        else
            CantidadInscritos = 0
        End If
        '& " and codcarr='" & CodCarr & "'"
		
    ElseIf TipoCupo = "CR" Then 'Cupo Repitente
		if TipoCurso = "T" then
	        StrSql = "Select count(s.codcli) from ra_carga s, mt_carrer c where"
		else
			StrSql = "Select count(s.codcli) from ra_cargaactividad s, mt_carrer c where"
		end if
        StrSql = StrSql & " s.ramoequiv='" & Trim(Ramo) & "'"
        StrSql = StrSql & " and s.codsecc='" & Trim(Seccion) & "'"
        StrSql = StrSql & " and s.ano=" & Ano & ""
        StrSql = StrSql & " and s.periodo=" & Periodo & ""
        StrSql = StrSql & " and s.inscrito='S'"
        StrSql = StrSql & " and s.codcarr = c.codcarr "
        StrSql = StrSql & " and c.sede = '" & CodSede & "' "
        StrSql = StrSql & " and s.codcli in"
		if TipoCupo = "T" then
			StrSql = StrSql & " (Select ra_nota.codcli from ra_nota"
			StrSql = StrSql & " where (ra_nota.codramo='" & Trim(Ramomalla) & "' or"
			StrSql = StrSql & " ra_nota.ramoequiv='" & Trim(Ramo) & "')"
			StrSql = StrSql & " and estado in ('R'))"		
		else
			StrSql = StrSql & " (Select ra_notaactividad.codcli from ra_notaactividad"
			StrSql = StrSql & " where (ra_notaactividad.codramo='" & Trim(Ramomalla) & "' or"
			StrSql = StrSql & " ra_notaactividad.ramoequiv='" & Trim(Ramo) & "')"
			StrSql = StrSql & " and estado in ('R'))"		
		end if
        If BCL_ADO(StrSql, Rst) Then
            CantidadInscritos = ValNulo(rst(0), NUM_)
            Rst.close()
        Else
            CantidadInscritos = 0
        End If
        ' " and codcarr='" & CodCarr & "'"
    ElseIf TipoCupo = "CN" Then 'Cupo No Repitente
		if TipoCurso = "T" then
	        StrSql = "Select count(s.codcli) from ra_carga s, mt_carrer c where"
		else
			StrSql = "Select count(s.codcli) from ra_cargaactividad s, mt_carrer c where"
		end if
        StrSql = StrSql & " s.ramoequiv='" & Trim(Ramo) & "'"
		StrSql = StrSql & " and s.codsecc='" & Trim(Seccion) & "'"
        StrSql = StrSql & " and s.ano=" & Ano & ""
        StrSql = StrSql & " and s.periodo=" & Periodo & ""
        StrSql = StrSql & " and s.inscrito='S'"
        StrSql = StrSql & " and s.codcarr = c.codcarr "
        StrSql = StrSql & " and c.sede = '" & CodSede & "' "
        StrSql = StrSql & " and s.codcli not in"
		if TipoCurso = "T" then
			StrSql = StrSql & " (Select ra_nota.codcli from ra_nota"
			StrSql = StrSql & " where (ra_nota.codramo='" & Trim(Ramo) & "' or"
			StrSql = StrSql & " ra_nota.ramoequiv='" & Trim(Ramo) & "')"
			StrSql = StrSql & " and estado in ('R'))"		
		else
			StrSql = StrSql & " (Select ra_notaactividad.codcli from ra_notaactividad"
			StrSql = StrSql & " where (ra_notaactividad.codramo='" & Trim(Ramomalla) & "' or"
			StrSql = StrSql & " ra_notaactividad.ramoequiv='" & Trim(Ramo) & "')"
			StrSql = StrSql & " and estado in ('R'))"		
		end if
        If BCL_ADO(StrSql, Rst) Then
            CantidadInscritos = ValNulo(Rst(0), NUM_)
            Rst.close()
        Else
            CantidadInscritos = 0
        End If
        '& " and codcarr='" & CodCarr & "'"
    Else
		if TipoCurso = "T" then
	        StrSql = "Select count(s.codcli) from ra_carga s, mt_carrer c where"
		else
			StrSql = "Select count(s.codcli) from ra_cargaactividad s, mt_carrer c where"
		end if
        StrSql = StrSql & " s.ramoequiv='" & Trim(Ramo) & "'"
        StrSql = StrSql & " and s.codsecc='" & Trim(Seccion) & "'"
        StrSql = StrSql & " and s.ano=" & Ano & ""
        StrSql = StrSql & " and s.periodo=" & Periodo & ""
        StrSql = StrSql & " and s.inscrito='S'"
        StrSql = StrSql & " and s.codcarr = c.codcarr "
        StrSql = StrSql & " and c.sede = '" & CodSede & "' "
        StrSql = StrSql & " and s.codcarr = '" & CodCarr & "'"
        If BCL_ADO(StrSql, Rst) Then
            CantidadInscritos = ValNulo(rst(0), NUM_)
            Rst.close()
        Else
            CantidadInscritos = 0
        End If
        '& " and codcarr='" & CodCarr & "'"
    End If
End Function
%>
<%

Function ValidaInscripcionLPT(Codcli,Ano,Periodo,Codpestud,Codcarr,Codsede)
	
dim Str
dim Rst
dim Rst2
dim Rst3
	
	ValidaInscripcionLPT = 1
	
	Str = "SELECT b.codramo, b.nombre, a.inscrito, a.codsecc, a.ramoequiv, b.credito, a.ramoequiv_i, a.codsecc_i, a.preinscrito "  
	Str = Str & " FROM ra_carga a, ra_ramo b" 
	Str = Str & " WHERE a.codramo = b.codramo and " 
	Str = Str & " a.codcli ='" & CodCli & "' and " 
	Str = Str & " a.ano = '" & ano & "' and " 
	Str = Str & " a.periodo = '" & periodo & "' and " 
	Str = Str & " a.preinscrito='S' and (coalesce(a.inscrito,'')<>'S' or (a.inscrito='S' and (a.codsecc<>a.codsecc_i or a.ramoequiv<>a.ramoequiv_i)))"
	'10/0
	
	'response.write(str)
	'response.end
	
	if bcl_ado(str,rst) then
		do while not rst.eof
			str ="select codramo from ra_seccio " 
			str = str & " where ano=" & ano & " "
			str = str & " and periodo = " & periodo & ""
			str = str & " and codsede = '" & CodSede & "'"
			str = str & " and codramo ='" & trim(rst("ramoequiv_i")) & "'"
			str = str & " and codseccteo ='" & trim(rst("codsecc_i")) & "'"
			str = str & " and coalesce(tipocurso,'') <>'T'"
			str = str & " and codramo + tipocurso COLLATE Modern_Spanish_CI_AS not in "
			str = str & " (select ramoequiv_i + tipocurso COLLATE Modern_Spanish_CI_AS from ra_cargaactividad" 
			str = str & " where codcli = '" & CodCli & "'" 
			str = str & " and Ano = " & ano & ""
			str = str & " and Periodo = " & periodo & ""
			Str = Str & " and preinscrito='S')"' and (coalesce(inscrito,'')<>'S' or (inscrito='S' and (codsecc<>codsecc_i or ramoequiv<>ramoequiv_i))) and tipocurso in ('L','P'))"			
	
	        'response.write(str)
	       	'response.end
			
			if bcl_ado(str,rst3) then
				while not rst3.eof
					ValidaInscripcionLPT = 0
					 %>
					 <script>
					 alert("Falta inscribir la actividad <%= rst("codramo")%> No se puede inscribir solo la asignatura teórica ....!!!!")
					 </script>
					 <%  			
					rst3.Movenext
				wend
				exit Do
			end if
			rst.movenext
		loop
		 'response.end
	end if	
End Function

Function ValidaInscripcionTeorico(Codcli,Ano,Periodo,Codpestud,Codcarr,Codsede)

dim Str
dim Rst
dim Rst2
dim Rst3
	
	ValidaInscripcionTeorico = 1
	
	Str = "SELECT b.codramo, b.nombre, a.inscrito, a.codsecc, a.ramoequiv, b.credito, a.ramoequiv_i, a.codsecc_i, a.preinscrito "  
	Str = Str & " FROM ra_cargaactividad a, ra_ramo b" 
	Str = Str & " WHERE a.codramo = b.codramo and " 
	Str = Str & " a.codcli ='" & CodCli & "' and " 
	Str = Str & " a.ano = '" & ano & "' and " 
	Str = Str & " a.periodo = '" & periodo & "' and " 
	Str = Str & " a.preinscrito='S' and (coalesce(a.inscrito,'')<>'S' or (a.inscrito='S' and (a.codsecc<>a.codsecc_i or a.ramoequiv<>a.ramoequiv_i)))"

	'10/0
	
	'response.write(str)
	'response.end
	
	if bcl_ado(str,rst) then
		do while not rst.eof
			str ="select codramo from ra_seccio " 
			str = str & " where ano=" & ano & " "
			str = str & " and periodo = " & periodo & ""
			str = str & " and codsede = '" & CodSede & "'"
			str = str & " and codramo ='" & trim(rst("ramoequiv_i")) & "'"
			str = str & " and codsecc ='" & trim(rst("codsecc_i")) & "'"
			str = str & " and coalesce(tipocurso,'') <>'T'"
			str = str & " and codramo COLLATE Modern_Spanish_CI_AS not in "
			str = str & " (select ramoequiv_i COLLATE Modern_Spanish_CI_AS from ra_carga" 
			str = str & " where codcli = '" & CodCli & "'" 
			str = str & " and Ano = " & ano & ""
			str = str & " and Periodo = " & periodo & ""
			Str = Str & " and preinscrito='S')"' and (coalesce(inscrito,'')<>'S' or (inscrito='S' and (codsecc<>codsecc_i or ramoequiv<>ramoequiv_i))))"			
	
	        'response.write(str)
	       	'response.end
			
			if bcl_ado(str,rst3) then
				while not rst3.eof
					 ValidaInscripcionTeorico = 0
					 %>
					 <script>
					 alert("Falta inscribir la asignatura <%= rst("codramo")%> No se puede inscribir solo la actividad ....!!!!")
					 </script>
					 <%  			
					rst3.Movenext
				wend
				exit Do
			end if
			rst.movenext
		loop
		 'response.end
	end if	
End Function

Function ValidaCuposPreinscritos (Codcli,Ano,Periodo,Codpestud,Codcarr,Codsede)
dim Str 
dim Rst
dim TipoRamo
dim Inscritos
dim Cupos
dim ValidaCuposPreinscritosAux
dim i 
dim s_CodPestud
dim TipoRamoOtros
dim EsRepit
dim TipoCursoInterno

ValidaCuposPreinscritosAux = 1
TipoCursoInterno = "T"

	
	Str = "SELECT b.codramo, b.nombre, a.inscrito, a.codsecc, a.ramoequiv, b.credito, a.ramoequiv_i, a.codsecc_i, a.preinscrito "  
	Str = Str & " FROM ra_carga a, ra_ramo b" 
	Str = Str & " WHERE a.codramo = b.codramo and " 
	Str = Str & " a.codcli ='" & CodCli & "' and " 
	Str = Str & " a.ano = '" & ano & "' and " 
	Str = Str & " a.periodo = '" & periodo & "' and " 
	Str = Str & " a.preinscrito='S' and (coalesce(a.inscrito,'')<>'S' or (a.inscrito='S' and (a.codsecc<>a.codsecc_i or a.ramoequiv<>a.ramoequiv_i)))"
	Str = Str & " Order By a.prioridad "
	
	if bcl_ado(str,rst) then
		do while not rst.eof
  	         session("OtrosRamos") = "N"
             EsRepit = EsRepitente(rst("codramo"), CodCli) 				 
             Session("EsRepit") = EsRepit
				
             TipoRamo = trim(GetTipoRamo(rst("codramo"), CodPestud))      
 	     	 s_CodPestud = trim(GetPlanOtrosRamos(CodPestud,rst("codramo")))
			 TipoRamoOtros = GetTipoRamo(rst("codramo"), s_CodPestud)
			 			 
             session("OtrosRamos" & rst("codramo")) = "N"						                   
			 Select Case TipoRamo
 			   Case ""

				  'response.end 	
			   	  if TipoRamoOtros = "OBLIGATORIO" then		
                  	session("OtrosRamos") = "S"						
				  end if    
			      'x = 10/0       

				  Cupos = SacaCupoAntiguo (rst("ramoequiv_i"), rst("codsecc_i"), Ano, Periodo, CodCarr, CodSede, EsRepitente( rst("ramoequiv_i"), CodCli) , TipoCupo,rst("codramo"))
  				  if session("RamoObliPlanificadoOp") = "N" then
	                  Inscritos = CantidadInscritos (rst("ramoequiv_i"), rst("codsecc_i"), CodSede, Ano, Periodo, TipoCupo, TipoCursoInterno)
				  else
				  	  'Se analiza el cupo asignado por otra carrera	
					  Inscritos = CantidadInscritosOptativos(rst("codramo"),  rst("ramoequiv_i"), rst("codsecc_i"), CodSede, Ano, Periodo, CodCarr, TipoCupo, TipoCursoInterno) 
				  end if
				if rst("ramoequiv_i") ="COMP0003" and rst("codramo")="OPT0005" then 
				'	response.write("Al confirmar ") & "</br>"
				'	response.write("Ramoequiv_i "  & rst("ramoequiv_i")) & "</br>"	
				'	response.write("Codsecc_i "  & rst("codsecc_i")) & "</br>"
				'	response.write("Año "  & Ano) & "</br>"
				'	response.write("Periodo "  & Periodo) & "</br>"
				'	response.write("CodCarr "  & CodCarr) & "</br>"
				'	response.write("CodSede "  & CodSede) & "</br>"
				'	response.write("Es Repitente "  & EsRepitente( rst("ramoequiv_i"),codcli)) & "</br>"
				'	response.write("TipoCupo "  & TipoCupo) & "</br>"
				'	response.write("CodRamo "  & rst("codramo")) & "</br>"
				'	response.write("Cupos "  & Cupos) & "</br>"
				'	response.end				   
				end if
			   Case "OBLIGATORIO"
				  Cupos = SacaCupoAntiguo (rst("ramoequiv_i"), rst("codsecc_i"), Ano, Periodo, CodCarr, CodSede, EsRepitente( rst("ramoequiv_i"), CodCli) , TipoCupo,rst("codramo"))			  
  				  if session("RamoObliPlanificadoOp") = "N" then
	                  Inscritos = CantidadInscritos (rst("ramoequiv_i"), rst("codsecc_i"), CodSede, Ano, Periodo, TipoCupo, TipoCursoInterno)
				  else
					  Inscritos = CantidadInscritosOptativos(rst("codramo"),  rst("ramoequiv_i"), rst("codsecc_i"), CodSede, Ano, Periodo, CodCarr, TipoCupo, TipoCursoInterno) 
				  end if	
				  
			   Case "OPCIONAL"
			   	  response.write("Opcional"  & rst("codramo")) & "</br>"
			   	  Cupos=0
				  Inscritos=0
				  Cupos = SacaCupoOptativo(rst("codramo"),rst("ramoequiv_i"), rst("codsecc_i"), Ano, Periodo, CodCarr,CodSede) 
				  Inscritos = CantidadInscritosOptativos(rst("codramo"),  rst("ramoequiv_i"), rst("codsecc_i"), CodSede, Ano, Periodo, CodCarr, TipoCupo, TipoCursoInterno) 

			   Case "OPTATIVO ESPECIAL"
				  'Response.Write("Opcional especial" & rst("codramo")) & "</br>"
				  
				  Cupos=0
				  Inscritos=0
				 
				  Cupos = SacaCupoOptativoEspecial( rst("ramoequiv_i"), rst("codsecc_i"), Ano, Periodo, CodCarr,rst("codramo")) 	
				  'response.write(TipoCupo)
				  Inscritos = CantidadInscritosOptativoEspecial(rst("ramoequiv_i"), rst("codsecc_i"), CodSede, Ano, Periodo,CodCarr, TipoCupo, TipoCursoInterno)
				
			 End Select
 			'response.end
			 if Inscritos < cupos  then
			 	ValidaCuposPreinscritos=1
			 else
			 %>
			 <script>
			 alert("Esta asignatura ya no tiene cupo, debes revisar y modificar tu inscripcion en el Ramo <%= rst("ramoequiv_i")%>  ya que la vacante, fuè Confirmado por otro alumno ....!!!!")
			 </script>
			 <%  
				session("RamoSinCupo")=rst("ramoequiv_i")
				ValidaCuposPreinscritosAux=0
				str=" update ra_carga set preinscrito=null, ramoequiv_i=null,codsecc_i=null "
				str = str & " where codcli='" & CodCli & "' and ramoequiv_i='"& rst("ramoequiv_i") &"' "
				str = str & " and codsecc_i='"& rst("codsecc_i") &"' and ano="& ano &" and periodo="& periodo &""
				conn.execute(str)
			 end if
			 rst.movenext
			 i=i+1
		loop
	else
		ValidaCuposPreinscritosAux=1
	end if
	
	ValidaCuposPreinscritos=ValidaCuposPreinscritosAux

End Function

Function ValidaCuposPreinscritosActividad (Codcli,Ano,Periodo,Codpestud,Codcarr,Codsede)
dim Str 
dim Rst
dim TipoRamo
dim Inscritos
dim Cupos
dim ValidaCuposPreinscritosActividadAux
dim i 
dim s_CodPestud
dim TipoRamoOtros
dim EsRepit
dim TipoCursoInterno

ValidaCuposPreinscritosActividadAux = 1

	Str = "SELECT b.codramo, b.nombre, a.inscrito, a.codsecc, a.ramoequiv, b.credito, a.ramoequiv_i, a.codsecc_i, a.preinscrito, a.TipoCurso "  
	Str = Str & " FROM ra_cargaactividad a, ra_ramo b" 
	Str = Str & " WHERE a.codramo = b.codramo and " 
	Str = Str & " a.codcli ='" & CodCli & "' and " 
	Str = Str & " a.ano = '" & ano & "' and " 
	Str = Str & " a.periodo = '" & periodo & "' and " 
	Str = Str & " a.preinscrito='S' and (coalesce(a.inscrito,'')<>'S' or (a.inscrito='S' and (a.codsecc<>a.codsecc_i or a.ramoequiv<>a.ramoequiv_i))) and "
	Str = Str & " a.TipoCurso <> 'T' "
	Str = Str & " Order By a.prioridad "
	
	'response.write(str)
	'response.end
	'x=0/0				  		 	
	if bcl_ado(str,rst) then
		do while not rst.eof
			
			 session("OtrosRamos") = "N"
             EsRepit = EsRepitente(rst("codramo"), CodCli) 				 
             Session("EsRepit") = EsRepit
					
             TipoRamo = trim(GetTipoRamo(rst("codramo"), CodPestud))      
 	     	 s_CodPestud = trim(GetPlanOtrosRamos(CodPestud,rst("codramo")))
			 TipoRamoOtros = GetTipoRamo(rst("codramo"), s_CodPestud)
			  
	'		 response.write ("Validando cupos TPL")
'			 response.end 			 
             session("OtrosRamos") = "N"						                   
			 Select Case TipoRamo
 			   Case ""

				  'response.end 	
			   	  if TipoRamoOtros = "OBLIGATORIO" then		
                  	session("OtrosRamos") = "S"						
				  end if    

			   	  'response.write("otros ramos") & "</br>"
				               
				  Cupos = SacaCupoAntiguo (rst("ramoequiv_i"), rst("codsecc_i"), Ano, Periodo, CodCarr, CodSede, EsRepitente( rst("ramoequiv_i"), CodCli) , TipoCupo,rst("codramo"))
  				  if session("RamoObliPlanificadoOp") = "N" then
	                  Inscritos = CantidadInscritos (rst("ramoequiv_i"), rst("codsecc_i"), CodSede, Ano, Periodo, TipoCupo, trim(rst("TipoCurso")) )
				  else
				  	  'Se analiza el cupo asignado por otra carrera	
					  Inscritos = CantidadInscritosOptativos(rst("codramo"),  rst("ramoequiv_i"), rst("codsecc_i"), CodSede, Ano, Periodo, CodCarr, TipoCupo,  trim(rst("TipoCurso"))) 
				  end if	
								   
			   Case "OBLIGATORIO"
			   	  'response.write("obligatorios") & "</br>"
				  Cupos = SacaCupoAntiguo (rst("ramoequiv_i"), rst("codsecc_i"), Ano, Periodo, CodCarr, CodSede, EsRepitente( rst("ramoequiv_i"), CodCli) , TipoCupo,rst("codramo"))			  
  				  if session("RamoObliPlanificadoOp") = "N" then
	                  Inscritos = CantidadInscritos (rst("ramoequiv_i"), rst("codsecc_i"), CodSede, Ano, Periodo, TipoCupo,  trim(rst("TipoCurso")))
				  else
					  Inscritos = CantidadInscritosOptativos(rst("codramo"),  rst("ramoequiv_i"), rst("codsecc_i"), CodSede, Ano, Periodo, CodCarr, TipoCupo,  trim(rst("TipoCurso"))) 
				  end if	
				  
			   Case "OPCIONAL"
			    
			   	  Cupos=0
				  Inscritos=0
  			      'Response.Write("Hola...")
				  'Response.end
				  Cupos = SacaCupoOptativo(rst("codramo"),rst("ramoequiv_i"), rst("codsecc_i"), Ano, Periodo, CodCarr,CodSede) 
				  Inscritos = CantidadInscritosOptativos(rst("codramo"),  rst("ramoequiv_i"), rst("codsecc_i"), CodSede, Ano, Periodo, CodCarr, TipoCupo,  trim(rst("TipoCurso"))) 

				'  if rst("ramoequiv_i") ="CFGE0013" and rst("codramo")="OPT0003" then 
				'	response.write("Al confirmar ") & "</br>"
				'	response.write("Ramoequiv_i "  & rst("ramoequiv_i")) & "</br>"	
				'	response.write("Codsecc_i "  & rst("codsecc_i")) & "</br>"
				'	response.write("Año "  & Ano) & "</br>"
				'	response.write("Periodo "  & Periodo) & "</br>"
				'	response.write("CodCarr "  & CodCarr) & "</br>"
				'	response.write("CodSede "  & CodSede) & "</br>"
				'	response.write("Es Repitente "  & EsRepitente( rst("ramoequiv_i"),codcli)) & "</br>"
				'	response.write("TipoCupo "  & TipoCupo) & "</br>"
				'	response.write("CodRamo "  & rst("codramo")) & "</br>"
				'	response.write("Cupos "  & Cupos) & "</br>"
				'	response.end				   
				'  end if
			   Case "OPTATIVO ESPECIAL"
				  'Response.Write("  Opcional especial") & "</br>"
				  'Response.Write(rst("codramo")) & "</br>"
				  'Response.Write(rst("ramoequiv_i")) & "</br>"
		
				  Cupos=0
				  Inscritos=0
				 
				  Cupos = SacaCupoOptativoEspecial( rst("ramoequiv_i"), rst("codsecc_i"), Ano, Periodo, CodCarr,rst("codramo")) 	
				  'response.write(TipoCupo)
				  Inscritos = CantidadInscritosOptativoEspecial(rst("ramoequiv_i"), rst("codsecc_i"), CodSede, Ano, Periodo,CodCarr, TipoCupo,  trim(rst("TipoCurso")))

			 End Select
			 'response.write ("Inscritos/Cupos: " & rst("ramoequiv_i") & " " & Inscritos & "/" & Cupos & " TipoCurso: " & trim(rst("TipoCurso"))) & "</br>"
			 if Inscritos < cupos  then
			 	ValidaCuposPreinscritosActividad = 1
			 else
			 %>
			 <script>
			 alert("Esta actividad ya no tiene cupo, debes revisar y modificar tu inscripcion en la Actividad <%= rst("ramoequiv_i")%>  ya que la vacante, fuè Confirmado por otro alumno ....!!!!")
			 </script>
			 <%  
				session("RamoSinCupo")=rst("ramoequiv_i")
				ValidaCuposPreinscritosActividadAux=0
				str=" delete from ra_cargaactividad "
				str = str & " where codcli='" & CodCli & "' and ramoequiv_i='"& rst("ramoequiv_i") &"' "
				str = str & " and codsecc_i='"& rst("codsecc_i") &"' and ano="& ano &" and periodo="& periodo &""
				conn.execute(str)
			 end if
			 rst.movenext
			 i=i+1
		loop
	else
		ValidaCuposPreinscritosActividadAux=1
	end if
	
	ValidaCuposPreinscritosActividad=ValidaCuposPreinscritosActividadAux

'	if ValidaCuposPreinscritosAux=0 then
'		ValidaCuposPreinscritosActividad=0
'	else
'		ValidaCuposPreinscritosActividad=1
'	end if

    ' response.write (Inscritos) & "</br>"
	' response.write (cupos) & "</br>"
	'response.Write("stop antes de salir") & "</br>"
	'response.write(ValidaCuposPreinscritosActividad)
	'response.end			   
End Function

Function GetRamoFueraMallaRep(Codramo, Codcli, Ano, Periodo) 
Dim Str
Dim Rst 

	str = "Select RamoRep from ra_carga where"
	str = str & " codramo='" & Trim(CodRamo) & "'"
	str = str & " and codcli = '" & trim(Codcli) & "'"
	str = str & " and ano=" & Ano & ""
	str = str & " and periodo=" & Periodo & ""
	If BCL_ADO(str, rst) Then
		GetRamoFueraMallaRep= trim(ValNulo(rst("ramorep"), str_))
	else
		GetRamoFueraMallaRep = ""
	End If
	
End Function

%>