<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenuExterno.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->

<%
	if Session("CodCli") = "" then
  		Response.Redirect("saltoinicio.htm")
	end if
      
	'Session("InscripcionDebeRealizada") = "N"
	
'Valida si el Alumno se encuentra Matriculado

	Matriculado = Session("EstadoMatriculado")

	if Matriculado = "N" then 
		Mensaje1 = "N"
    	MensajeMat = "* Usted DEBE Matricularse. "
	else 
		Mensaje1 = "S"
	 	MensajeMat = "* Usted se Encuentra Matriculado. "
	End if 

' Valida si el alumno tiene deuda

	IdRut= Session("Rut")
	Sstr=" exec matricula.sp_alumno_deuda_ra '"& (IdRut) &"'"
	Session("Conn").Execute Sstr

	Sstr=" Select Deuda from mt_client Where Codcli = '"& (IdRut) &"'"
	Set Rrst = Session("Conn").Execute(Sstr)
	Bloqueo=valnulo(Rrst("DEUDA"), num_)
	
	if Bloqueo <> 0 then 
		Mensaje2 = "N"
		MensajeDeuda = "* Usted DEBE regularizar su Deuda Financiera. " 
	else
		Mensaje2 = "S"
		MensajeDeuda = "* Usted tiene su Situacion Financiera al Dia. " 
	End if
	Rrst.close()

' Valida Docuemntos Pendientes.

	str = "select b.descripcion from mt_certifalu a, mt_certif b, mt_alumno c"
	str = str & " where a.codcli=c.rut and c.rut='" & IdRut & "'"
	str = str & " and a.codcertif = b.codcertif and entregado='N'"
	str = str & " and a.codcertif in ('1','2','4') "
	str = str & " and c.estacad  in ('egresado','vigente')"

	Set Rst1 = Session("Conn").Execute(str)
		
	if not Rst1.eof  then 
		Mensaje3 = "N"
		MensajeDoctos = "* Usted DEBE entregar Documentos Pendientes. " 
	else
		Mensaje3= "S"
		MensajeDoctos = "* Usted NO tiene Documentos Pendientes por Entregar. " 
    End if 
	Rst1.close()

' Valida Si el alumno Contesto la Evaluacion Docente.

	codcli =  session("Codcli")
	
	sql = "select distinct codcli from  ra_nota  where ano = 2010 and periodo = 2 and CODCLI = '" & Codcli & "'"
	
	Set Rst4 = Session("Conn").Execute(sql)
	
		if not Rst4.eof then 
				
			if Encuestado (session("Codcli"),session("ano_ed"),session("periodo_ed")) = 0 or Encuestado (session("Codcli"),session("ano_ed"),session("periodo_ed"))= 2  then
			
						if  Encuestado (session("Codcli"),session("ano_ed"),session("periodo_ed")) = 0 then 
						   'Mensaje = Mensaje & " - Usted No a Realizado la Evaluacion Docente. "
							Estado = 1
							vrmensaje = 1
							'//	alert("Usted No a Realizado Encuestas");
					    elseif Encuestado (session("Codcli"),session("ano_ed"),session("periodo_ed")) = 2 then 
							'Mensaje = Mensaje & " - Alumno no tiene Carga Academica, NO podra realizar Evaluacion Docente. "
							Estado = 1 
							vrmensaje = 1
							'//	alert("Alumno no tiene Carga Academica");
						end if
							if TieneCarga(session("Codcli"),Session("Ano_Ed"),Session("Periodo_Ed")) = 0 then
							'if TieneCarga(session("Codcli"),2010,2) = 0 then
									'if vrmensaje = 0 then 
									'	Mensaje = Mensaje & " - Alumno no tiene Carga Academica, NO podra realizar Evaluacion Docente. "
										vrmensaje = 1
									'end if 
										'//alert("Alumno no tiene Carga Academica, wcwec");
									if GetHabilitaEncuesta(GetCarreraAlumno(session("Codcli"))) = 0 then 
										'if vrmensaje = 0 then 
											'Mensaje = Mensaje & " - Alumno no tiene Carga Academica, NO podra realizar Evaluacion Docente. "
											vrmensaje = 1
										'end if 	
										
											'//alert("- Alumno no tiene Carga Academica, NO podra realizar Encuesta Docente. ");
											'if TieneNotas(session("Codcli"),Session("Ano_Ed"),Session("Periodo_Ed")) <> 0 then  
														if TieneNotas(session("Codcli"),2010,2) <> 0 then
														'Mensaje = Mensaje & " - Alumno tiene notas pendientes, NO podra realizar Evaluacion Docente. "
															vrmensaje = 1
															
															'  // alert("Usted DEBE Realizar la Pauta de Evaluacion Docente");
															'	//window.top.location.href='menu_encuestas.asp';
														End if 
										End if 
							End if 
			End if 
		
		End if 
		Rst4.Close()
		if vrmensaje = 1 Then
			Mensaje4 ="N"
			MensajeEvaDoc = "* Usted DEBE Realizar la Evaluacion Docente. "
		else
			Mensaje4= "S"
			MensajeEvaDoc = "* Usted Realizo la Evaluacion Docente. "
		end if 
	
' Valida si el Alumno realizo Encuesta de Satisfaccion.

	Codcli = Session("Codcli")
		
	sql34 = "select distinct codcli from  ra_nota  where ano = 2010 and periodo = 2 and CODCLI = '" & Codcli & "'"
	
	Set Rst34 = Session("Conn").Execute(sql34)
	
		if not Rst34.eof then 
				
			sql1 = "select * from esac_encuestas e, mt_alumno a, ra_nota n"
			sql1 = sql1 & " where e.ano = '2010' "
			sql1 = sql1 & " and e.periodo = '2' "
			sql1 = sql1 & " and e.rut = '" & IdRut & "'" 
			sql1 = sql1 & " and CONVERT (varchar (20),e.rut) = a.rut "
			sql1 = sql1 & " and a.codcli = n.codcli "
			sql1 = sql1 & " and n.codcli = (select distinct n.codcli from  ra_nota n where n.ano = 2010 and n.periodo = 2 and n.CODCLI = '" & Codcli & "')"
				
			Set Rstq = Session("Conn").Execute(sql1)
				
				if rstq.eof then
					Mensaje5 = "N" 
					MensajeEncSat= "* Usted DEBE Realizar Encuesta de Satisfaccion. " 
						'//alert("Usted DEBE Realizar Encuesta de Satisfaccion");
						'//window.top.location.href='menu_encuestas.asp';
				else
					Mensaje5 = "S"
					MensajeEncSat= "* Usted Realizo la Encuesta de Satisfaccion. " 
				End If	
			 Rstq.Close()
		else
			Mensaje5 = "N"
			MensajeEncSat= "* Usted DEBE Realizar la Encuesta de Satisfaccion. " 
        End if 		
        Rst34.close()
		
		if session("CarreraAlumno") <> "" then 
		    
			Sql = ""
			Sql =  Sql & " Select ti.descripcion from mt_alumno a , ra_tipositu ti"
			Sql =  Sql & " where ti.codigo = a.tipositu"
			Sql =  Sql & " and codcli ='" & Codcli & "' "
			Sql =  Sql & " and codcarpr ='" & session("CarreraAlumno") & "' "

			Set RstTi = Session("Conn").Execute(Sql)
			if not rstTi.eof then 
				if RstTi("descripcion") <> "" then 
					TipoSitu = RstTi("descripcion")
				end if 
			end if	
			RstTi.Close()		
		
		else
			
			Sql = ""
			Sql =  Sql & " Select ti.descripcion from mt_alumno a , ra_tipositu ti"
			Sql =  Sql & " where ti.codigo = a.tipositu"
			Sql =  Sql & " and codcli ='" & Codcli & "' "

			Set RstTi = Session("Conn").Execute(Sql)
			if not rstTi.eof then 
				if RstTi("descripcion") <> "" then 
					TipoSitu = RstTi("descripcion")
			   	end if 
			end if	
			RstTi.Close()
								
		end if 


Dim Rut, strSql, Estado, GlosaError

' ****** Esta pag revisa la situación por la cual se impidió el logueo al alumno.


rutcompleto=trim((Session("logrut")))

'Rut = RutSinDV(Session("RutAlum"))

codcli=session("codcli")
carreraselecciona =  session("CarreraAlumno")

'response.Write(Session("RutCli"))
'response.Write("HOal")
'response.End()
if Session("logrut") = "" then
  'rutcompleto =(Session("RutCli")
   rutcompleto= Session("RutCliente")
end if 

'response.write("Carrera: ") & session("CarreraAlumno") & " </br>"
'response.End()

if session("CarreraAlumno") <> "" then 
            Sql30 = ""
			Sql30 = "Select estacad, codcarpr from mt_alumno where rut = '" & Session("RutCli")& "'"
			Sql30 =  Sql30 & " and codcarpr ='" & session("CarreraAlumno") & "' "
			Sql30 =  Sql30 & " and codcli ='" & (codcli) & "' "
     		'response.Write(sql30)
			'response.End()
			
			Set Rst30 = Session("Conn").Execute(Sql30)
			if not rst30.eof then 
				if rst30("estacad") <>"" then 
				 GlosaError = rst30("estacad")
				 Codcarr = rst30("codcarpr")
				 'else 
				 'GlosaError = "Año Actual"
			   end if 
			end if
		
			Sql21 = "select AnoAdmision, PeriodoAdmi, MatXPromo from mt_parame"
			set Rst21 = Session("Conn").Execute(Sql21)
			if not Rst21.eof then 
				AnoAdmi = Rst21("Anoadmision")
				PeriodoAdmi = Rst21("PeriodoAdmi")
				paramMat = rst21("MatXPromo")
			end if 
			
else
            Sql30 = "Select estacad, codcarpr from mt_alumno where rut = '" &  Session("RutCli") & "'"
			'Sql30 = sql30 & "matriculado ='S' "
			'response.Write(sql30)
			'response.End()
			
			Set Rst30 = Session("Conn").Execute(Sql30)
			if not rst30.eof then 
				if rst30("estacad") <>"" then 
				 GlosaError = rst30("estacad")
				 Codcarr = rst30("codcarpr")
				 'else 
'				 GlosaError = "Año Actual"
			   end if 
			end if
			
			Sql21 = "select AnoAdmision, PeriodoAdmi, MatXPromo from mt_parame"
			set Rst21 = Session("Conn").Execute(Sql21)
			if not Rst21.eof then 
				AnoAdmi = Rst21("Anoadmision")
				PeriodoAdmi = Rst21("PeriodoAdmi")
				paramMat = rst21("MatXPromo")
			end if 			
			
end if 			
	
	sql9 = ""
	sql9 = sql9 & " select nombre_c from mt_carrer where codcarr = '" & codcarr & "' "
	
	Set RstCar = Session("Conn").Execute(Sql9)
		if not RstCar.eof then 
			Carrera = RstCar("Nombre_c") 
		end if 
	RstCar.Close()

'Datos.Open strSql, Conn
'Estado = Trim(Datos(0))
'Select Case estacad
'	Case "ELIMINADO":
'	     GlosaError = "Ud. se encuentra Eliminado de los registros de la  Universidad."
'	Case "EGRESADO"	 
'	     GlosaError = "Su estado actual es de Egresado en nuestros registros."
'	Case "TITULADO"	 
'	     GlosaError = "Su estado actual es de Titulado en nuestros registros."
'	Case "SUSPENDIDO"	 		
'	     GlosaError = "Ud. se encuentra suspendido de la Universidad."
'	case else  		
'		 GlosaError = "Vigente/Alumno Regular"
'End Select
'Datos.Close()



sql1= ""
sql1= sql1 &"select deuda_biblioteca from mt_client "
sql1= sql1 &"where codcli ='" & Rut & "' and coalesce(deuda_biblioteca,'N')='S'"

if bcl_ado(sql1,rst2)  then
	 deuda = "Usted Tiene deuda Pendiente en Biblioteca, favor regularizar situación"
else
	 deuda ="Usted No tiene Deuda en Biblioteca"
end if   
		
strsql2 = ""
strsql2 =strsql2 & "select b.nombre,a.ctadocnum as Numero,a.fecven,a.monto,a.saldo "
strsql2= strsql2 & "from mt_ctadoc a, mt_docum b,mt_estdoc c, mt_item i, mt_detubi d "
strsql2= strsql2 & "where a.codcli = '" & trim(Session("RutCli")) & "' "
strsql2= strsql2 & "and a.ctadoc = b.tipodoc and a.estado = c.estado  "
strsql2= strsql2 & "and a.item*=i.coditem and a.ubicacion=d.codubifis "
strsql2= strsql2 & "and a.saldo > 0 and convert(varchar,a.fecven,112) < convert(varchar,getdate(),112) and a.estado = 2 order by a.fecven asc "

'response.Write(strsql2)
'response.End()

if bcl_ado(strsql2,rst5)  then
	 dinero = "Usted Tiene Deuda(s) de Arancel, favor regularizar situaci&oacute;n en Finanzas"
else
	 dinero ="Usted No tiene Deudas de Arancel con la Universidad"
end if   

strsql="select distinct a.codsecc,a.ramoequiv AS CODRAMO,c.nombre,a.periodo,a.ano,a.np,a.nf,a.escala ,a.nej,a.ncert1,a.ncert2,a.ncert3,a.ncert4,a.ncert5,a.ncert6,a.ncert7,a.ncert8,a.ncert9,a.ncert10,a.asistencia,a.ne,a.ner,a.nfcontrol"
'strsql=strsql & " from ra_nota a, ra_ramo c,mt_parame p"
strsql=strsql & " from ra_nota a, ra_ramo c, mt_alumno m, ra_curric cu"
strsql=strsql & " where a.codcli='" & (codcli) & "' and "
strsql=strsql & " (a.nf=0 or a.nf is null) and "
strsql=strsql & " a.ramoequiv=c.codramo and " 
strsql=strsql & " a.codcli=m.codcli and  " 
strsql=strsql & " m.codpestud=cu.codpestud and " 
strsql=strsql & " a.codramo=cu.codramo and "
strsql=strsql & " cu.tipoeval = 'N' "  
'strsql=strsql & " b.codramo=c.codramo "
'strsql=strsql & " and a.ano = p.ano"
strsql=strsql & " group by a.codsecc,a.ramoequiv,c.nombre,a.periodo,a.ano,a.np,a.nf,a.escala,a.nej,a.ncert1,a.ncert2,a.ncert3,a.ncert4,a.ncert5,a.ncert6,a.ncert7,a.ncert8,a.ncert9,a.ncert10,a.asistencia,a.ne,ner,a.nfcontrol"
strsql=strsql & " Union select distinct a.codsecc,a.ramoequiv AS CODRAMO,c.nombre,a.periodo,a.ano,a.np,a.nf,a.escala ,a.nej,a.ncert1,a.ncert2,a.ncert3,a.ncert4,a.ncert5,a.ncert6,a.ncert7,a.ncert8,a.ncert9,a.ncert10,a.asistencia,a.ne,a.ner,a.nfcontrol"
'strsql=strsql & " from ra_nota a, ra_ramo c,mt_parame p"
strsql=strsql & " from ra_nota a, ra_ramo c, mt_alumno m, ra_curric cu"
strsql=strsql & " where a.codcli='" & (codcli) & "' and "
strsql=strsql & " (a.estado='' or a.estado is null) and "
strsql=strsql & " a.ramoequiv=c.codramo and " 
strsql=strsql & " a.codcli=m.codcli and  " 
strsql=strsql & " m.codpestud=cu.codpestud and " 
strsql=strsql & " a.codramo=cu.codramo and"
strsql=strsql & " cu.tipoeval = 'C' "  
'strsql=strsql & " b.codramo=c.codramo "
'strsql=strsql & " and a.ano = p.ano"
strsql=strsql & " group by a.codsecc,a.ramoequiv,c.nombre,a.periodo,a.ano,a.np,a.nf,a.escala,a.nej,a.ncert1,a.ncert2,a.ncert3,a.ncert4,a.ncert5,a.ncert6,a.ncert7,a.ncert8,a.ncert9,a.ncert10,a.asistencia,a.ne,ner,a.nfcontrol"
strsql= strsql & " order by a.ano, a.periodo  "

'response.Write(strsql)
'response.End()

'Set Rst = Conn.Execute(StrSql)

if bcl_ado(strsql,Rst)  then
	 notas = "Usted tiene notas pendientes"
else
	 notas ="Usted No tiene notas pendientes" 
end if 
sql20= ""
sql20= sql20 & " select ano from ra_nota"
sql20= sql20 & " where codcli ='" & (codcli) & "' "
sql20= sql20 & " order by ano desc"

'response.Write(sql20)
'response.End()
Set Rst20 = Session("Conn").Execute(Sql20)
if not rst20.eof then 
    if rst20("ano") <>"" then 
	 ano = rst20("ano")
	 else 
	 ano = "Año Actual"
   end if 
end if

%>

<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title>Portal de Alumnos en L&iacute;nea</title>
<body >
<table border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right"><% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td valign="top" >
			<% CargarTop1()%><% SubMenu()%>
			<table width="790" border="0" cellspacing="0" cellpadding="15" bgcolor="#FFFFFF">
			  <tr> 
				<td valign="top"><p><img src="imagenes/titulos/T-consulta-sit-alumno.gif" width="400" height="38"></p>
                <td><div align="right"></div></td>
			  </tr>
                <tr> 
                  <td width="790" height="0" colspan="2" bgcolor="#FFFFFF"><table width="810" border="0" cellpadding="3" cellspacing="2">
                    <tr>
                      <td width="133" height="25" background="Imagenes/fdo-cabecera-cel.jpg" class="text-cabecera-celda" ><strong class="text-cabecera-celda">NOMBRE : </strong></td>
                      <td width="435" height="25" bgcolor="#9FDBE8" class="tex-totales-celda"><font face="Arial, Helvetica, sans-serif" class="tex-totales-celda"><%=Session("NomAlum")%></font></td>
                    </tr>
                    <tr>
                      <td height="25" background="Imagenes/fdo-cabecera-cel.jpg" bgcolor="#D7EAFD" class="text-cabecera-celda" ><strong class="text-cabecera-celda">RUT : </strong></td>
                      <td height="25" bgcolor="#9FDBE8" class="tex-totales-celda"><%=rutcompleto%></td>
                    </tr>
                    <tr>
                      <td height="25" background="Imagenes/fdo-cabecera-cel.jpg" bgcolor="#D7EAFD" class="text-cabecera-celda" ><strong class="text-cabecera-celda">CARRERA : </strong></td>
                      <td height="25" bgcolor="#9FDBE8" class="tex-totales-celda"><%=Carrera%></td>
                    </tr>                    
                    <tr>
                      <td background="Imagenes/fdo-cabecera-cel.jpg" bgcolor="#D7EAFD" class="text-cabecera-celda"><strong class="text-cabecera-celda">&Uacute;LTIMO A&Ntilde;O : </strong></td>
                      <td height="25" bgcolor="#9FDBE8" class="tex-totales-celda"><%=ano%> </td>
                    </tr>
                    
                   	   </table>
                    <table width="582" border="0" height="5">
               		  <tr>
                        		<td width="568" height="10"><strong class="tex-totales-celda">REQUISITOS PARA INSCRIPCION DE TOMA DE RAMOS</strong></td>
                        	</tr>
                       </table>
                       <table width="810" height="48" border="0" cellpadding="3" cellspacing="2" >
                       		<tr class="casillas-form">
                       		 	<td width="810" height="10" bgcolor="#FFFFFF" class="textos" align="justify"><p class="text-normal-celdas">Documentos requisitos entregados, Encuesta de Satisfacci&oacute;n respondida, Evaluaci&oacute;n Docente respondida, estar al d&iacute;a en el Pago de sus Colegiaturas y firmar el Pagar&eacute; Anual (Matr&iacute;cula y Colegiaturas) para el a&ntilde;o 2011. Los casos especiales (causal de eliminaci&oacute;n, reincorporados, congelados,  etc.) ser&aacute;n atendidos por las jefaturas de carrera.</p></td>                            
                            </tr>
                       </table>
                       <br>
                       <table width="810" height="168" border="0" cellpadding="3" cellspacing="2" class="text-normal-celdas">                       
                      		<tr>
							<%	if session("CarreraAlumno") <> "" then %>
                       		 	<td width="562" height="10" bgcolor="#9FDBE8" class="textos"><span class="tex-totales-celda">* SU SITUACI&Oacute;N ACADEMICA ES : 								<%=GlosaError%> - <%=TipoSitu%></span></td>
                   	 	  <%else%>
                            	<td width="562" height="10" bgcolor="#9FDBE8" class="textos"><span class="tex-totales-celda">* SU SITUACI&Oacute;N ACADEMICA ES :        <%=GlosaError%> - <%=TipoSitu%></span></td>
                            <%end if%>
                         </tr>
                      		<tr>
							<%	if mensaje1 = "N" then %>
                       		 	<td width="562" height="10" bgcolor="#9FDBE8"><span class="tex-totales-celda"><%=UCASE(MensajeMat)%></span></td>
                                <td width="562" height="10" bgcolor="#9FDBE8"><span class="tex-totales-celda"><%=UCASE("* Debe dirigirse a Concredito.")%></span></td>
                   	 	  <%else%>
                            	<td width="562" height="10" bgcolor="#9FDBE8"><span class="tex-totales-celda"><%=UCASE(MensajeMat)%></span></td>
                            <%end if%>
                         </tr>
                          	<tr>
                            <% if mensaje2 = "N" then %>
                       			<td bgcolor="#9FDBE8" height="10" class="textos"><span class="tex-totales-celda"><%=UCASE(MensajeDeuda)%></span></td>
                                <td bgcolor="#9FDBE8" height="10" class="textos"><span class="tex-totales-celda"><%=UCASE("* Debe dirigirse a finanzas.")%></span></td>
                          <%else%>
                            	<td bgcolor="#9FDBE8" height="10" class="textos"><span class="tex-totales-celda"><%=UCASE(MensajeDeuda)%></span></td>
                        	<%end if%>
                            </tr>
                       		<tr>
                            <% if mensaje3 = "N" then %>
                        		<td bgcolor="#9FDBE8" height="10" class="textos"><span class="tex-totales-celda"><%=UCASE(MensajeDoctos)%></span></td>
                                <td bgcolor="#9FDBE8" height="10" class="textos"><span class="tex-totales-celda"><%=UCASE("Debe dirigirse a su escuela.")%></span></td>
                   		  <%else%>
                            	<td bgcolor="#9FDBE8" height="10" class="textos"><span class="tex-totales-celda"><%=UCASE(MensajeDoctos)%></span></td>
                            <%end if%>
                           </tr>
                            
                   	</table>
         	      </td>
                </tr>
			</table>
			</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>
</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->