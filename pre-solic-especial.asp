<!--#INCLUDE file="include/conexion.inc" -->
<!--#INCLUDE file="include/funciones.inc" -->
<!--#INCLUDE file="include/periodos.inc" -->
<%
Session("InscripcionDebeRealizada") = "N"
%>
		<script>
		//alert("Se esta probando inscripcion sin validaciones. pre-inscrip-asigna.asp");
	    //window.location.href='inscrip-asigna.asp';
		</script>
        
<%
if Session("CodCli") = "" then
 ' Response.Redirect("saltoinicio.htm") %>
		<script>
			//alert("Tienmpo de Session Expirado");
			window.top.location.href='saltoinicio.htm';
		</script> 
<%
end if
%>
        
<%

	Mensaje = ""
	Estado = 0
	Session("InscripcionDebeRealizada") = "N"
	
'Valida si el Alumno se encuentra Matriculado

	Matriculado = Session("EstadoMatriculado")

	if Matriculado = "N" then 
    	Mensaje = "- Usted DEBE Matricularse. " 
		Estado =  1
		%>
		<script>
		//alert("Usted DEBE matricularse");
		//window.top.location.href='menu_tomaderamos.asp';
		</script>
	<%
	End if 
	%>
	<%

' Valida si el alumno tiene deuda

	IdRut= Session("Rut")
	Sstr=" exec sp_alumno_deuda_ra '"& (IdRut) &"'"
	Session("Conn").Execute Sstr

	Sstr=" Select Deuda from mt_client Where Codcli = '"& (IdRut) &"'"
	Set Rrst = Session("Conn").Execute(Sstr)
	Bloqueo=valnulo(Rrst("DEUDA"), num_)
	
	if Bloqueo <> 0 then 
		Mensaje = Mensaje & " - Usted DEBE regularizar su Deuda Financiera. " 
		Estado = 1
	%>
		<script>
		//alert("Usted DEBE regularizar su deuda financiera");
		//window.top.location.href='menu_tomaderamos.asp';
		</script>
	<%
	End if
	Rrst.close()
	%>
    <%
	
' Valida Docuemntos Pendientes.
	strV = "SELECT COALESCE(VALIDADOCPEND,'N') AS VALIDADOCPEND FROM mt_carrer where codcarr = '" & session("Codcarr") & "'"
	'response.Write(strV)
	'response.End()
	Set RstV = Session("Conn").Execute(strV)
	if not RstV.eof then 
	if RstV("VALIDADOCPEND") = "S" then


	str = "select b.descripcion from mt_certifalu a, mt_certif b, mt_alumno c"
	str = str & " where a.codcli=c.rut and c.rut='" & IdRut & "'"
	str = str & " and a.codcertif = b.codcertif and entregado='N'"
	str = str & " and a.codcertif in ('1','2','4') "
	str = str & " and c.estacad  in ('egresado','vigente')"

	'response.Write(str)
	'response.end()
	Set Rst1 = Session("Conn").Execute(str)
		
	if not Rst1.eof  then 
		Mensaje= Mensaje & " - Usted DEBE entregar Documentos Pendientes. " 
		Estado = 1
	 %>
	<script>
		//alert("Usted DEBE entregar documentos pendientes");
		//window.top.location.href='menu_tomaderamos.asp';
	</script>	
	<%	
    End if 
	Rst1.close()
	
	
	End if
	End if
	RstV.close()
	%>
    
    <%
' Valida Si el alumno Contesto la Evaluacion Docente.
	vrmensaje = 0
	Codcli = session("Codcli")
	
	sql = "select distinct codcli from  ra_nota  where ano = '"& session("ano_ed") &"' and periodo = '"& session("periodo_ed") &"' and CODCLI = '" & Codcli & "'"
	
	'response.Write(sql)
	'response.End()
	Set Rst4 = Session("Conn").Execute(sql)
		if not Rst4.eof then 

			if Encuestado (session("Codcli"),session("ano_ed"),session("periodo_ed")) = 0 or Encuestado (session("Codcli"),session("ano_ed"),session("periodo_ed"))= 2  then
			
						if  Encuestado (session("Codcli"),session("ano_ed"),session("periodo_ed")) = 0 then 
						   'Mensaje = Mensaje & " - Usted No a Realizado la Evaluacion Docente. "
							Estado = 2
							vrmensaje = 1
							
						 %>
							<script>
							//	alert("Usted No a Realizado Encuestas");
							</script>
						 <% elseif Encuestado (session("Codcli"),session("ano_ed"),session("periodo_ed")) = 2 then 
							'Mensaje = Mensaje & " - Alumno no tiene Carga Academica, NO podra realizar Evaluacion Docente. "
							Estado = 2 
							vrmensaje = 1
							%>
							<script>
							//	alert("Alumno no tiene Carga Academica");
							</script>
						<%
						end if
							if TieneCarga(session("Codcli"),Session("Ano_Ed"),Session("Periodo_Ed")) = 0 then
							'if TieneCarga(session("Codcli"),2010,2) = 0 then
									'if vrmensaje = 0 then 
									'	Mensaje = Mensaje & " - Alumno no tiene Carga Academica, NO podra realizar Evaluacion Docente. "
										vrmensaje = 1
									'end if 
									Estado = 2
								%> 
									<script>
										//alert("Alumno no tiene Carga Academica, wcwec");
									</script>
									<%	if GetHabilitaEncuesta(GetCarreraAlumno(session("Codcli"))) = 0 then 
										'if vrmensaje = 0 then 
											'Mensaje = Mensaje & " - Alumno no tiene Carga Academica, NO podra realizar Evaluacion Docente. "
											vrmensaje = 1
										'end if 	
										Estado = 2
									%> 
										<script>
											//alert("- Alumno no tiene Carga Academica, NO podra realizar Encuesta Docente. ");
										</script>
										
													<%	'if TieneNotas(session("Codcli"),Session("Ano_Ed"),Session("Periodo_Ed")) <> 0 then  
														if TieneNotas(session("Codcli"),2010,2) <> 0 then
														'Mensaje = Mensaje & " - Alumno tiene notas pendientes, NO podra realizar Evaluacion Docente. "
														Estado = 2
														vrmensaje = 1
													%>
															
															<script>
															  // alert("Usted DEBE Realizar la Pauta de Evaluacion Docente");
																//window.top.location.href='menu_encuestas.asp';
														   </script>
													  <%else 
														'Estado = 0
														Estado= 0
													  %>     
						   
			
			<% 											End if 
										End if 
							End if 
			End if 
		
		End if 
		Rst4.Close()
		if vrmensaje = 1 Then
			Mensaje = Mensaje & " - Usted DEBE Realizar la Evaluacion Docente.  "
		
		end if 
	
' Valida si el Alumno realizo Encuesta de Satisfaccion.

	Codcli = Session("Codcli")
		
	sql34 = "select distinct codcli from  ra_nota  where ano = 2010 and periodo = 2 and CODCLI = '" & Codcli & "'"
	
	Set Rst34 = Session("Conn").Execute(sql34)
	 'NO EXISTE ENCUESTA DOCENTE AUN ESTO QUEDA PENDIENTE
		if not Rst34.eof  and   1 = 2	then 		
			sql1 = "select * from esac_encuestas e, mt_alumno a, ra_nota n"
			sql1 = sql1 & " where e.ano = '2010' "
			sql1 = sql1 & " and e.periodo = '2' "
			sql1 = sql1 & " and e.rut = '" & IdRut & "'" 
			sql1 = sql1 & " and CONVERT (varchar (20),e.rut) = a.rut "
			sql1 = sql1 & " and a.codcli = n.codcli "
			sql1 = sql1 & " and n.codcli = (select distinct n.codcli from  ra_nota n where n.ano = 2010 and n.periodo = 2 and n.CODCLI = '" & Codcli & "')"
				
			Set Rstq = Session("Conn").Execute(sql1)
				
				if rstq.eof then 
					Mensaje = Mensaje & " - Usted DEBE Realizar Encuesta de Satisfaccion " 
					Estado = 2 
					 %> 
					<script>
						//alert("Usted DEBE Realizar Encuesta de Satisfaccion");
						//window.top.location.href='menu_encuestas.asp';
					</script>	
						
		<%		End If	
			 Rstq.Close()
        End if 
        Rst34.close()
        %>
    
		<%
	If Estado = 0 then 
		%>
		<script>
		//alert("<%=Mensaje%>");
		window.location.href='esperainf.htm';
		</script>
    <% elseif Estado = 2 then %>
		<script>
		alert("<%=Mensaje%>");
		window.top.location.href='menu_encuestas.asp';
		</script>
	<% else %>
		<script>
		alert("<%=Mensaje%>");
		window.top.location.href='menu_tomaderamos.asp';
		</script>
	<%End if %>
