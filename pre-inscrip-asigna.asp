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
	'response.Write(Session("InscripcionDebeRealizada"))
	'response.End()
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
	'response.Write(IdRut)
	'response.End()
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
	
	
	'if( 1 = 2 ) then 'En espera que la institución defina los documentos obligatorios 
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
	'end if 'En espera que la institución defina los documentos obligatorios 
	
	end if
	end if 
	RstV.close()
	%>
    
		<%
	If Estado = 0 then 
		%>
		<script>
		//alert("<%=Mensaje%>");
		window.location.href='inscrip-asigna.asp';
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
