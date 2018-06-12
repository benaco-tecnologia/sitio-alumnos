<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->

<%
dim carrera,estado,bloqueoF,bloqueoB,Nivel,Permiso,MensajeP,bloqueoA
dim Opcion1,Opcion2,Opcion3,Opcion4,Opcion5,Opcion6,Opcion7,Opcion8,Opcion9,Opcion10
dim permiso1,permiso2,permiso3,permiso4,permiso5,permiso6,permiso7,permiso8,permiso9,permiso10
dim habilitada1,habilitada2,habilitada3,habilitada4,habilitada5,habilitada6,habilitada7,habilitada8,habilitada9,habilitada10
dim url1,url2,url3,url4,url5,url6,url7,url8,url9,url10,url11


'response.Write("Hola")
'response.End
'response.write(session("RutAlum"))

carrera=Session("codcarr")
estado=session("estado")
EstadoMatriculado=session("EstadoMatriculado")
bloqueoF=session("BloqueoF")
bloqueoB=session("BloqueoB")
bloqueoA=session("BloqueoA")

'response.Write(carrera) & "</br>"
'response.Write("-Estado")
'response.Write(estado) & "</br>"
'response.Write("-BF")
'response.Write(bloqueoF) & "</br>"
'response.Write("-BB")
'response.Write(bloqueoB) & "</br>"
'response.Write("-Ma")
'response.Write(EstadoMatriculado) & "</br>"
'response.Write("-BA")
'response.Write(bloqueoA) & "</br>"


Opcion1="INT_001"
Opcion2="INT_002"
Opcion3="INT_003"
Opcion4="INT_004"
Opcion5="INT_005"
Opcion6="INT_006"
Opcion7="INT_007"
Opcion8="INT_008"
Opcion9="475"
Opcion10="INT_010"
Opcion11="474"

Nivel=0
GetPeriodoActivo
dim enlaces(12)

if session("logueado")="1" then

	   if Session("perfilNW") = 0 then
	   permiso1 = 0
	   else
	   permiso1 = 1
	   end if
	   
	   'permiso1=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion1,nivel,EstadoMatriculado,bloqueoA)
	   Nivel=1
	   permiso2=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion2,nivel,EstadoMatriculado,bloqueoA)
	
	   permiso3=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion3,nivel,EstadoMatriculado,bloqueoA)
	   permiso4=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion4,nivel,EstadoMatriculado,bloqueoA)
	   permiso5=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion5,nivel,EstadoMatriculado,bloqueoA)
	   permiso6=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion6,nivel,EstadoMatriculado,bloqueoA)
	   permiso7=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion7,nivel,EstadoMatriculado,bloqueoA)
	   permiso8=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion8,nivel,EstadoMatriculado,bloqueoA)
	   permiso9=1'GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion9,nivel,EstadoMatriculado,bloqueoA)
	   permiso10=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion10,nivel,EstadoMatriculado,bloqueoA)
	   permiso11=1'GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion11,nivel,EstadoMatriculado,bloqueoA)
	   'response.Write(permiso11)
	   'response.End()
	   habilitada1=EstaHabilitada (Opcion1)
	   habilitada2=EstaHabilitada (Opcion2)
	   habilitada3=EstaHabilitada (Opcion3)
	   habilitada4=EstaHabilitada (Opcion4)
	   habilitada5=EstaHabilitada (Opcion5)
	   habilitada6=EstaHabilitada (Opcion6)
	   habilitada7=EstaHabilitada (Opcion7)
	   habilitada8=EstaHabilitada (Opcion8)
	   'habilitada9=EstaHabilitada (Opcion9)
	   url9=GetPermisoNW (Opcion9)
	   habilitada9=EstaHabilitadaNW (Opcion9)
	   if url9 = "N" then
	   permiso9 = 0
	   else
	   permiso9 = 1
	   end if
	   
	   
	   habilitada10=EstaHabilitada (Opcion10)
	  
	  'habilitada11=EstaHabilitada (Opcion11)
	   url11=GetPermisoNW (Opcion11)
	   habilitada11=EstaHabilitadaNW (Opcion11) 
	   
	   if url11 = "N" then 
	   permiso11 = 0
	   else
	   permiso11 = 1
	   end if
	   
	   

	  if Permiso1=1 then
		  'enlaces(0)="adm-acad.asp"
		  enlaces(0)="menu"
	  else
		  enlaces(0)="MensajesBloqueos.asp" 
	  end if
	
    if trim(Session("Logueado")) <> "" then
		if trim(habilitada2)="S" then
		  if Permiso2=1 then
			  'enlaces(1)="inscrip-asigna.htm" 
			  enlaces(1)="frame-inscrip-asigna.asp" 			  
			else
			  enlaces(1)="MensajesBloqueos.asp" 
		  end if
		else
			  enlaces(1)= "MensajeBloqueoHabilita.asp"
		end if
	else	
		  enlaces(1)= "MensajeBloqueoNoLog.asp"
	end if
	if trim(habilitada3)="S"  then
	  if Permiso3=1 then
		'enlaces(2)="solicitud.htm"
		enlaces(2)="frame-solicitud.asp" 
	  else
		enlaces(2)="MensajesBloqueos.asp" 
	  end if
	else
		enlaces(2)= "MensajeBloqueoHabilita.asp"
	end if

	if trim(habilitada4)="S"  then
	  if Permiso4=1 then
		'enlaces(3)="resultado.htm"
		enlaces(3)="resultado.asp" 
		else
		enlaces(3)="MensajesBloqueos.asp" 
	  end if
	 else
		enlaces(3)= "MensajeBloqueoHabilita.asp"
	end if

	if trim(habilitada5)="S"  then
	  if Permiso5=1 then
		enlaces(4)="malla-curri.asp"
	  else
		enlaces(4)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(4)= "MensajeBloqueoHabilita.asp"
	end if

	if trim(habilitada6)="S"  then
	 if Permiso6=1 then
		enlaces(5)="ficha.asp"  
	  else
		enlaces(5)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(5)= "MensajeBloqueoHabilita.asp"
	end if

	if trim(habilitada7)="S"    then
	  if Permiso7=1 then
		enlaces(6)="certificado.asp"
		else
		enlaces(6)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(6)= "MensajeBloqueoHabilita.asp"
	end if

	if trim(habilitada8)="S"  then  
	  if Permiso8=1 then
		enlaces(7)="cambioclave.asp"
		else
		enlaces(7)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(7)= "MensajeBloqueoHabilita.asp"
	end if

  'TENGO DUDA EN ESTE****************************************************************
  	enlaces(8)="javascript:Terminar()"
	if trim(habilitada8)="S"    then
	 
	   if Permiso8=1 then
		enlaces(9)="concent-notas.asp"
	   else
		enlaces(9)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(9)= "MensajeBloqueoHabilita.asp"
	end if
	
	if trim(habilitada9)="S"    then
	  if Permiso9=1 then
	  'enlaces(10)="ed.htm"
	   enlaces(10)=url9'"frame-ed.asp"	  
	   else
	  enlaces(10)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(10)= "MensajeBloqueoHabilita.asp"
	end if

	if trim(habilitada10)="S"    then
	  if Permiso10=1 then
	  enlaces(11)="Actualizador.asp"
	   else
	  enlaces(11)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(11)= "MensajeBloqueoHabilita.asp"
	end if
	
	if trim(habilitada11)="S"    then
	  if Permiso11=1 then
	  enlaces(12)=url11'"encuesta_ei.asp"
	   else
	  enlaces(12)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(12)= "MensajeBloqueoHabilita.asp"
	end if	
  
  '********************************************TENGO DUDA EN ESTO****************************
  
	if trim(habilitada2)="S"    then
		  if Permiso2=1 then
			  if InscritoNormal() then
				'enlaces(1)="resultado.htm" 
				enlaces(1)="resultado.asp" 
			  end if
		  else
				enlaces(1)="MensajesBloqueos.asp" 
		 end if		  
	 else
		  enlaces(1)= "MensajeBloqueoHabilita.asp"
	end if
				  
  	if trim(habilitada3)="S"    then
		  if Permiso3=1 then
			  if InscritoEspecial() then
				'enlaces(2)="resultado.htm" 
				enlaces(2)="resultado.asp" 
			  end if
		   else
  				enlaces(2)="MensajesBloqueos.asp" 
		 end if
	 else
		  enlaces(2)= "MensajeBloqueoHabilita.asp"
	end if		  
	
		 if SESSION("PER_ID") = "-1" then
			'enlaces(1)="sinperiodo.htm" 
			enlaces(1)="sinperiodo.asp" 
			'enlaces(2)="sinperiodo.htm" 
			enlaces(2)="sinperiodo.asp" 
		 end if
else

	if session("SW")=1 then	
		  enlaces(0)="alumnos.asp"
	  	  enlaces(1)="alumnos.asp"
	  	  enlaces(2)="alumnos.asp"
	  	  enlaces(3)="alumnos.asp"
	  	  enlaces(4)="alumnos.asp"
	  	  enlaces(5)="alumnos.asp"
	  	  enlaces(6)="alumnos.asp"
	  	  enlaces(7)="alumnos.asp"
	  	  enlaces(8)="alumnos.asp"
	  	  enlaces(9)="alumnos.asp"
	  	  enlaces(10)="alumnos.asp"
	  	  enlaces(11)="alumnos.asp"
		  enlaces(12)="alumnos.asp"
	else
	  enlaces(0)="MensajeBloqueoNoLog.asp"
	  enlaces(1)="MensajeBloqueoNoLog.asp"
	  enlaces(2)="MensajeBloqueoNoLog.asp"
	  enlaces(3)="MensajeBloqueoNoLog.asp"
	  enlaces(4)="MensajeBloqueoNoLog.asp"
	  enlaces(5)="MensajeBloqueoNoLog.asp"  
	  enlaces(6)="MensajeBloqueoNoLog.asp"
	  enlaces(7)="MensajeBloqueoNoLog.asp"
	  'enlaces(8)="adm-log.htm"
	  enlaces(8)="alumnos.asp"
	  enlaces(9)="MensajeBloqueoNoLog.asp"
	  enlaces(10)="MensajeBloqueoNoLog.asp"
	  enlaces(11)="MensajeBloqueoNoLog.asp"
	  enlaces(12)="MensajeBloqueoNoLog.asp"
	  
	 end if
end if


'response.write(session("ano_ed")) 
'response.write(session("periodo_ed"))
'response.end()

%>

<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if
%>
<%
dim rut,clave,mensaje
rut=session("logrut")
clave=session("logclave")
%>
<script languaje="javascript">
//function Terminar()
//{
<%' if session("logueado")="1" then %>
   //if (confirm("Desea abandonar?") )
   //{
      <%' Session("MiValor")=0%>
      //window.location = "salir.asp";
        //window.top.location = "alumn-udd.asp";
   //}
<%' end if%>   
//}

//function MensajePermiso(mensaje)
//{
// if mensaje!= "" {
//    alert("Bloqueado ");
//    return;
// }

function Enviar(pag,encrealizada,carga,haynotas,notas)
{


 if (encrealizada==1) {
    alert("Encuesta realizada.");
    return;
 }
 if (encrealizada==3) {
    alert("No existen encuestas vigentes."); 
    return;
 }
 if (encrealizada==4) {
    alert("No tiene asignaturas disponible para encuestar."); 
    return;
 }
 if (carga==0) {
    alert("Alumno no tiene varga para realizar encuesta.");
    return;
 }

    if (haynotas==1){
	 if (notas==0) {
		alert("Alumno tiene Notas Pendiente.\No puede Realizar la encuesta");
		return;
 	 }
	 }
  window.location.href = pag;	
}

function Enviar2(pag,valor)
{
 if (valor==0) {
    alert("Debe Realizar la Encuesta");
    return;
 }
   window.location.href = pag;
}
</script>

<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title><%=Session("NombrePestana")%></title>
<% if Logueado = "1" then %>
<body onUnload="javascript:window.open('VerifInscrip.asp')">
<%end if%>

<table border="0" cellpadding="0" cellspacing="0"  align="left">
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
				
			    %>
                 </td>
			<td valign="top">
			<% CargarTop1()%><% SubMenu()%>
			<table width="100%" border="0" cellspacing="0" cellpadding="20" height="550" bgcolor="#FFFFFF">
           
			  <tr> 
				<td valign="top"><p><img src="Imagenes/titulos/T-encuestas.gif" width="271" height="38"></p>
				  <table width="800" border="0" cellpadding="0" cellspacing="0" bgcolor="#333333">
                 
					<tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada9)%>>
					  <td >
                      
					<%'if habilitada9="S" then
					if Cstr(session("ano_ed")) <> "0" and Cstr(session("periodo_ed")) <> "0" then
						VariableX = Encuestado(session("Codcli"),session("ano_ed"),session("periodo_ed"))
						
						if session("debeEncDoc")="SI" then
							ValorEncuestado = 0
						elseif session("SinRamosEncDoc")="SI" then
							ValorEncuestado = 4 
						else
							sqlenc="select numero from RA_ENCUESTA Where CONVERT(DATE,getDate())>=FecIni And CONVERT(DATE,getDate()) <=FecFin AND TIPO_ENCUESTA=1"				
							if bcl_ado(sqlenc,rstenc) then
								ValorEncuestado = 1
							else
								ValorEncuestado = 3
							end if 
						end if 
						
						if VariableX = 4 then
							ValorEncuestado = VariableX
						end if
					else
						ValorEncuestado = 3
					end if 
					%>
                    
					 <%if Permiso9=1 then %>
					  	<a href="javascript:Enviar('<%=enlaces(10)%>','<%=ValorEncuestado%>','<%=TieneCarga(session("Codcli"),Session("Ano_Ed"),Session("Periodo_Ed"))%>','<%=GetHabilitaEncuesta(GetCarreraAlumno(session("Codcli")))%>','<%=TieneNotas(session("Codcli"),Session("Ano_Ed"),Session("Periodo_Ed"))%>')" >
						<p class="text-menu">Evaluaci&oacute;n Docente </p>
					    </a>
					<%else%>
	   				  	<% if trim(Session("Logueado")) <> "" then %> 
					  		<a href="MensajesBloqueos.asp">
					        <p class="text-menu">Evaluaci&oacute;n Docente </p>
	   					    </a>
					    <%else%>
					  		<a href="MensajeBloqueoNoLog.asp">
					        <p class="text-menu">Evaluaci&oacute;n Docente </p>
					 		</a>
					    <% end if%>
                    	<%end if%>
					 <%'end if%>
					  
					  <p class="text-normal-celdas">Te permite evaluar a tus Docentes.</p></td>
					  <td width="20">&nbsp;</td>
				      <td width="160">
					<%if Permiso9=1 then%>
					  	<a href="javascript:Enviar('<%=enlaces(10)%>','<%=Encuestado(session("Codcli"),session("ano_ed"),session("periodo_ed"))%>','<%=TieneCarga(session("Codcli"),Session("Ano_Ed"),Session("Periodo_Ed"))%>','<%=GetHabilitaEncuesta(GetCarreraAlumno(session("Codcli")))%>','<%=TieneNotas(session("Codcli"),Session("Ano_Ed"),Session("Periodo_Ed"))%>')" >
						<img src="Imagenes/Ingreso-de-asistencia-1.jpg" height="100" width="150">					    </a>
					<%else%>
	   				  	<% if trim(Session("Logueado")) <> "" then %> 
					  		<a href="MensajesBloqueos.asp">
					        <img src="Imagenes/Ingreso-de-asistencia-1.jpg" height="100" width="150">	   					    </a>
					    <%else%>
					  		<a href="MensajeBloqueoNoLog.asp">
					        <img src="Imagenes/Ingreso-de-asistencia-1.jpg" height="100" width="150">					 		</a>
					    <% end if%>
                    <%end if%>					  </td>
					</tr>
					<tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada9)%>>
					  <td height="15">&nbsp;</td>
					  <td height="15">&nbsp;</td>
					  <td height="15">&nbsp;</td>
				    </tr>
					<% Opcion11="474"
					habilitada11=EstaHabilitadaNW (Opcion11)%>
					<tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada11)%>>
					  <td height="100">
					<%if Permiso11=1 then%>
					  	<a href="<%=enlaces(12)%>">
					    	<p class="text-menu">Encuesta de Satisfacci&oacute;n </p>
					  	</a>
					<%else%>	
					  	<% if trim(Session("Logueado")) <> "" then %>
					  		<a href="MensajesBloqueos.asp">
							<p class="text-menu">Encuesta de Satisfacci&oacute;n </p>
							</a>
						<%else%>
							<a href="MensajeBloqueoNoLog.asp">
							<p class="text-menu">Encuesta de Satisfacci&oacute;n </p>
							</a>
						<% end if%>
					<% end if%>	
                        	<p id="lblSubtitEncuestaSatis" class="text-normal-celdas">Encuesta de Satisfacci&oacute;n del Alumno. </p></td>
					  <td width="20">&nbsp;</td>
				      <td width="160">
					  <%if Permiso11=1 then%>
					  	<a href="<%=enlaces(12)%>">
					    	<img src="Imagenes/Ingreso-de-asistencia-1.jpg" height="100" width="150">				  	</a>
					<%else%>	
					  	<% if trim(Session("Logueado")) <> "" then %>
					  		<a href="MensajesBloqueos.asp">
							<img src="Imagenes/Ingreso-de-asistencia-1.jpg" height="100" width="150">							</a>
						<%else%>
							<a href="MensajeBloqueoNoLog.asp">
							<img src="Imagenes/Ingreso-de-asistencia-1.jpg" height="100" width="150">							</a>
						<% end if%>
					<% end if%>					  </td>
					</tr>
					<tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada11)%>> 
					  <td height="15">&nbsp;</td>
					  <td height="15">&nbsp;</td>
					  <td height="15">&nbsp;</td>
				    </tr>
					<tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada9)%>>
					  <td height="100">
                      	<a href="Reglamentos/Pauta%20de%20Evaluacion%20Docente.pdf">
					    <p class="text-menu">Descargar Instructivo Evaluaci&oacute;n Docente. </p></a>
                      <p class="text-normal-celdas">Te permite ver como poder constestar la Encuesta Docente. </p></td>
					  <td width="20">&nbsp;</td>
				      <td width="160"><a href="Reglamentos/Pauta%20de%20Evaluacion%20Docente.pdf"><img src="Imagenes/Ingreso-de-asistencia-1.jpg" height="100" width="150"></a></td>
					</tr>
                    <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(Permiso9)%>>
                      <td height="15">&nbsp;</td>
                      <td height="15">&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <%
					strLink="sp_traeFwlinkPortalAlumno 'menu_encuestas.asp'"
					Set rsLink = Session("Conn").execute(strLink)
							
					if not rsLink.eof then       
						while not rsLink.eof
						%>
						
						 <tr valign="top" bgcolor="#FFFFFF">
						  <td height="100"><a href="<%="RedirectNet.asp?PaginaRedirect="&rsLink("id")%>" target="<%=rsLink("propiedad")%>">
							<p class="text-menu"><%=rsLink("nombre_link")%></p>
							</a>
							  <p align="justify" class="text-normal-celdas"><span class="tex-normales"><%=rsLink("descripcion")%></span></p></td>
							  
						  <td>&nbsp;</td>
						  <%if rsLink("imagen") <> "" then
								RutaFoto = Server.MapPath("Imagenes") & rsLink("imagen")
								if ExisteArchivo(RutaFoto) then%>
									<td><a href="<%="RedirectNet.asp?PaginaRedirect="&rsLink("id")%>" target="<%=rsLink("propiedad")%>"><img src="<%="Imagenes\"&rsLink("imagen")%>" height="100" width="150"></a></td>
								<%else%>
									<td>&nbsp;</td>    
								<%end if%>	
						  <%else%>
							<td>&nbsp;</td>    			  
						  <%end if%>
						</tr>
							
						<%rsLink.movenext		
						wend
					end if %>
			    </table>      
			      <table width="738">
                    <tr>
                    </tr>
                    <tr>
                    </tr>
                  </table>
		        <p>&nbsp;</p></td>
			  </tr>
			</table>
			</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>
</body>

<%ObjetosLocalizacion("menu_encuestas.asp")%>
</html>

<!--#INCLUDE file="include/desconexion.inc" -->
