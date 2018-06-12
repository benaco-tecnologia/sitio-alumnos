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
Opcion5="485"
Opcion6="486"
Opcion7="INT_007"
Opcion8="INT_008"
Opcion9="INT_009"
Opcion10="INT_010"
Opcion11="INT_011"
Opcion13="721"
Opcion14="762" 'CUPON ARANCEL
Opcion15="763"
Opcion16="903"
Opcion17="905"
Opcion18="913" 'MATWEB


Nivel=0
GetPeriodoActivo
dim enlaces(20)

if session("logueado")="1" then

	   permiso1=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion1,nivel,EstadoMatriculado,bloqueoA)
	   Nivel=1
	   permiso2=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion2,nivel,EstadoMatriculado,bloqueoA)
	
	   permiso3=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion3,nivel,EstadoMatriculado,bloqueoA)
	   permiso4=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion4,nivel,EstadoMatriculado,bloqueoA)
	   permiso5=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion5,nivel,EstadoMatriculado,bloqueoA)
	   permiso6=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion6,nivel,EstadoMatriculado,bloqueoA)
	   permiso7=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion7,nivel,EstadoMatriculado,bloqueoA)
	   permiso8=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion8,nivel,EstadoMatriculado,bloqueoA)
	   permiso9=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion9,nivel,EstadoMatriculado,bloqueoA)
	   permiso10=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion10,nivel,EstadoMatriculado,bloqueoA)
	   permiso11=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion11,nivel,EstadoMatriculado,bloqueoA)
	   
	   habilitada1=EstaHabilitada (Opcion1)
	   habilitada2=EstaHabilitada (Opcion2)
	   habilitada3=EstaHabilitada (Opcion3)
	   habilitada4=EstaHabilitada (Opcion4)
	   'habilitada5=EstaHabilitada (Opcion5)
	   url5=GetPermisoNW (Opcion5)
	   habilitada5=EstaHabilitadaNW (Opcion5)
	   if url5 = "N" then
	   permiso5 = 0
	   else
	   permiso5 = 1
	   end if
	   
'	   habilitada6=EstaHabilitada (Opcion6)
	   url6=GetPermisoNW (Opcion6)
	   habilitada6=EstaHabilitadaNW (Opcion6)
	   if url6 = "N" then
	   permiso6 = 0
	   else
	   permiso6 = 1
	   end if
	   
	   habilitada7=EstaHabilitada (Opcion7)
	   habilitada8=EstaHabilitada (Opcion8)
	   habilitada9=EstaHabilitada (Opcion9)
	   habilitada10=EstaHabilitada (Opcion10)
	   habilitada11=EstaHabilitada (Opcion11)
	   
	    url13=GetPermisoNW (Opcion13)
		habilitada13=EstaHabilitadaNW (Opcion13)
		if url13 = "N" then
		Permiso13 = 0
		else
		Permiso13 = 1
		end if
		
		url14=GetPermisoNW (Opcion14)
		habilitada14=EstaHabilitadaNW (Opcion14)
		if url14 = "N" then
			Permiso14 = 0
		else
			Permiso14 = 1
		end if
		
		url15=GetPermisoNW (Opcion15)
		habilitada15=EstaHabilitadaNW (Opcion15)
		if url15 = "N" then
			Permiso15 = 0
		else
			Permiso15 = 1
		end if
		
		url16=GetPermisoNW (Opcion16)
		habilitada16=EstaHabilitadaNW (Opcion16)
		if url16 = "N" then
			Permiso16 = 0
		else
			Permiso16 = 1
		end if
		
		url17=GetPermisoNW (Opcion17)
		habilitada17=EstaHabilitadaNW (Opcion17)
		if url17 = "N" then
			Permiso17 = 0
		else
			Permiso17 = 1
		end if
		url18=GetPermisoNW (Opcion18)
		habilitada18=EstaHabilitadaNW (Opcion18)
		if url18= "N" then
			Permiso18 = 0
		else
			Permiso18 = 1
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
			  enlaces(1)="inscrip-asigna.htm" 
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
		enlaces(2)="solicitud.htm"
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
		enlaces(4)=url5'"malla-curri.asp"
	  else
		enlaces(4)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(4)= "MensajeBloqueoHabilita.asp"
	end if

	if trim(habilitada6)="S"  then
	 if Permiso6=1 then
		enlaces(5)=url6'"ficha.asp"  
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
	  enlaces(10)="ed.htm"
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
	  enlaces(12)="encuesta_ei.asp"
	   else
	  enlaces(12)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(12)= "MensajeBloqueoHabilita.asp"
	end if	

  if trim(habilitada13)="S"  then
	 if Permiso13=1 then
		enlaces(13)="RedirectNet.asp?PaginaRedirect="&url13
	  else
		enlaces(13)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(13)= "MensajeBloqueoHabilita.asp"
	end if

	if trim(habilitada14)="S"  then
	 if Permiso14=1 then
		enlaces(14)=url14 
	  else
		enlaces(14)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(14)= "MensajeBloqueoHabilita.asp"
	end if
	
  if trim(habilitada15)="S"  then
	 if Permiso15=1 then
		enlaces(15)=url15 
	  else
		enlaces(15)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(15)= "MensajeBloqueoHabilita.asp"
	end if
	
	if trim(habilitada16)="S"  then
	 if Permiso16=1 then
		enlaces(16)="RedirectNet.asp?PaginaRedirect="&url16
	  else
		enlaces(16)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(16)= "MensajeBloqueoHabilita.asp"
	end if
	
	if trim(habilitada17)="S"  then
	
	 if Permiso17=1 then	 
		enlaces(17)="RedirectNet.asp?PaginaRedirect="&url17
	  else
		enlaces(17)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(17)= "MensajeBloqueoHabilita.asp"
	end if
	
	if trim(habilitada18)="S"  then
	 if Permiso18=1 then
		enlaces(18)="RedirectNet.asp?PaginaRedirect="&url18
	  else
		enlaces(18)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(18)= "MensajeBloqueoHabilita.asp"
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
		  enlaces(13)="alumnos.asp"
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
	  enlaces(13)="MensajeBloqueoNoLog.asp"
	  
	 end if
end if

%>
<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if

	sql67 = ""
	sql67 = sql67 & "select ctapagnum, monto from mt_ctapag where ctapag = 4 and codcli ='" & Trim(Session("RutCli")) &"' and ano ='" & year(date()) & "'"
	if BCL_ADO(sql67,rst67) then
		Habilitado = "S"
	else 
		Habilitado = "N"
    end if 

%>
<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title><%=Session("NombrePestana")%></title>
<body>
    <table border="0" cellpadding="0" cellspacing="0" align="left">
        <tr>
            <td>
                <table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
                    <tr>
                        <td colspan="2" valign="top" align="right">
                            <% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
                            %>
                        </td>
                        <td valign="top">
                            <% CargarTop1()%><% SubMenu()%>
                            <table width="100%" border="0" cellspacing="0" cellpadding="15" height="550" bgcolor="#FFFFFF">
                                <tr>
                                    <td valign="top">
                                        <p>
                                            <img src="Imagenes/titulos/T-mis-finanzas.gif" width="271" height="38"></p>
                                        <table width="800" border="0" cellpadding="0" cellspacing="0" bgcolor="#333333">
                                            
                                            <!--modificacion RT-->
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada5)%>>
                                                <td height="100">
                                                    <!-- <a href="Pago-Web-Cuotas.asp"> -->
                                                    <a href="<%=enlaces(4)%>">
                                                        <p class="text-menu">
                                                            Consulta de cuentas</p>
                                                    </a>
                                                    <p align="justify" class="text-normal-celdas">
                                                        <span class="tex-normales">Te permite revisar tus compromisos pendientes.</span></p>
                                                </td>
                                                <td>&nbsp;
                                                    
                                                </td>
                                                <td>
                                                    <a href="<%=enlaces(4)%>">
                                                        <img src="Imagenes/menus/concentracion-notas.GIF" height="100" width="150"></a>
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada6)%>>
                                                <td height="100">
                                                    <!-- <a href="Pago-Web-Cuotas.asp"> -->
                                                    <a href="<%=enlaces(5)%>">
                                                        <p class="text-menu">
                                                            Pago de cuotas</p>
                                                    </a>
                                                    <p align="justify" class="text-normal-celdas">
                                                        <span class="tex-normales">Te permite pagar tus cuotas en Servipag.</span></p>
                                                </td>
                                                <td>&nbsp;
                                                    
                                                </td>
                                                <td>
                                                 <a href="<%=enlaces(5)%>">
                                                  <!--   <a href="Pago-Web-Cuotas.asp">  --> 
                                                   
                                                        <img src="Imagenes/menus/concentracion-notas.GIF" height="100" width="150"></a>
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada13)%>>
                                                <td height="100">
                                                   <a href="<%=enlaces(13)%>">
                                                   <!--<a href="RedirectNet.asp?redirect=Certificados.aspx">-->
                                                        <p class="text-menu">
                                                            Certificados Online</p>
                                                    </a>
                                                    <p align="justify" class="text-normal-celdas">
                                                        <span class="tex-normales">El Alumno en esta opci&oacute;n podr&aacute; comprar certificados de manera online.</span></p>
                                                </td>
                                                <td>&nbsp;
                                                    
                                                </td>
                                                <td>
                                                 <a href="<%=enlaces(13)%>">
                                                        <img src="Imagenes/menus/concentracion-notas.GIF" height="100" width="150"></a>
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada17)%>>
                                                <td height="100">
                                                   <a href="<%=enlaces(17)%>">
                                                   	<p class="text-menu">Historial de Certificados</p>
                                                   </a>
                                                    <p align="justify" class="text-normal-celdas">
                                                        <span class="tex-normales">El Alumno en esta opci&oacute;n podr&aacute; revisar los certificados que ha comprado y obtener los certificados vigentes.</span></p>
                                                </td>
                                                <td>&nbsp;
                                                    
                                                </td>
                                                <td>
                                                 <a href="<%=enlaces(17)%>">
                                                        <img src="Imagenes/menus/concentracion-notas.GIF" height="100" width="150"></a>
                                                </td>
                                            </tr>
                                           <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada14)%>>
                                                <td height="100">
                                                   <a href="<%=enlaces(14)%>">
                                                   <!--<a href="RedirectNet.asp?redirect=Certificados.aspx">-->
                                                        <p class="text-menu">
                                                            Cup&oacute;n para Pago de Cuotas</p>
                                                    </a>
                                                    <p align="justify" class="text-normal-celdas">
                                                        <span class="tex-normales">Te permite imprimir el cup&oacute;n para cancelar tus cuotas.</span></p>
                                                </td>
                                                <td>&nbsp;
                                                    
                                                </td>
                                                <td>
                                                 <a href="<%=enlaces(14)%>">
                                                        <img src="Imagenes/menus/concentracion-notas.GIF" height="100" width="150"></a>
                                                </td>
                                            </tr>
                                            
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada16)%>>
                                                <td height="100">
                                                    <a href="<%=enlaces(16)%>">
                                                        <p class="text-menu">Consulta de boletas</p>
                                                    </a>
                                                    <p align="justify" class="text-normal-celdas">
                                                        <span class="tex-normales">Te permite visualizar los documentos emitidos.</span></p>
                                                </td>
                                                <td>&nbsp;</td>
                                                <td>
                                                    <a href="<%=enlaces(16)%>"><img src="Imagenes/menus/concentracion-notas.GIF" height="100" width="150"></a>
                                                </td>
                                            </tr>
                                            
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada18)%>>
                                                <td height="100">
                                                
                                                <%
													'Validacion Previa P
													strPago="SP_VALIDA_PAGO_CUOTA_BASICA '"& Session("RutCli") &"' ,'"& session("codcarr") &"'"
													if BCL_ADO(strPago, rstPago) then
														PAGO = rstPago("PAGO")
													end if
													
													if PAGO = "NO" and enlaces(18) = "RedirectNet.asp?PaginaRedirect="&url18 then
														session("MensajeBloqueosVarios") ="Para acceder a esta opci&oacute;n debe pagar su cuota b&aacute;sica de matr&iacute;cula."
														enlaces(18) = "MensajeBloqueo.asp"
													end if 
													
													'Validacion calendario
													strCalenMat="SP_GET_CALENDARIO_MATWEB '"& Session("codcli") &"' "
													if BCL_ADO(strCalenMat, rstCalenMat) then
														HAYCALENDARIO = rstCalenMat("HAYCALENDARIO")
													end if
													
													if HAYCALENDARIO = "NO" then
														session("MensajeBloqueosVarios") ="A&uacute;n no se apertura el proceso o ha terminado el periodo de Matr&iacute;cula web."
														enlaces(18) = "MensajeBloqueo.asp"
													end if
													
													'Validaciones Previa Chile
													strVP = "SP_VALIDA_REQUISITOS_MAT_WEB '"& Session("codcli") &"' "
													
													if BCL_ADO(strVP, rstVP) then
														RESULTADO = rstVP("RESULTADO")
													end if
													
													if RESULTADO <> "OK" then
														session("MensajeBloqueosVarios") ="Para acceder a esta opci&oacute;n debe cumplir con los requisitos de matr&iacute;cula."
														enlaces(18) = "MensajeBloqueo.asp"
													end if 
													
												%>
                                                    <a href="<%=enlaces(18)%>">
                                                        <p class="text-menu">Matr&iacute;cula Web</p>
                                                    </a>
                                                    <p align="justify" class="text-normal-celdas">
                                                        <span class="tex-normales">Esta opci&oacute;n permite realizar tu matr&iacute;cula v&iacute;a online.</span></p>
                                                </td>
                                                <td>&nbsp;</td>
                                                <td>
                                                    <a href="<%=enlaces(18)%>"><img src="Imagenes/menus/concentracion-notas.GIF" height="100" width="150"></a>
                                                </td>
                                            </tr>
                                            <%
											strLink="sp_traeFwlinkPortalAlumno 'menu_finanzas.asp'"
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
                                        <p>&nbsp;
                                            </p>
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
