<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%

if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if
dim Opcion1,Opcion2,url1,permiso1,habilitada1,url2,permiso2,habilitada2,Opcion3,url3,permiso3,habilitada3
dim enlaces(4)


Opcion1="488"
Opcion2="489"
Opcion3="802"
Opcion4="840"


strParame="SELECT coalesce(dbo.Fn_ValorParame('MUESTRALINKDESCARGAUMOBILE'),'')Parame,coalesce(dbo.Fn_ValorParame('MoodleUcevalpo'),'')MoodleUcevalpo"
set rstParame= Session("Conn").Execute(strParame)

if not rstParame.eof then
	MuestraLink=rstParame("Parame")
	MoodleUcevalpo = rstParame("MoodleUcevalpo")
else
	MuestraLink=""
	MoodleUcevalpo=""
end if

url1=GetPermisoNW (Opcion1)
habilitada1=EstaHabilitadaNW (Opcion1)
if url1 = "N" then
permiso1 = 0
else
permiso1 = 1
end if


url2=GetPermisoNW (Opcion2)
habilitada2=EstaHabilitadaNW (Opcion2)
if url2 = "N" then
permiso2 = 0
else
permiso2 = 1
end if

url3=GetPermisoNW (Opcion3)
habilitada3=EstaHabilitadaNW (Opcion3)
if url3 = "N" then
permiso3 = 0
else
permiso3 = 1
end if

url4=GetPermisoNW (Opcion4)
habilitada4=EstaHabilitadaNW (Opcion4)
if url4 = "N" then
permiso4 = 0
else
permiso4 = 1
end if


if trim(Session("Logueado")) <> "" then
	if trim(habilitada1)="S" then
	  if Permiso1=1 then
		  enlaces(1)=url1 
		else
		  enlaces(1)="MensajesBloqueos.asp" 
	  end if
	else
		  enlaces(1)= "MensajeBloqueoHabilita.asp"
	end if
	
	
	if trim(habilitada2)="S"  then
	  if Permiso2=1 then
		enlaces(2)=url2
	  else
		enlaces(2)="MensajesBloqueos.asp" 
	  end if
	else
		enlaces(2)= "MensajeBloqueoHabilita.asp"
	end if
	
	'habilitada3="S"
	'Permiso3=1
	'enlaces(3)="EvaluacionOnline.asp"
	
	if trim(habilitada3)="S"  then
	  if Permiso3=1 then
		enlaces(3)=url3
	  else
		enlaces(3)="MensajesBloqueos.asp" 
	  end if
	else
		enlaces(3)= "MensajeBloqueoHabilita.asp"
	end if
	'enlaces(3)="EvaluacionOnline.asp"
	if trim(habilitada4)="S"  then
	  if Permiso4=1 then
		enlaces(4)=url4
	  else
		enlaces(4)="MensajesBloqueos.asp" 
	  end if
	else
		enlaces(4)= "MensajeBloqueoHabilita.asp"
	end if
	
else	
	  enlaces(1)= "MensajeBloqueoNoLog.asp"
end if





%>

<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title><%=Session("NombrePestana")%></title>
<body >
<table border="0" cellpadding="0" cellspacing="0"  align="left">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right" ><% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td valign="top">
			<% CargarTop1()%><% SubMenu()%>
			<table width="100%" border="0" cellspacing="0" cellpadding="20" height="550" bgcolor="#FFFFFF">
			  <tr> 
				<td valign="top"><p><img src="Imagenes/Titulos/T-material_apoyo.gif" width="271" height="38"></p>
				  <table width="800" border="0" cellpadding="0" cellspacing="0" bgcolor="#333333">
					<tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada1)%>>
					  <td width="620" height="100"><a href="<%=enlaces(1)%>">
					    <p class="text-menu">Descarga de Apuntes </p>
					  </a>
					    <p class="text-normal-celdas"><span class="tex-normales">Esta opci&oacute;n te permite descargar 
los apuntes por Asignatura.</span></p></td>
					  <td width="20">&nbsp;</td>
				      <td width="160"><a href="<%=enlaces(1)%>"><img src="Imagenes/Ingreso-de-asistencia-1.jpg" height="100" width="150"></a></td>
					</tr>
					<tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada1)%>>
					  <td height="15">&nbsp;</td>
					  <td height="15">&nbsp;</td>
					  <td height="15">&nbsp;</td>
				    </tr>
					<tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada2)%>>
					  <td><a href="<%=enlaces(2)%>">
					    <p class="text-menu">Avisos</p>
					  </a>
                        <p class="text-normal-celdas"><span class="tex-normales">Esta opci&oacute;n te permite ver los avisos ingresados.</span></p></td>
					  <td>&nbsp;</td>
				      <td><a href="<%=enlaces(2)%>"><img src="Imagenes/Ingreso-de-asistencia-1.jpg" height="100" width="150"></a></td>
					</tr>
                    <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada2)%>>
					  <td height="15">&nbsp;</td>
					  <td height="15">&nbsp;</td>
					  <td height="15">&nbsp;</td>
				    </tr>
                    <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada3)%>> 
					  <td><a href="<%=enlaces(3)%>">
					    <p class="text-menu">Evaluaciones Online</p>
					  </a>
                        <p class="text-normal-celdas"><span class="tex-normales">Esta opci&oacute;n te permite contestar las evaluaciones online.</span></p></td>
					  <td>&nbsp;</td>
				      <td><a href="<%=enlaces(3)%>"><img src="Imagenes/Ingreso-de-asistencia-1.jpg" height="100" width="150"></a></td>
					</tr>
                    <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada3)%>> 
                      <td height="15">&nbsp;</td>
                      <td height="15">&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    
                    <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada4)%>> 
					  <td><a href="<%=enlaces(4)%>">
					    <p class="text-menu" id="lblProgramaAsignaturas">Programas de Asignaturas</p> 
					  </a>
                        <p class="text-normal-celdas"><span class="tex-normales" id="lblProgramaAsignaturasDet">Esta opci&oacute;n te permite ver los programas de las asignaturas.</span></p></td>
					  <td>&nbsp;</td>
				      <td><a href="<%=enlaces(4)%>"><img src="Imagenes/Ingreso-de-asistencia-1.jpg" height="100" width="150"></a></td>
					</tr>
                    <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada4)%>> 
                      <td height="15">&nbsp;</td>
                      <td height="15">&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    
                    <%if MoodleUcevalpo="SI" then
					
					if ObtieneEstadoMoodleUCV <> 10 then   
						session("MensajeBloqueosVarios") ="No puede acceder a esta opci&oacute;n."
						paginaMoodle="MensajeBloqueo.asp"
						target =""
					else
						paginaMoodle="accesoMoodle.asp"
						target ="target='_blank'"
					end if 
					
					%>
                    <tr valign="top" bgcolor="#FFFFFF">
					  <td height="15">&nbsp;</td>
					  <td height="15">&nbsp;</td>
					  <td height="15">&nbsp;</td>
				    </tr>
					<tr valign="top" bgcolor="#FFFFFF">
					  <td><a href="<%=paginaMoodle%>" <%=target%> > 
					    <p class="text-menu">Moodle</p>
					  </a>
                        <p class="text-normal-celdas"><span class="tex-normales">Este link te permite acceder a la plataforma moodle.</span></p></td>
					  <td>&nbsp;</td>
				      <td>&nbsp;</td>
					</tr>
                    <%end if %>
                    
                    <%if MuestraLink="SI" then%>
                    <tr valign="top" bgcolor="#FFFFFF">
					  <td height="15">&nbsp;</td>
					  <td height="15">&nbsp;</td>
					  <td height="15">&nbsp;</td>
				    </tr>
					<tr valign="top" bgcolor="#FFFFFF">
					  <td><a href="Aplicacion/UMobile.apk"> 
					    <p class="text-menu">Descargar Aplicaci&oacute;n</p>
					  </a>
                        <p class="text-normal-celdas"><span class="tex-normales">Este link te permite descargar la aplicaci&oacute;n UMobile para tu celular en donde podr&aacute;s revisar tu situaci&oacute;n financiera, cursos, mensajes y m&aacute;s.</span></p></td>
					  <td>&nbsp;</td>
				      <td>&nbsp;</td>
					</tr>
                    <%end if %>
                    
                    <%
					strLink="sp_traeFwlinkPortalAlumno 'menu_material.asp'"
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
<%ObjetosLocalizacion("menu_material.asp")%>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->