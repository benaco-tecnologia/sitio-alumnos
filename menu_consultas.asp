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

strParame="SELECT coalesce(dbo.Fn_ValorParame('CERTIFICADOSENMENUCONSULTASPA'),'NO')Parame"
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
	CERTIFICADOSENMENUCONSULTASPA=rstParame("Parame")
end if

carrera=Session("codcarr")
estado=session("estado")
EstadoMatriculado=session("EstadoMatriculado")
bloqueoF=session("BloqueoF")
bloqueoB=session("BloqueoB")
bloqueoA=session("BloqueoA")

Opcion1="INT_001"
Opcion2="INT_002"
Opcion3="INT_003"
Opcion4="484"
Opcion5="467"
Opcion6="476"
Opcion7="483"
Opcion8="477"
Opcion9="480"
Opcion10="481"
Opcion11="482"
Opcion12="760" 'PAGARE
Opcion13="761" 'CUPON MATRICULA
Opcion14="762" 'CUPON ARANCEL
Opcion15="800" 'CONTENIDOS POR CLASE
Opcion16="763" 'Consulta arancel y matricula
Opcion17="896" 'Detalle de asistencia
Opcion18="898" 'Ingreso de solicitudes
Opcion19="899" 'Listado de solicitudes
Opcion20="721" 'Certificados Online

Nivel=0
GetPeriodoActivo
dim enlaces(20)

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
	   permiso5=1'GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion5,nivel,EstadoMatriculado,bloqueoA)
	   permiso6=1'GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion6,nivel,EstadoMatriculado,bloqueoA)
	   permiso7=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion7,nivel,EstadoMatriculado,bloqueoA)
	   permiso8=1'GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion8,nivel,EstadoMatriculado,bloqueoA)
	   permiso9=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion9,nivel,EstadoMatriculado,bloqueoA)
	   permiso10=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion10,nivel,EstadoMatriculado,bloqueoA)
	   permiso11=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion11,nivel,EstadoMatriculado,bloqueoA)
	   
	   habilitada1=EstaHabilitada (Opcion1)
	   habilitada2=EstaHabilitada (Opcion2)
	   habilitada3=EstaHabilitada (Opcion3)
	   
	   
	   'habilitada4=EstaHabilitada (Opcion4)
	   	   url4=GetPermisoNW (Opcion4)
	   habilitada4=EstaHabilitadaNW (Opcion4)
	   if url4 = "N" then
	   permiso4 = 0
	   else
	   permiso4 = 1
	   end if
	   
	   
	   'habilitada5=EstaHabilitada (Opcion5)
	   url5=GetPermisoNW (Opcion5)
	   habilitada5=EstaHabilitadaNW (Opcion5)
	   if url5 = "N" then
	   permiso5 = 0
	   else
	   permiso5 = 1
	   end if
	   
	   
	   'habilitada6=EstaHabilitada (Opcion6)
	   url6=GetPermisoNW (Opcion6)	
	   habilitada6=EstaHabilitadaNW (Opcion6)  
	   if url6 = "N" then
	   permiso6 = 0
	   else
	   permiso6 = 1
	   end if
	   
	   
	   'habilitada7=EstaHabilitada (Opcion7)
	   url7=GetPermisoNW (Opcion7)	
	   habilitada7=EstaHabilitadaNW (Opcion7)  
	   if url7 = "N" then
	   permiso7 = 0
	   else
	   permiso7 = 1
	   end if
	   
	   
	   
	   
	   'habilitada8=EstaHabilitada (Opcion8)
	   url8=GetPermisoNW (Opcion8)	 
	   habilitada8=EstaHabilitadaNW (Opcion8)   
	   if url8 = "N" then
	   permiso8 = 0
	   else
	   permiso8 = 1
	   end if
	   
	   'habilitada9=EstaHabilitada (Opcion9)
	   url9=GetPermisoNW (Opcion9)	 
	   habilitada9=EstaHabilitadaNW (Opcion9)   
	   if url9 = "N" then
	   permiso9 = 0
	   else
	   permiso9 = 1
	   end if
	   
	   
	   'habilitada10=EstaHabilitada (Opcion10)
	   url10=GetPermisoNW (Opcion10)	 
	   habilitada10=EstaHabilitadaNW (Opcion10)   
	   if url10 = "N" then
	   permiso10 = 0
	   else
	   permiso10 = 1
	   end if
	   
	   
	   
	   'habilitada11=EstaHabilitada (Opcion11)
	   url11=GetPermisoNW (Opcion11)	 
	   habilitada11=EstaHabilitadaNW (Opcion11)   
	   if url11 = "N" then
	   permiso11 = 0
	   else
	   permiso11 = 1
	   end if
	   
	   url12=GetPermisoNW (Opcion12)	 
	   habilitada12=EstaHabilitadaNW (Opcion12)   
	   if url12 = "N" then
	   		permiso12 = 0
	   else
	   		permiso12 = 1
	   end if
	   
	   url13=GetPermisoNW (Opcion13)	 
	   habilitada13=EstaHabilitadaNW (Opcion13)   
	   if url13 = "N" then
	   		permiso13 = 0
	   else
	   		permiso13 = 1
	   end if
	   
	   url14=GetPermisoNW (Opcion14)	 
	   habilitada14=EstaHabilitadaNW (Opcion14)   
	   if url14 = "N" then
	   		permiso14 = 0
	   else
	   		permiso14 = 1
	   end if
	   
	   url15=GetPermisoNW (Opcion15)	 
	   habilitada15=EstaHabilitadaNW (Opcion15)   
	   if url15 = "N" then
	   		permiso15 = 0
	   else
	   		permiso15 = 1
	   end if
	   
	   url16=GetPermisoNW (Opcion16)	 
	   habilitada16=EstaHabilitadaNW (Opcion16)   
	   if url16 = "N" then
	   		permiso16 = 0
	   else
	   		permiso16 = 1
	   end if
	   
	   url17=GetPermisoNW (Opcion17)	 
	   habilitada17=EstaHabilitadaNW (Opcion17)   
	   if url17 = "N" then
	   		permiso17 = 0
	   else
	   		permiso17 = 1
	   end if
	  
	  
	   url18=GetPermisoNW (Opcion18)	 
	   habilitada18=EstaHabilitadaNW (Opcion18)   
	   if url18 = "N" then
	   		permiso18 = 0
	   else
	   		permiso18 = 1
	   end if
	   
	   url19 = GetPermisoNW (Opcion19)	 
	   habilitada19 = EstaHabilitadaNW (Opcion19)   
	   if url19 = "N" then
	   		permiso19 = 0
	   else
	   		permiso19 = 1
	   end if
	   
	   url20 = GetPermisoNW (Opcion20)	 
	   habilitada20 = "S"
	   if url20 = "N" then
	   		permiso20 = 0
	   else
	   		permiso20 = 1
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
		enlaces(3)=url4'"resultado.asp" 
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
		enlaces(6)=url7'"certificado.asp"
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
		enlaces(9)=url8'"concent-notas.asp"
	   else
		enlaces(9)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(9)= "MensajeBloqueoHabilita.asp"
	end if
	
	if trim(habilitada9)="S"    then
	  if Permiso9=1 then
	  enlaces(10)=url9'"ed.htm"
	   else
	  enlaces(10)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(10)= "MensajeBloqueoHabilita.asp"
	end if

	if trim(habilitada10)="S"    then
	  if Permiso10=1 then
	  enlaces(11)=url10'"Actualizador.asp"
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
	
	if trim(habilitada12)="S"    then
	  if Permiso12=1 then
	  enlaces(13)=url12'"pagare.asp"
	   else
	  enlaces(13)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(13)= "MensajeBloqueoHabilita.asp"
	end if	
	
	if trim(habilitada13)="S"    then
	  if Permiso13=1 then
	  enlaces(14)=url13'"cuponmatricula.asp"
	   else 
	  enlaces(14)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(14)= "MensajeBloqueoHabilita.asp"
	end if
	
	if trim(habilitada14)="S"    then
	  if Permiso14=1 then 
	  enlaces(15)=url14'"cuponarancel.asp"
	   else 
	  enlaces(15)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(15)= "MensajeBloqueoHabilita.asp"
	end if	
	
	if trim(habilitada15)="S"    then
	  if Permiso15=1 then 
	  enlaces(16)=url15'"ContenidosSelClase.asp"
	   else 
	  enlaces(16)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(16)= "MensajeBloqueoHabilita.asp"
	end if	
  
  if trim(habilitada17)="S"    then
	  if Permiso17=1 then 
	  enlaces(17)=url17
	   else 
	  enlaces(17)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(17)= "MensajeBloqueoHabilita.asp"
	end if	
	
	if trim(habilitada18)="S"    then
		if Permiso18=1 then 
			enlaces(18)=url18
		else 
			enlaces(18)="MensajesBloqueos.asp" 
		end if
	else
		enlaces(18)= "MensajeBloqueoHabilita.asp"
	end if	
	
	
	if trim(habilitada19)="S"    then
		if Permiso19=1 then 
			enlaces(19)=url19
		else 
			enlaces(19)="MensajesBloqueos.asp" 
		end if
	else
		enlaces(19)= "MensajeBloqueoHabilita.asp"
	end if	
	
	if trim(habilitada20)="S"    then
		if Permiso20=1 then 
			enlaces(20)="RedirectNet.asp?PaginaRedirect="&url20
		else 
			enlaces(20)="MensajesBloqueos.asp" 
		end if
	else
		enlaces(20)= "MensajeBloqueoHabilita.asp"
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


strParame="SELECT dbo.Fn_ValorParame('USACONCENTNOTASPA_CENFOTUR')USACONCENTNOTASPA_CENFOTUR"
if BCL_ADO(strParame, rstParame) then
	USACONCENTNOTASPA_CENFOTUR = rstParame("USACONCENTNOTASPA_CENFOTUR")
else
	USACONCENTNOTASPA_CENFOTUR=""
end if

%>
<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if

%>
<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1"> 
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
                                            <img src="Imagenes/titulos/T-consultas.gif" width="271" height="38"></p>
                                            
                                        <table width="800" border="0" cellpadding="0" cellspacing="0" bgcolor="#333333">
                                        	<%if CERTIFICADOSENMENUCONSULTASPA = "SI" then%>                                      
                                                <tr valign="top" bgcolor="#FFFFFF">
                                                    <td height="100">
                                                       <a href="<%=enlaces(20)%>">
                                                            <p class="text-menu">
                                                                Certificados Online</p>
                                                        </a>
                                                        <p align="justify" class="text-normal-celdas">
                                                            <span class="tex-normales">Te permite comprar certificados con Servipag.</span></p>
                                                    </td>
                                                    <td>&nbsp;
                                                        
                                                    </td>
                                                    <td>
                                                     <a href="<%=enlaces(20)%>">
                                                            <img src="Imagenes/menus/concentracion-notas.GIF" height="100" width="150"></a>
                                                    </td>
                                                </tr>
                                                <tr valign="top" bgcolor="#FFFFFF">
                                                  <td height="15">&nbsp;</td>
                                                  <td height="15">&nbsp;</td>
                                                  <td>&nbsp;</td>
                                                </tr> 
                                            <%end if%> 
                                        
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada5)%>>
                                                <td width="619" height="100">
                                                    <a href="<%=enlaces(4)%>">
                                                        <p class="text-menu">
                                                            Malla Curricular
                                                        </p>
                                                    </a>
                                                    <p align="justify" class="text-normal-celdas">
                                                        <span class="tex-normales">Te permite ver el plan de estudios de tu carrera y las asignaturas
                                                            cursadas. La malla tiene incluida las actividades del proceso de titulaci&oacute;n.</span></p>
                                                </td>
                                                <td width="20">&nbsp;
                                                    
                                                </td>
                                                <td width="161">
                                                    <a href="<%=enlaces(4)%>">
                                                        <img src="Imagenes/menus/malla-curricular.GIF" height="100" width="150"></a>
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada5)%>>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                            </tr>
                                          <!--  <tr valign="top" bgcolor="#FFFFFF">
                                                <td height="100">
                                                    <a href="<%'=enlaces(5)%>">
                                                        <p class="text-menu">
                                                            Ficha Curricular</p>
                                                    </a>
                                                    <p align="justify" class="text-normal-celdas">
                                                        <span class="tex-normales">Te permite ver el plan de estudios de tu carrera y las asignaturas
                                                            cursadas. La malla tiene incluida las actividades del proceso de titulaci&oacute;n
                                                            .</span></p>
                                                </td>
                                                <td>&nbsp;
                                                    
                                                </td>
                                                <td>
                                                    <a href="<%'=enlaces(5)%>">
                                                        <img src="Imagenes/Ingreso-de-asistencia-1.jpg" height="100" width="150"></a>
                                                </td>
                                            </tr> 
                                            <tr valign="top" bgcolor="#FFFFFF">
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                            </tr>-->
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada8)%>>
                                            	<%if USACONCENTNOTASPA_CENFOTUR = "SI" then%>
                                                    <td height="100">
                                                        <a href="<%=enlaces(9)%>">
                                                            <p class="text-menu">
                                                                Hist&oacute;rico de notas
                                                            </p>
                                                        </a>
                                                        <p align="justify" class="text-normal-celdas">
                                                            <span class="tex-normales">Te permite ver tu Hist&oacute;rico de notas y asistencia
                                                                acumulada.</span></p>
                                                    </td>
                                                <%else%>
                                                	                                                    <td height="100">
                                                        <a href="<%=enlaces(9)%>">
                                                            <p class="text-menu">
                                                                Concentraci&oacute;n de Notas
                                                            </p>
                                                        </a>
                                                        <p align="justify" class="text-normal-celdas">
                                                            <span class="tex-normales">Te permite ver tu concentraci&oacute;n de notas y asistencia
                                                                acumulada.</span></p>
                                                    </td>

                                                <%end if%>
                                                <td>&nbsp;
                                                    
                                                </td>
                                                <td>
                                                    <a href="<%=enlaces(9)%>">
                                                        <img src="Imagenes/menus/concentracion-notas.GIF" height="100" width="150"></a>
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada8)%>>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada9)%>>
                                                <td height="100">
                                                    <a href="<%=enlaces(10)%>">
                                                        <p class="text-menu" id="lblSituActual">
                                                            Situaci&oacute;n Actual del Alumno.</p>
                                                    </a>
                                                    <p align="justify" class="text-normal-celdas">
                                                        <span class="tex-normales">Te permite ver tu situaci&oacute;n actual.</span></p>
                                                </td>
                                                <td>&nbsp;
                                                    
                                                </td>
                                                <td>
                                                    <a href="<%=enlaces(10)%>">
                                                        <img src="Imagenes/menus/concentracion-notas.GIF" height="100" width="150"></a>
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada9)%>>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada10)%>>
                                                <td height="100">
                                                    <a href="<%=enlaces(11)%>">
                                                        <p class="text-menu">
                                                            Consulta de Biblioteca.</p>
                                                    </a>
                                                    <p align="justify" class="text-normal-celdas">
                                                        <span class="tex-normales">Te permite ver tu situación de biblioteca.</span></p>
                                                </td>
                                                <td>&nbsp;
                                                    
                                                </td>
                                                <td>
                                                    <a href="<%=enlaces(11)%>">
                                                        <img src="Imagenes/menus/concentracion-notas.GIF" height="100" width="150"></a>
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada10)%>>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                            </tr><%Opcion11="482"
													habilitada11=EstaHabilitadaNW (Opcion11)%>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada11)%>>
                                                <td height="100">
                                                    <a href="<%=enlaces(12)%>">
                                                        <p class="text-menu">
                                                            Consulta Financiera.</p>
                                                    </a>
                                                    <p align="justify" class="text-normal-celdas">
                                                        <span class="tex-normales">Te permite ver si actualmente tienes deuda con la Institución.</span></p>
                                                </td>
                                                <td>&nbsp;
                                                    
                                                </td>
                                                <td>
                                                    <a href="<%=enlaces(12)%>">
                                                        <img src="Imagenes/Ingreso-de-asistencia-1.jpg" height="100" width="150"></a>
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada11)%>>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada7)%>>
                                                <td height="100">
                                                    <a href="<%=enlaces(6)%>">
                                                        <p class="text-menu">
                                                            Consulta Curricular.</p>
                                                    </a>
                                                    <p align="justify" class="text-normal-celdas">
                                                        <span class="tex-normales">Te permite ver tu situaci&oacute;n curricular.</span></p>
                                                </td>
                                                <td>&nbsp;
                                                    
                                                </td>
                                                <td>
                                                    <a href="<%=enlaces(6)%>">
                                                        <img src="Imagenes/menus/concentracion-notas.GIF" height="100" width="150"></a>
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada7)%>>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada16)%>>
                                                <td height="100">
                                                    <a href="<%=enlaces(3)%>">
                                                        <p class="text-menu">
                                                            Consulta Arancel y matr&iacute;cula.</p>
                                                    </a>
                                                    <p align="justify" class="text-normal-celdas">
                                                        <span class="tex-normales">Te permite consultar por el valor de arancel y matr&iacute;cula.</span></p>
                                                </td>
                                                <td>&nbsp;
                                                    
                                                </td>
                                                <td>
                                                    <a href="<%=enlaces(3)%>">
                                                        <img src="Imagenes/menus/concentracion-notas.GIF" height="100" width="150"></a>
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada16)%>>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada15)%>>
                                                <td height="100">
                                                    <a href="<%=enlaces(16)%>">
                                                        <p class="text-menu">
                                                            Consulta Contenidos por Clase.</p>
                                                    </a>
                                                    <p align="justify" class="text-normal-celdas">
                                                        <span class="tex-normales">Te permite consultar los contenidos y avance que tiene cada asignatura.</span></p>
                                                </td>
                                                <td>&nbsp;
                                                    
                                                </td>
                                                <td>
                                                    <a href="<%=enlaces(16)%>">
                                                        <img src="Imagenes/menus/concentracion-notas.GIF" height="100" width="150"></a>
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada15)%>>
                                              <td height="15">&nbsp;</td>
                                              <td height="15">&nbsp;</td>
                                              <td>&nbsp;</td>
                                            </tr>
                                            
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada17)%>>
                                                <td height="100">
                                                   <a href="<%=enlaces(17)%>">
                                                        <p class="text-menu">
                                                            Detalle de asistencia.</p>
                                                    </a>
                                                    <p align="justify" class="text-normal-celdas">
                                                        <span class="tex-normales">Te permite consultar el detalle de asistencia que tiene cada asignatura.</span></p>
                                                </td>
                                                <td>&nbsp;
                                                    
                                                </td>
                                                <td>
                                                    <a href="<%=enlaces(17)%>">
                                                        <img src="Imagenes/detalleasistencia.jpg" height="100" width="150"></a>
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=enlaces(17)%>>
                                              <td height="15">&nbsp;</td>
                                              <td height="15">&nbsp;</td>
                                              <td>&nbsp;</td>
                                            </tr>
                                            
                                             <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada18)%>>
                                                <td height="100">
                                                    <a href="<%=enlaces(18)%>">
                                                        <p class="text-menu">
                                                            Ingreso de solicitudes.</p>
                                                    </a>
                                                    <p align="justify" class="text-normal-celdas">
                                                        <span class="tex-normales">Te permite ingresar solicitudes.</span></p>
                                                </td>
                                                <td>&nbsp;
                                                    
                                                </td>
                                                <td>
                                                    <a href="<%=enlaces(18)%>"><img src="Imagenes/ingresodesolicitudes.jpg" height="100" width="150"></a>
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada18)%>>
                                              <td height="15">&nbsp;</td>
                                              <td height="15">&nbsp;</td>
                                              <td>&nbsp;</td>
                                            </tr> 
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada19)%>>
                                                <td height="100">
                                                    <a href="<%=enlaces(19)%>">
                                                        <p class="text-menu">
                                                            Listado de solicitudes.</p> 
                                                    </a>
                                                    <p align="justify" class="text-normal-celdas">
                                                        <span class="tex-normales">Te permite revisar tus solicitudes.</span></p>
                                                </td>
                                                <td>&nbsp;
                                                    
                                                </td>
                                                <td>
                                                    <a href="<%=enlaces(19)%>"><img src="Imagenes/listadodesolicitudes.jpg" height="100" width="150"></a>
                                                </td>
                                            </tr>
                                            <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada19)%>>
                                              <td height="15">&nbsp;</td>
                                              <td height="15">&nbsp;</td>
                                              <td>&nbsp;</td>
                                            </tr>     
                                            
                                            
                                           <%
											strLink="sp_traeFwlinkPortalAlumno 'menu_consultas.asp'"
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
                                            <tr valign="top" bgcolor="#FFFFFF">
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                                <td height="15">&nbsp;
                                                    
                                                </td>
                                            </tr>
                                            
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
<%ObjetosLocalizacion("menu_consultas.asp")%>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
