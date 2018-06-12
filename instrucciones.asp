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
Opcion5="INT_005"
Opcion6="INT_006"
Opcion7="INT_007"
Opcion8="INT_008"
Opcion9="INT_009"
Opcion10="INT_010"
Opcion11="INT_011"

Nivel=0
GetPeriodoActivo
dim enlaces(12)

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
	   habilitada5=EstaHabilitada (Opcion5)
	   habilitada6=EstaHabilitada (Opcion6)
	   habilitada7=EstaHabilitada (Opcion7)
	   habilitada8=EstaHabilitada (Opcion8)
	   habilitada9=EstaHabilitada (Opcion9)
	   habilitada10=EstaHabilitada (Opcion10)
	   habilitada11=EstaHabilitada (Opcion11)

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
<title>Portal de Alumnos en L&iacute;nea</title>
<body >
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
			    %></td>
			<td valign="top">
			<% CargarTop1()%><% SubMenu()%>
			<table width="100%" border="0" cellspacing="0" cellpadding="15" height="550" bgcolor="#FFFFFF">
			  <tr> 
				<td valign="top"><p><img src="Imagenes/titulos/T-consultas.gif" width="271" height="38"></p>
				  <table width="800" border="0" cellpadding="0" cellspacing="0" bgcolor="#333333">
					<tr valign="top" bgcolor="#FFFFFF">
					  <td height="100" colspan="3">
					    <p class="text-menu">Instrucciones.</p>
					    <p align="justify" class="text-normal-celdas">1.- Imprime el pagar&eacute; y Convenio del&nbsp; Cr&eacute;dito  Solidario Universitario<br>
					      2.- Saca fotocopia a tu c&eacute;dula de identidad<br>
					      3.- Firma tu&nbsp; pagar&eacute; y convenio ante  un Notario P&uacute;blico<br>
					      4.- Fotocopia estos documentos(pagar&eacute; y  convenio legalizados)<br>
					      5.- Entrega todos los documentos (pagar&eacute;,  convenio, y fotocopias) en las Oficinas del Fondo Solidario de Cr&eacute;dito.<br>
				        Universitario, ubicada en el campus Macul de la UMCE,( frente a la direcci&oacute;n de b&aacute;sica), en el siguiente horario de  atenci&oacute;n:</p>
                        <p class="text-normal-celdas">&nbsp;&nbsp;&nbsp; &nbsp; Lunes a  Jueves:&nbsp; de 9:00 a 13:.00 Horas y <br>
  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  de 15:00 a 16:30 Horas</p>
                        <p class="text-normal-celdas">&nbsp;&nbsp;&nbsp;  &nbsp;&nbsp;Viernes:&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;de 9:00 a  13:.00 Horas</p>
                        <p class="text-normal-celdas">6.- Plazo de entrega, &nbsp;entre el lunes  02 &nbsp;y el Viernes 13 de Mayo de 2011</p>
                        <p class="text-normal-celdas"><strong><u>NOTA:</u></strong></p>
                        <ul>
                          <li class="text-normal-celdas">&ldquo;Recuerda  la no firma y entrega de tu Pagar&eacute; y convenio dentro del plazo establecido  importa la p&eacute;rdida de tu beneficio&rdquo;.</li>
                          <li class="text-normal-celdas">&ldquo;Los  estudiantes que hayan ingresado a la Universidad este a&ntilde;o 2011 ser&aacute;n convocados  posteriormente.&rdquo;</li>
                          <li class="text-normal-celdas"><strong>&ldquo;Se recomienda utilizar Internet Explorer 7.0, superior o  &nbsp;Mozilla 4&rdquo;</strong>. </li>
                          <p></p>
                          <li class="text-normal-celdas">Cualquier  problema o dificultad por favor contactarse con la Oficina de Normalizaci&oacute;n de  la UMCE </li>
                          <li class="text-normal-celdas">Fonos 2412718-2412550- 2412549- 2412732.</li>
                        </ul></td>
				    </tr>
					<tr valign="top" bgcolor="#FFFFFF">
					  <td width="910" height="15" align="justify"><strong>SI TIENES PROBLEMAS CON LA IMPRESI&Oacute;N DE TU PAGAR&Eacute;, DEBES CONCURRIR A LA OFICINA DE NORMALIZACI&Oacute;N. &nbsp;&nbsp;NO SE RECIBIR&Aacute;N PAGAR&Eacute; MAL IMPRESOS.</strong></td>
					  <td width="20" height="15">&nbsp;</td>
				      <td width="161" height="15">&nbsp;</td>
					</tr>
					<tr valign="top" bgcolor="#FFFFFF">
					  <% if Habilitado = "S" then %>
					  <td height="15"><a target="_blank" href="pagare.asp"><br>
					    <p class="text-menu">Impresi&oacute;n Pagar&eacute;</p>
					    </a>
					    <p align="justify" class="text-normal-celdas">&nbsp;</p></td>
					  <%end if %>
					  <td>&nbsp;</td>
					  <% if Habilitado = "S" then %>
					  <td>&nbsp;</td>
					  <%end if %>
				    </tr>
					<tr valign="top" bgcolor="#FFFFFF">
					  <% if Habilitado = "S" then %>
					  <td height="100"><a target="_blank" href="convenio.asp">
					    <p class="text-menu">Impresi&oacute;n  Convenio.</p>
					    </a>
					    <p align="justify" class="text-normal-celdas">&nbsp;</p></td>
					  <%end if %>
					  <td>&nbsp;</td>
					  <% if Habilitado = "S" then %>
					  <td>&nbsp;</td>
					  <%end if %>
				    </tr>                                                                                
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
</html>
<!--#INCLUDE file="include/desconexion.inc" -->