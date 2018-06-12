<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->



<script> 
function abrir() { 
open('popup.html','','top=200,left=200,width=920,height=550') ; 
} 
</script> 

<%
dim carrera,estado,bloqueoF,bloqueoB,Nivel,Permiso,MensajeP,bloqueoA
dim Opcion1,Opcion2,Opcion3,Opcion4,Opcion5,Opcion6,Opcion7,Opcion8,Opcion9,Opcion10
dim permiso1,permiso2,permiso3,permiso4,permiso5,permiso6,permiso7,permiso8,permiso9,permiso10
dim habilitada1,habilitada2,habilitada3,habilitada4,habilitada5,habilitada6,habilitada7,habilitada8,habilitada9,habilitada10
dim url1,url2,url3,url4,url5,url6,url7,url8,url9,url10,url11


strParamePOP="SELECT dbo.Fn_ValorParame('ACTIVAPOPUPPORTALALUMNO')ACTIVAPOPUPPORTALALUMNO,coalesce(dbo.Fn_ValorParame('ACTIVATRAUTO'),'')ACTIVATRAUTO"
set rstParamePOP= Session("Conn").Execute(strParamePOP)

if not rstParamePOP.eof then
	ACTIVAPOPUPPORTALALUMNO=rstParamePOP("ACTIVAPOPUPPORTALALUMNO")
	ACTIVATRAUTO = rstParamePOP("ACTIVATRAUTO")
else
	ACTIVAPOPUPPORTALALUMNO=""
	ACTIVATRAUTO =""
end if 

strParame="SELECT dbo.Fn_ValorParame('ActivaGalyleo')Parame"
set rstParame= Session("Conn").Execute(strParame)

if not rstParame.eof then
	ActivaGalyleo=rstParame("Parame")
else
	ActivaGalyleo=""
end if 

StrA = "select Acepto from SIS_REG_INSCRIPCION WHERE codcli= '" & trim(Session("CodCli")) & "' and ano= '" & trim(Session("ano")) & "' AND periodo = '" & trim(Session("periodo")) & "'"
if bcl_ado(StrA, rstA) then
  if trim(valnulo(rstA("Acepto"),str_))="SI" then
  	AceptoCarga = "SI"
  else
  	AceptoCarga = "NO"
  end if   
else 
  AceptoCarga = "SI"
end if

strAN="select nuevo from mt_alumno where codcli='"& session("codcli") &"'"
if bcl_ado(strAN,rstAN) then
	if ucase(trim(valnulo(rstAN("nuevo"),str_)))="S" then
		ACTIVATRAUTO ="NO"
	end if 
end if

carrera=Session("codcarr")
estado=session("estado")
EstadoMatriculado=session("EstadoMatriculado")
bloqueoF=session("BloqueoF")
bloqueoB=session("BloqueoB")
bloqueoA=session("BloqueoA")



strHS="SELECT solicita from mt_carrer where codcarr='"& carrera &"'"
set rstHs= Session("Conn").Execute(strHS)

if not rstHs.eof then 
	HabilitaSolicitud=rstHs("solicita")
else
	HabilitaSolicitud ="N"
end if 
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


Opcion1="466" 'id de la ca_menu
Opcion2="473"
Opcion3="470"
Opcion4="470"
Opcion5="467"
Opcion6="INT_006"
Opcion7="INT_007"
Opcion8="INT_008"
Opcion9="INT_009"
Opcion10="INT_010"
Opcion11="INT_011"

Nivel=0
GetPeriodoActivo
dim enlaces(12)

AnoAdmi=AnoAdmision()
PeriodoAdmi=PeriodoAdmision()


if session("logueado")="1" then
	   
	   if Session("perfilNW") = 0 then
	   permiso1 = 0
	   else
	   permiso1 = 1
	   end if
       'permiso1=1'GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion1,nivel,EstadoMatriculado,bloqueoA)
	   Nivel=1
	   permiso2=1'GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion2,nivel,EstadoMatriculado,bloqueoA)
	
	   permiso3=1'GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion3,nivel,EstadoMatriculado,bloqueoA)
	   permiso4=1'GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion4,nivel,EstadoMatriculado,bloqueoA)
	   permiso5=1'GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion5,nivel,EstadoMatriculado,bloqueoA)
	   permiso6=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion6,nivel,EstadoMatriculado,bloqueoA)
	   permiso7=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion7,nivel,EstadoMatriculado,bloqueoA)
	   permiso8=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion8,nivel,EstadoMatriculado,bloqueoA)
	   permiso9=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion9,nivel,EstadoMatriculado,bloqueoA)
	   permiso10=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion10,nivel,EstadoMatriculado,bloqueoA)
	   permiso11=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion11,nivel,EstadoMatriculado,bloqueoA)
	   
	   habilitada1=EstaHabilitada (Opcion1)
'	   habilitada2=EstaHabilitada (Opcion2)
	   if EsAlumnoAntiguo(Session("RutCli"),Carrera,AnoAdmi,PeriodoAdmi) = 1 then
	       'habilitada2=EstaHabilitada (Opcion1)
		    url2=GetPermisoNW (Opcion1)
			habilitada2=EstaHabilitadaNW (Opcion1)
			'response.Write(url2+habilitada2)
			'response.End()
			
		   if url2 = "N" then
		   permiso2 = 0
		   else
		   permiso2 = 1
		   end if
		    
		   
	       'habilitada3=EstaHabilitada (Opcion2)
		   url3=GetPermisoNW (Opcion2)
		   habilitada3=EstaHabilitadaNW (Opcion2)
		   if url3 = "N" then
		   permiso3 = 0
		   else
		   permiso3 = 1
		   end if
'		response.write("antiguo")
'		response.end

           else
	       habilitada2="X"
	       habilitada3="X"
'		response.write("nuevo")
'		response.end

	   end if 
		
		'response.Write(habilitada2)
		'response.Write(habilitada3)
		
	   'habilitada3=EstaHabilitada (Opcion3)
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
		
		if trim(habilitada2)="X" then
			enlaces(1)= "MensajeBloqueoHabilitaAlumNuevo.asp"
		elseif trim(habilitada2)="S" then
		  if Permiso2=1 then			  
			  if ACTIVATRAUTO = "SI" then
			      enlaces(1)="resultadotr.asp"			
			  else
				  enlaces(1)=url2				  		  
			  end if 
			  session("url2")=enlaces(1)	
			else
			  enlaces(1)="MensajesBloqueos.asp" 
		  end if

		elseif trim(habilitada2)="N" then

		 	 enlaces(1)= "MensajeBloqueoHabilita.asp"
		else
	
			  enlaces(1)= "MensajeBloqueoHabilitaAlumNuevo.asp"
		end if

	else	
		  enlaces(1)= "MensajeBloqueoNoLog.asp"
	end if

	if trim(habilitada3)="X" then
			enlaces(1)= "MensajeBloqueoHabilitaAlumNuevo.asp"
	elseif trim(habilitada3)="S"  then
	  if Permiso3=1 then
		'enlaces(2)="solicitud.htm"
		if HabilitaSolicitud ="S" then
			enlaces(2)=url3'"frame-solicitud.asp"
		else
			enlaces(2)="MensajesBloqueos.asp"
		end if
	  else
		enlaces(2)="MensajesBloqueos.asp" 
	  end if
	elseif trim(habilitada3)="N" then
		  enlaces(2)= "MensajeBloqueoHabilita.asp"
	else
		  enlaces(2)= "MensajeBloqueoHabilitaAlumNuevo.asp"
	end if

	if trim(habilitada4)="S"  then
	  if Permiso4=1 then
		'enlaces(3)="resultado.htm"
			if ACTIVATRAUTO = "SI" and SESSION("PER_ID") <> "-1" then
				if AceptoCarga = "SI" then
					enlaces(3)="resultado.asp"
				else
					enlaces(3)="resultadotr.asp"		
				end if
		  	else
			  	enlaces(3)=url4	  		  
		  	end if 		
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
	   enlaces(10)="frame-ed.asp"	  
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
if ACTIVATRAUTO <> "SI" then

	if trim(habilitada2)<>""    then
		  if Permiso2=1 then
			  if InscritoNormal() then
				'enlaces(1)="resultado.htm" 
				enlaces(1)="resultado.asp" 
			  end if
		  else
				enlaces(1)="MensajesBloqueos.asp" 
		 end if		  
	 elseif trim(habilitada2)="N"    then
		enlaces(1)= "MensajeBloqueoHabilita.asp"
	 else
	
		enlaces(1)= "MensajeBloqueoHabilitaAlumNuevo.asp"
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

	elseif trim(habilitada3)="N" then
		  enlaces(2)= "MensajeBloqueoHabilita.asp"
	else
		  enlaces(2)= "MensajeBloqueoHabilitaAlumNuevo.asp"
	end if	  
	
end if 	
		 if SESSION("PER_ID") = "-1" then		
			enlaces(1)="sinperiodo.asp" 		
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
%>


<style type="text/css">
<!--
#popup {
	position:absolute;
	left:198px;
	top:80px;
	width:871px;
	height:345px;
	z-index:2000;
	visibility: visible;
}
#popup3 {
	position:absolute;
	left:1115px;
	top:78px;
	width:40px;
	height:40px;
	z-index:2000;
	visibility: visible; 
}
-->
</style>
<script language='javascript'>
<!--

function MM_showHideLayers() { //v9.0
  var i,p,v,obj,args=MM_showHideLayers.arguments;
  for (i=0; i<(args.length-2); i+=3) 
  with (document) if (getElementById && ((obj=getElementById(args[i]))!=null)) { v=args[i+2];
    if (obj.style) { obj=obj.style; v=(v=='show')?'visible':(v=='hide')?'hidden':v; }
    obj.visibility=v; }
}
-->
</script>
<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title><%=Session("NombrePestana")%></title>
<body>

<% if Logueado = "1" then %>
<body onUnload="javascript:window.open('VerifInscrip.asp')" ><%end if%>


<%if ACTIVAPOPUPPORTALALUMNO = "SI" and session("PrimeraVez")<>"1" THEN
	session("PrimeraVez")="1"%>
    <div id="popup">
    	<img src="fotoportal/popup.jpg" width="940" height="635" hspace="0" vspace="0" border="0" usemap="#Map">
    </div>
    <div id="popup3">
    <a href="#" > 
       <img src="imagenes/cerrar.png" width="25" height="25" border="0" usemap="#Map" onClick="MM_showHideLayers('popup3','','hide');MM_showHideLayers('popup','','hide');" /> 
    </a> 
 </div>
<%end if%>


 

<table  width="1024" border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
    <td><table width="200" border="0" cellpadding="0" cellspacing="0" >
      <tr>
        <td colspan="2" valign="top" align="right"><% CargarTop()%>
              <%
			  ' response.Write(enlaces(1))
			  ' response.End()
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
        <td valign="top" ><% CargarTop1()%><% SubMenu()%>
              <table width="100%" border="0" cellspacing="0" cellpadding="20" height="550" bgcolor="#FFFFFF" >
                <tr >
                  <td valign="top"><p><span style="font-size: 14px"><p id ="lblTituloInscripcion" style="font-size: 25px" class="text-menu">Inscripci&oacute;n de Asignaturas</p></span></p>
                      <table width="800" border="0" cellpadding="0" cellspacing="0" bgcolor="#333333">
                      <%
					    MensajeGlosa=""
					  	strGlosa="sp_traeFwGlosasPortales 'menu_tomaderamos.asp','PORTAL ALUMNO'"
						Set rsGlosa = Session("Conn").execute(strGlosa)
								
						if not rsGlosa.eof then       
							while not rsGlosa.eof
								MensajeGlosa =  MensajeGlosa + rsGlosa("glosa") + "</br>"
							rsGlosa.movenext		
							wend
						end if
					  %>  		
                      <%if MensajeGlosa<> ""then%>
                          <link rel="stylesheet" href="http://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/css/bootstrap.min.css">    
                          <div class="panel panel-primary">
                              <!--<div class="panel-heading">Mensaje</div>-->
                              <div class="panel-body"><p align="justify" style="font-size:12px" class="text-normal-celdas"><span class="tex-normales"><%=MensajeGlosa%></span></p></div>
                          </div>
                      <%end if%>
                      <%if session("m_bloque") <> ""then%>
                          <link rel="stylesheet" href="http://maxcdn.bootstrapcdn.com/bootstrap/3.3.5/css/bootstrap.min.css">    
                          <div class="panel panel-primary">
                              <!--<div class="panel-heading">Mensaje</div>-->                              		
                              <div class="panel-body"><p align="justify" style="font-size:12px" class="text-normal-celdas"><span class="tex-normales"><%=session("m_bloque")%></span></p></div>
                          </div>
                      <%end if%>
                              	                      
					  <p style='color: white; font-size: 12pt'><%=session("perfilNW")%></p>
                          
                        <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada2)%>>
                          <td width="620" height="100" bgcolor="#FFFFFF"><%if trim(habilitada2)="S" then%>
                          
                              <% if Permiso2=1 then  %>
                              <a href="<%=enlaces(1)%>">
                              <p class="text-menu" id="lblInscripcion">Inscripci&oacute;n Normal de Asignaturas</p>
                            </a>
                              <%else%>
                              <a href="MensajesBloqueos.asp" >
                              <p class="text-menu" id="lblInscripcion">Inscripci&oacute;n Normal de Asignaturas </p>
                            </a>
                              <%end if%>
                              <%else%>
                              <% if trim(Session("Logueado")) <> "" then %>
                              <a href="MensajeBloqueoHabilita.asp">
                              <p class="text-menu" id="lblInscripcion">Inscripci&oacute;n Normal de Asignaturas </p>
                            </a>
                              <%else%>
                              <a href="MensajeBloqueoNoLog.asp">
                              <p class="text-menu" id="lblInscripcion">Inscripci&oacute;n Normal de Asignaturas </p>
                            </a>
                              <% end if%>
                              <% end if%>
                              <p align="justify" class="text-normal-celdas"><span id="lblGlosaInscripcion"class="tex-normales">Este 
              proceso te permite realizar tu inscripci&oacute;n normal de asignaturas, 
              en los per&iacute;odos fijados para ello.</span></p></td>
                          <td width="20">&nbsp;</td>
                          <td width="160">
						  	  <%if trim(habilitada2)="S" then%>
                              <% if Permiso2=1 then%>
                              <a href="<%=enlaces(1)%>">
                              <img src="Imagenes/menus/inscrip-asignaturas.GIF" height="100" width="150">                            </a>
                              <%else%>
                              <a href="MensajesBloqueos.asp" >
                              <img src="Imagenes/menus/inscrip-asignaturas.GIF" height="100" width="150">                            </a>
                              <%end if%>
                              <%else%>
                              <% if trim(Session("Logueado")) <> "" then %>
                              <a href="MensajeBloqueoHabilita.asp">
                              <img src="Imagenes/menus/inscrip-asignaturas.GIF" height="100" width="150">                            </a>
                              <%else%>
                              <a href="MensajeBloqueoNoLog.asp">
                             <img src="Imagenes/menus/inscrip-asignaturas.GIF" height="100" width="150">                            </a>
                              <% end if%>
                              <% end if%>
						  </td>
                        </tr>
                        <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada2)%>>
                          <td height="15">&nbsp;</td>
                          <td height="15">&nbsp;</td>
                          <td>&nbsp;</td>
                        </tr>
                        <%if ActivaGalyleo="SI" then%>
                        <tr valign="top" bgcolor="#FFFFFF">
                          <td height="100">
                           <a target="_blank" href="http://instituto.galyleo.net/login/index.php">                          
                            <p class="text-menu">Ingreso al Portal de Galyleo </p>
                            </a>
                              <p align="justify" class="text-normal-celdas" style="color: #FF0000; font-size: 12px; font-weight: bold" ><span class="tex-normales">Te 
              permite acceder al portal de matem&aacute;ticas. </span></p></td>
                          <td>&nbsp;</td>
						  <td>&nbsp;</td>
                        </tr>
                        <%end if%>
                        <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada3)%>>
                          <td height="100">
                           <a href="<%=enlaces(2)%>">
                          
                            <p id ="lblSolicitudInscripcion" class="text-menu">Solicitud Especial de Inscripci&oacute;n de Asignaturas </p>
                            </a>
                              <p align="justify" class="text-normal-celdas"><span id="lblGlosaSolicitudInscripcion" class="tex-normales">Te 
              permite solicitar una inscripci&oacute;n especial. Debes usar esta 
              opci&oacute;n s&oacute;lo para las asignaturas que no puedas inscribir 
              a trav&eacute;s de la inscripci&oacute;n normal.</span></p></td>
                          <td>&nbsp;</td>
						  <td><a href="<%=enlaces(2)%>"><img src="Imagenes/menus/solicitud-especial.GIF" height="100" width="150"></a></td>
                        </tr>
                        <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada3)%>>
                          <td height="15">&nbsp;</td>
                          <td height="15">&nbsp;</td>
                          <td>&nbsp;</td>
                        </tr>
                        <tr valign="top" bgcolor="#FFFFFF"  <%=MuestraOpcion(habilitada4)%>>
                          <td height="100"><a href="<%=enlaces(3)%>">
                            <p id="lblResumenInscripcion" class="text-menu">Resumen de Inscripci&oacute;n de Asignaturas </p>
                            </a>
                              <p class="text-normal-celdas"><span id="lblGlosaResumenInscripcion" class="tex-normales">Te 
              permite ver el detalle de las asignaturas inscritas en este per&iacute;odo.</span></p></td>
                          <td>&nbsp;</td>
                          <td><a href="<%=enlaces(3)%>"><img src="Imagenes/menus/resumen de inscripcion.GIF" height="100" width="150"></a></td>
                        </tr>
                        <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada4)%>>
                          <td height="15">&nbsp;</td>
                          <td height="15">&nbsp;</td>
                          <td>&nbsp;</td>
                        </tr>
                        <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada5)%>>
                          <td height="100"><a href="<%=enlaces(4)%>">
                            <p class="text-menu">Malla Curricular </p>
                            </a>
                              <p align="justify" class="text-normal-celdas"><span id="lblGlosaMallaCurri" class="tex-normales">Te 
                                permite ver el plan de estudios de tu carrera y las asignaturas 
                                cursadas. La malla tiene incluida las actividades del proceso de 
                                titulaci&oacute;n .</span></p></td>
                          <td>&nbsp;</td>
                          <td><a href="<%=enlaces(4)%>"><img src="Imagenes/menus/malla-curricular.GIF" height="100" width="150"></a></td>				  
                        </tr>
                        <tr valign="top" bgcolor="#FFFFFF" <%=MuestraOpcion(habilitada5)%>>
                          <td height="15">&nbsp;</td>
                          <td height="15">&nbsp;</td>
                          <td>&nbsp;</td>
                        </tr>
                        
					<%
					strLink="sp_traeFwlinkPortalAlumno 'menu_tomaderamos.asp'"
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
                        <tr> </tr>
                        <tr> </tr>
                    </table>
                    <p>&nbsp;</p></td>
                </tr>
            </table></td>
      </tr>
    </table></td>
  </tr>
</table>
</body>
<%
dim rut,clave,mensaje
rut=session("logrut")
clave=session("logclave")
%>
<script languaje="javascript">

function Terminar()
{
<% if session("logueado")="1" then %>
   if (confirm(String.fromCharCode(191)+"Desea abandonar el Portal de Alumnos?") )
   {
      <% Session("MiValor")=0%>
      window.location = "salir.asp";
        //window.top.location = "alumn-udd.asp";
   }
<% end if%>   
}

//function MensajePermiso(mensaje)
//{
// if mensaje!= "" {
//    alert("Bloqueado ");
//    return;
// }

function Enviar(pag,encrealizada,carga,haynotas,notas)
{


 if (encrealizada==1) {
    alert("Encuesta Realizada");
    return;
 }
 if (carga==0) {
    alert("Alumno no tiene Carga para realizar Encuesta");
    return;
 }

    if (haynotas==1){
	 if (notas==0) {
		alert("Alumno tiene Notas Pendiente.\No puede Realizar la encuesta");
		return;
 	 }
	 }
  	 window.top.location.href = "menu_encuestas.asp";	
}

function Enviar2(pag,valor)
{
 if (valor==0) {
    alert("Debe Realizar o Confirmar Evaluaci\u00f3n Docente"); 
    return;
 }
  window.top.location.href = pag;
}
</script>

<%ObjetosLocalizacion("menu_tomaderamos.asp")%>

</html>


<!--#INCLUDE file="include/desconexion.inc" -->
