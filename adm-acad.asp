<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/Encuesta.inc" -->

<%  response.buffer = false %>
<%

if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if

%>
<link rel="stylesheet" href="css/tex-normales.css" type="text/css">

<script language="JavaScript" type="text/JavaScript">
<!--
function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}
//-->
</script>
<body background="ima/iso-uhc.gif" onLoad="MM_preloadImages('imag/botones/salir-on.gif','imag/f-der/ficha-on.gif','imag/f-der/concentracion-on.gif','imag/f-der/encuesta-on.gif','imag/f-der/inscripcion-on.gif','ImagenBoton/actualiza-on.gif','ImagenBoton/descarga-on.gif','ImagenBoton/avisos-on.gif')">
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=consulta name=consulta >
</OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=consulta1 name=consulta > </OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=consulta2 name=consulta > </OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=consulta3 name=consulta > </OBJECT>
<%
dim carrera,estado,bloqueoF,bloqueoB,Nivel,PermisoB,MensajeP,msg,bloqueoA
dim Opcion1,Opcion2,Opcion3,Opcion4,Opcion5,Opcion6,Opcion7,Opcion8,Opcion9, Opcion10
dim Habilita1,Habilita2,Habilita3,Habilita4,Habilita5,Habilita6,Habilita7,Habilita8,Habilita9, Habilita10
Dim AnoAdmi
Dim PeriodoAdmi

'dim PermisoA,PermisoB,PermisoC,PermisoD,PermisoE,PermisoF,PermisoG,PermisoH
carrera=Session("codcarr")
RutAlumno=Session("Rut")
estado=session("estado")
bloqueoF=session("BloqueoF")
bloqueoB=session("BloqueoB")
EstadoMatriculado=session("EstadoMatriculado")
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
'response.End()

Opcion1="INT_002"
Opcion2="INT_003"
Opcion3="INT_004"
Opcion4="INT_005"
Opcion5="INT_006"
Opcion6="INT_007"
Opcion7="INT_008"
' NO EXISTE EN BOTONES
Opcion8="INT_009"
Opcion10="INT_010"
Opcion11="INT_011"
'Nivel=0

AnoAdmi=AnoAdmision()
PeriodoAdmi=PeriodoAdmision()
	   
Nivel=1
PermisoB=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion1,nivel,EstadoMatriculado,bloqueoA)
PermisoC=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion2,nivel,EstadoMatriculado,bloqueoA)
PermisoD=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion3,nivel,EstadoMatriculado,bloqueoA)
PermisoE=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion4,nivel,EstadoMatriculado,bloqueoA)
PermisoF=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion5,nivel,EstadoMatriculado,bloqueoA)
PermisoG=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion6,nivel,EstadoMatriculado,bloqueoA)
PermisoH=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion8,nivel,EstadoMatriculado,bloqueoA)
PermisoI=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion10,nivel,EstadoMatriculado,bloqueoA)
PermisoJ=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion11,nivel,EstadoMatriculado,bloqueoA)

if EsAlumnoAntiguo(RutAlumno,Carrera,AnoAdmi,PeriodoAdmi) = 1 then
	habilitada1=EstaHabilitada (Opcion1)
	habilitada2=EstaHabilitada (Opcion2)
else
    habilitada1="X"
    habilitada2="X"
end if 
   
'habilitada1=EstaHabilitada (Opcion1)
'habilitada2=EstaHabilitada (Opcion2)

habilitada3=EstaHabilitada (Opcion3)
habilitada4=EstaHabilitada (Opcion4)
habilitada5=EstaHabilitada (Opcion5)
habilitada6=EstaHabilitada (Opcion6)
habilitada7=EstaHabilitada (Opcion8)
habilitada8=EstaHabilitada (Opcion10)
habilitada_Opcion11=EstaHabilitada (Opcion11)


'*********************************************Segunda Oarte de la Validacion**********************************************
dim strSql,strSql1,strSql2, myStr
dim Rt
ano=session("ano")
codcarr=session("Codcarr")

strSql ="select texto_promocion " &_
        "From promocion_carrera with (nolock)" &_
		"where codcarr='" & codcarr & "' and promocion= '" & ano & "'"

strSql1 ="select texto_promocion " &_
        "From mt_parame with (nolock)"

strSql2 ="select texto_promocion " &_
        "From mt_carrer with (nolock)" &_
		"where codcarr='" & codcarr & "'"

myStr="select distinct b.nombre from mt_carrer a with (nolock), ra_institucion b with (nolock) "
myStr=myStr & " where a.institucion=b.codinst and "
myStr=myStr & " a.codcarr='"& codcarr & "'"

consulta.Open strSql,Conn
consulta1.Open strSql1,Conn
consulta2.Open strSql2,Conn
consulta3.Open myStr,Conn


GetPeriodoActivo
'GetEncuestado
dim enlaces(2)

	enlaces(1) = "inscrip-asigna.htm"
	enlaces(2) = "solicitud.htm"

	if trim(habilitada1)="S"    then
		  if PermisoB=1 then
			  if InscritoNormal() then
				enlaces(1)="resultado.htm" 
			  end if
		  else
				enlaces(1)="MensajesBloqueos.asp" 
		 end if		  
	elseif trim(habilitada1)="N" then
		  enlaces(1)= "MensajeBloqueoHabilita.asp"
	else
		  enlaces(1)= "MensajeBloqueoHabilitaAlumNuevo.asp"
	end if
				  
  	if trim(habilitada2)="S"    then
		  if PermisoC=1 then
			  if InscritoEspecial() then
				enlaces(2)="resultado.htm" 
			  end if
		   else
  				enlaces(2)="MensajesBloqueos.asp" 
		 end if
	elseif trim(habilitada2)="N" then
		  enlaces(2)= "MensajeBloqueoHabilita.asp"
	else
		  enlaces(2)= "MensajeBloqueoHabilitaAlumNuevo.asp"
	end if


'	enlaces(1) = "inscrip-asigna.htm"
'	enlaces(2) = "solicitud.htm"
'	  if InscritoNormal() then
'		enlaces(1)="resultado.htm" 
'	  end if
'	  
'	  if InscritoEspecial() then
'		if PermisoC=1 or PermisoC <> 0 then
'			enlaces(2)="resultado.htm" 
'		   else
'			enlaces(2)="resultado.htm" 
'		end if
'	  end if
	
	  'Response.write("ID = " & SESSION("PER_ID"))
	 ' if SESSION("PER_ID") = "-1" then
	'	enlaces(1)="" 
	'	enlaces(2)="" 
	'  end if
	  
	  'response.Write(enlace(1))
	  'response.End
%>
<html>
<head>
<title>Registro Académico en línea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
<link rel="stylesheet" href="untitled.css" type="text/css">
<link rel="stylesheet" href="css/tex-normales.css" type="text/css">
<script language="JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('imag/f-der/inscripcion-on.gif','imag/f-der/solicitud-on.gif','imag/f-der/malla-on.gif','imag/f-der/ficha-on.gif','imag/f-der/concentracion-on.gif','imag/botones/resumen-on.gif')"> 
<div align="left">
  <table border="0" cellpadding="0" cellspacing="0" height="293" align="left" width="847">
    <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
    <tr> 
      <td><img name="fder_r1_c1" src="imag/f-der/f-der_r1_c1.gif" width="934" height="24" border="0"></td>
    </tr>
    <tr> 
      <td><img name="fder_r2_c1" src="Imagenes/titulo_inicio.gif" width="336" height="28" border="0"><font color="#2166B1"><img src="imag/f-der/no-ayuda.gif" width="619" height="28"></font> 
      </td>
    </tr>
    <tr> 
      <td> <table width="722" height="1" cellpadding="0" cellspacing="0" border="0">
          <tr> 
            <td width="722" valign="top"> <p class="tex-normales">Bienvenido al 
                sistema de inscripci&oacute;n de ramos de <br>
                <%=consulta3(0)%>. A trav&eacute;s de este sistema puedes realizar 
                los procesos de inscripci&oacute;n normal de asignaturas y proceso 
                de solicitudes especiales de inscripci&oacute;n de asignaturas. 
                Adem&aacute;s puedes tener informaci&oacute;n sobre el resultado 
                de tu inscripci&oacute;n de asignaturas y de tu situaci&oacute;n 
                acad&eacute;mica. Si necesitas ayuda por favor selecciona &quot;ayuda 
                animada&quot;, y sigue las instrucciones que se te indican.</p></td>
          </tr>
        </table></td>
    </tr>
    <tr valign="top"> 
      <td height="12"> <table width="584" height="1" cellpadding="0" cellspacing="0" border="0">
          <tr> 
            <td valign="top"> <p>&nbsp;</p></td>
          </tr>
        </table></td>
    </tr>
    <tr valign="top"> 
      <% if not consulta1.eof then%>
      <td height="0" class="tex-normales"><span class="tex-normales"><font face="Arial, Helvetica, sans-serif"><i> 
        <font color="#2661C1"><%=Normaliza(consulta1("texto_promocion"))%></font> 
        </i></font></span></td>
      <%end if %>
    </tr>
    <tr valign="top"> 
      <% if not consulta2.eof then%>
      <td height="0" class="tex-normales"><font face="Arial, Helvetica, sans-serif" color="#2661C1"><%=Normaliza(consulta2("texto_promocion"))%></font></td>
      <%end if%>
    </tr>
    <tr valign="top"> 
      <% if not consulta.eof then%>
      <td height="0" class="tex-normales"><span class="tex-normales"><font face="Arial, Helvetica, sans-serif"><i><font color="#2661C1"><%=Normaliza(consulta("texto_promocion"))%></font></i></font></span></td>
      <%end if%>
    </tr>
    <tr valign="top"> 
      <td height="0">&nbsp;</td>
    </tr>
    <tr valign="top"> 
      <td height="21"> <table width="584" border="0" cellspacing="0" cellpadding="0" height="21">
          <tr> 
		    <%if habilitada1="S" then%>
				<%if PermisoB=1 or PermisoB <> 0 then%>
					<%if trim(enlaces(1)) <> "" then %>
					<td width="175"><a href="javascript:Enviar2('<%=enlaces(1)%>','<%=TomadeRamo(session("Codcarr"),session("Codcli"),session("ano_ed"),session("periodo_ed"))%>')" onMouseOver="MM_swapImage('Image11','','imag/f-der/inscripcion-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/inscripcion-of.gif" name="Image11" width="218" height="20" border="0" id="Image11"></a></td>
					<%else %>
					<td width="175"><img src="imag/f-der/inscripcion-of.gif" width="218" height="20" name="Image1" border="0"></td>
					<%end if%>
				<%else%>
				<td width="175"><a href='<%=enlaces(1)%>' onMouseOver="MM_swapImage('Image11','','imag/f-der/inscripcion-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/inscripcion-of.gif" name="Image11" width="218" height="20" border="0" id="Image11"></a></td>
				<%end if%>
				<%else%>
				<td width="175"><a href='<%=enlaces(1)%>' onMouseOver="MM_swapImage('Image11','','imag/f-der/inscripcion-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/inscripcion-of.gif" name="Image11" width="218" height="20" border="0" id="Image11"></a></td>
				<%end if%>
			
            <td rowspan="2" height="2" width="5">&nbsp;</td>
            <td rowspan="2" width="399" valign="top" class="tex-normales"> Esto 
              proceso te permite realizar tu inscripci&oacute;n normal de asignaturas, 
              en los per&iacute;odos fijados para ello.</td>
          </tr>
          <tr> 
            <td width="175">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr valign="top"> 
      <td height="21"> <table width="584" border="0" cellspacing="0" cellpadding="0" height="21">
          <tr> 
		 <%if habilitada2="S" then%>
				<%if PermisoC=1 or PermisoC <> 0 then%>
					<%if trim(enlaces(2)) <> "" then %>
					<td width="175"><a href="javascript:Enviar2('<%=enlaces(2)%>','<%=TomadeRamo(session("Codcarr"),session("Codcli"),session("ano_ed"),session("periodo_ed"))%>')" onMouseOver="MM_swapImage('Image2','','imag/f-der/solicitud-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/solicitud-of.gif" width="298" height="20" name="Image2" border="0"></a></td>
					<%else%>
					<td width="175"><img src="imag/f-der/solicitud-of.gif" width="298" height="20" name="Image2" border="0"></td>
					<%end if%>
				<%else%>
				<td width="175"><a href='<%=enlaces(2)%>' onMouseOver="MM_swapImage('Image2','','imag/f-der/solicitud-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/solicitud-of.gif" width="298" height="20" name="Image2" border="0"></a></td>
				<%end if%>
		<%else%>
			<td width="175"><a href='<%=enlaces(2)%>' onMouseOver="MM_swapImage('Image2','','imag/f-der/solicitud-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/solicitud-of.gif" width="298" height="20" name="Image2" border="0"></a></td>				
		<%end if%>
            <td rowspan="2" height="2" width="5">&nbsp;</td>
            <td rowspan="2" width="399" valign="top" class="tex-normales"> Te 
              permite solicitar una inscripci&oacute;n especial. Debes usar esta 
              opci&oacute;n s&oacute;lo para las asignaturas que no puedas inscribir 
              a trav&eacute;s de la inscripci&oacute;n normal.</td>
          </tr>
          <tr> 
            <td width="175" valign="top">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr valign="top"> 
      <td height="21"> <table width="584" border="0" cellspacing="0" cellpadding="0" height="21">
          <tr> 
		 <%if habilitada3="S" then%>
				<%if PermisoD=1 or PermisoD <> 0 then %>
					<td width="175"><a href="resultado.htm" onMouseOver="MM_swapImage('Image3','','imag/botones/resumen-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/botones/resumen-of.gif" width="255" height="20" name="Image3" border="0"></a></td>
				<%else%>
					<td width="175"><a href="MensajesBloqueos.asp" onMouseOver="MM_swapImage('Image3','','imag/botones/resumen-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/botones/resumen-of.gif" width="255" height="20" name="Image3" border="0"></a></td>
				<%end if%>
		<%else%>
					<td width="175"><a href="MensajeBloqueoHabilita.asp" onMouseOver="MM_swapImage('Image3','','imag/botones/resumen-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/botones/resumen-of.gif" width="255" height="20" name="Image3" border="0"></a></td>
		<%end if%>
            <td rowspan="2" height="2" width="5">&nbsp;</td>
            <td rowspan="2" width="399" valign="top" class="tex-normales"> Te 
              permite ver el detalle de las asignaturas inscritas en este per&iacute;odo.</td>
          </tr>
          <tr> 
            <td width="175">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr valign="top"> 
      <td height="21"> <table width="584" border="0" cellspacing="0" cellpadding="0" height="15">
          <tr> 
  		 <%if habilitada4="S" then%>
            <%if PermisoE=1 or PermisoE <> 0then %>
	            <td width="105"><a href="malla-curri.asp" onMouseOver="MM_swapImage('Image4','','imag/f-der/malla-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/malla-of.gif" width="114" height="20" name="Image4" border="0"></a></td>
            <%else%>
    	        <td width="105"><a href="MensajesBloqueos.asp" onMouseOver="MM_swapImage('Image4','','imag/f-der/malla-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/malla-of.gif" width="114" height="20" name="Image4" border="0"></a></td>
            <%end if%>
			   
			<%else%>
				 <td width="105"><a href="MensajeBloqueoHabilita.asp" onMouseOver="MM_swapImage('Image4','','imag/f-der/malla-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/malla-of.gif" width="114" height="20" name="Image4" border="0"></a></td>
			<%end if%>						
            <td rowspan="2" height="2" width="5">&nbsp;</td>
            <td rowspan="2" width="451" valign="top" class="tex-normales"> Te 
              permite ver el plan de estudios de tu carrera y las asignaturas 
              cursadas. La malla tiene incluida las actividades del proceso de 
              titulaci&oacute;n .</td>
          </tr>
          <tr> 
            <td width="105">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr valign="top"> 
      <td height="21"> <table width="584" border="0" cellspacing="0" cellpadding="0" height="17">
          <tr> 
  		 <%if habilitada5="S" then%>
            <%if PermisoF=1 or PermisoF <> 0 then %>
            	<td width="93"><a href="ficha.asp" onMouseOver="MM_swapImage('Image53','','imag/f-der/ficha-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/ficha-of.gif" name="Image53" width="114" height="20" border="0" id="Image53"></a></td>
            <%else%>
            	<td width="93"><a href="MensajesBloqueos.asp" onMouseOver="MM_swapImage('Image5','','imag/f-der/ficha-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/ficha-of.gif" width="114" height="20" name="Image5" border="0"></a></td>
            <%end if%>
			<%else%>
				<td width="93"><a href="MensajeBloqueoHabilita.asp" onMouseOver="MM_swapImage('Image5','','imag/f-der/ficha-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/ficha-of.gif" width="114" height="20" name="Image5" border="0"></a></td>
			<%end if%>
            <td rowspan="2" height="2" width="5">&nbsp;</td>
            <td rowspan="2" width="451" valign="top" class="tex-normales"> Te 
              permite ver tu historia acad&eacute;mica.</td>
          </tr>
          <tr> 
            <td width="93">&nbsp;</td>
          </tr>
        </table>
        <table width="584" border="0" cellspacing="0" cellpadding="0" height="17">
          <tr> 
   		 <%if habilitada6="S" then%>
            <%if PermisoG=1 or PermisoG <> 0 then %>
            	<td width="93"><a href="concent-notas.asp" onMouseOver="MM_swapImage('Image51','','imag/f-der/concentracion-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/concentracion-of.gif" name="Image51" width="162" height="20" border="0" id="Image51"></a></td>
            <%else%>
            	<td width="93"><a href="MensajesBloqueos.asp" onMouseOver="MM_swapImage('Image51','','imag/f-der/concentracion-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/concentracion-of.gif" name="Image51" width="162" height="20" border="0" id="Image51"></a></td>
            <%end if%>
			<%else%>
				<td width="93"><a href="MensajeBloqueoHabilita.asp" onMouseOver="MM_swapImage('Image51','','imag/f-der/concentracion-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/concentracion-of.gif" name="Image51" width="162" height="20" border="0" id="Image51"></a></td>
			<%end if%>
            <td rowspan="2" height="2" width="5">&nbsp;</td>
            <td rowspan="2" width="451" valign="top" class="tex-normales"> Te 
              permite ver tu concentraci&oacute;n de notas y asistencia acomulada.</td>
          </tr>
          <tr> 
            <td width="93">&nbsp;</td>
          </tr>
        </table>
        <table width="584" border="0" cellspacing="0" cellpadding="0" height="17">
          <tr> 
            <%if habilitada7="S" then%>
            <%if PermisoH=1 or PermisoH <> 0 then %>
            <td width="93"><a href="javascript:Enviar('ed.htm','<%=Encuestado(session("Codcli"),session("ano_ed"),session("periodo_ed"))%>','<%=TieneCarga(session("Codcli"),Session("Ano_Ed"),Session("Periodo_Ed"))%>','<%=GetHabilitaEncuesta(GetCarreraAlumno(session("Codcli")))%>','<%=TieneNotas(session("Codcli"),Session("Ano_Ed"),Session("Periodo_Ed"))%>')" onMouseOver="MM_swapImage('Image521','','imag/f-der/encuesta-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/encuesta-of.gif" name="Image521" width="192" height="20" border="0" id="Image52"></a></td>
            <%else%>
            <td width="93"><a href="MensajesBloqueos.asp" onMouseOver="MM_swapImage('Image522','','imag/f-der/encuesta-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/encuesta-of.gif" name="Image522" width="192" height="20" border="0" id="Image52"></a></td>
            <%end if%>
            <%else%>
            <td width="93"><a href="MensajeBloqueoHabilitaED.asp" onMouseOver="MM_swapImage('Image52','','imag/f-der/encuesta-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="imag/f-der/encuesta-of.gif" name="Image52" width="192" height="20" border="0" id="Image52"></a></td>
            <%end if%>
            <td rowspan="2" height="2" width="5">&nbsp;</td>
            <td rowspan="2" width="451" valign="top" class="tex-normales"> Te 
              permite evaluar a tus docentes.</td>
            <td rowspan="2" width="451" valign="top" class="linkrojo">
<p><font size="2" face="Arial, Helvetica, sans-serif"><a href="Pauta%20de%20Evaluacion%20Docente.doc" target="_blank" class="linkrojo">Baja 
                este instructivo de Evaluaci&oacute;n Docente</a> </font></p></td>
          </tr>
          <tr> 
            <td width="93">&nbsp;</td>
          </tr>
        </table></td>
    </tr>
    <tr valign="top"> 
      <td height="21"> <table width="584" border="0" cellspacing="0" cellpadding="0" height="108">
          <tr> 
   		 <%if habilitada8="S" then%>
            <%if PermisoI=1 or PermisoI <> 0 then %>
            <td width="178"><a href="Actualizador.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image21','','ImagenBoton/actualiza-on.gif',1)"><img src="ImagenBoton/actualiza-of.gif" name="Image21" width="131" height="20" border="0"></a></td>
            <%else%>
            <td width="133"><a href="MensajesBloqueos.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image21','','ImagenBoton/actualiza-on.gif',1)"><img src="ImagenBoton/actualiza-of.gif" name="Image21" width="131" height="20" border="0"></a></td>
            <%end if%>
			<%else%>
	            <td width="273"><a href="MensajeBloqueoHabilita.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image21','','ImagenBoton/actualiza-on.gif',1)"><img src="ImagenBoton/actualiza-of.gif" name="Image21" width="131" height="20" border="0"></a></td>
			<%end if%>
            <td height="0" colspan="2" class="tex-normales">Esta opci&oacute;n 
              te permite actualizar tus datos personales</td>
          </tr>
		  
		  
		  
		  
          <tr>
            <td>&nbsp;</td>
            <td height="1">&nbsp;</td>
            <td valign="top" class="tex-normales">&nbsp;</td>
          </tr>
          <tr> 
   		 <%if habilitada_Opcion11="S" then%>
            <%if PermisoJ=1 or PermisoJ <> 0 then %>
            <td width="178"><p><a href="encuesta_ei.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image21','','ImagenBoton/actualiza-on.gif',1)"><img src="imag/botones/encuesta_de_sat-of.gif" name="Image21" width="169" height="20" border="0"></a></p>            </td>
            <%else%>
            <td width="133"><a href="MensajesBloqueos.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image21','','ImagenBoton/actualiza-on.gif',1)"><img src="imag/botones/encuesta_de_sat-of.gif" name="Image21" width="169" height="20" border="0"></a></td>
            <%end if%>
			<%else%>
	            <td width="273"><a href="MensajeBloqueoHabilita.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image21','','ImagenBoton/actualiza-on.gif',1)"><img src="imag/botones/encuesta_de_sat-of.gif" name="Image21" width="169" height="20" border="0"></a></td>
			<%end if%>
            <td height="0" colspan="2" class="tex-normales"><p>Encuesta de satisfacci&oacute;n</p>
              </td>
          </tr>
          <tr>
            <td ><a href="vClave.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image31','','ImagenBoton/descarga-on.gif',1)"> <img src="ImagenBoton/descarga-of.gif" name="Image31"  height="20" border="0"  onClick=""> </a></td>
            <td  class="tex-normales">Esta opci&oacute;n te permite descargar <br>
                los apuntos por Asignatura.</td>
            <td valign="top" class="tex-normales">&nbsp;</td>
          </tr>
          <tr>
            <td><a href="vClave1.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image33','','ImagenBoton/avisos-on.gif',1)"><img src="ImagenBoton/avisos-of.gif" name="Image33"  height="20" border="0"  onClick=""> </a></td>
            <td height="0" colspan="2" class="tex-normales">Esta opci&oacute;n te permite ver los avisos ingresados.</td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td height="0">&nbsp;</td>
            <td valign="top" class="tex-normales">&nbsp;</td>
          </tr>
          <tr> 
            <td><a href="javascript:Terminar()"><img src="imag/botones/salir-of.gif" width="114" height="20" name="Image6" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image6','','imag/botones/salir-on.gif',1)" border="0"></a></td>
            <td height="0">&nbsp;</td>
            <td width="273" valign="top" class="tex-normales">&nbsp;</td>
          </tr>
          <tr> 
            <td width="178" height="25">&nbsp;</td>
            <td width="133" height="1">&nbsp;</td>
            <td width="273" valign="top" class="tex-normales">&nbsp;</td>
          </tr>
          <tr> 
            <td height="25">&nbsp;</td>
            <td width="133" height="0">&nbsp;</td>
            <td width="273" valign="top" class="tex-normales">&nbsp;</td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td width="133" height="0">&nbsp;</td>
            <td width="273" valign="top" class="tex-normales">&nbsp;</td>
          </tr>
        </table>
        <p><font size="2" face="Arial, Helvetica, sans-serif"> </font></p></td>
    </tr>
  </table>
  <p class="tex-normales">&nbsp;</p>
</div>
</body>

<script languaje="javascript">
function Terminar()
{
<% if session("logueado")="1" then %>
   if (confirm("Desea abandonar?") )
   { 
   	  <% Session("MiValor")=0%>
      window.location = "salir.asp";
        //window.top.location = "alumn-udd.asp";
   }
<% end if%>   
}
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
</html>
