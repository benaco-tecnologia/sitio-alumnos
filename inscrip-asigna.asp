<%
  Response.Expires = -1
%>
<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%

Dim strRamosDebe, strRamosPuede, Rut, CodCli, strPaso, Rs, SumaOtrosCreditos, RstCred, RstCombo, Strcombo
Dim CredMaxCombo

Audita 466,"Ingresa a Inscripcion de Asignaturas"

CodCli = Session("CodCli")
CodSede = Session("CodSede")
CodPestud = Session("CodPestud")
NivelALumno = session("Nivel")
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if

CredMaxCombo = 100

OCULTABOTONOTRASASIGNATURAS_PA ="NO"
strParame="SELECT coalesce(dbo.Fn_ValorParame('OCULTABOTONOTRASASIGNATURAS_PA'),'NO')Parame"
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
	OCULTABOTONOTRASASIGNATURAS_PA=rstParame("Parame")
end if


Str = "Select dbo.Fn_ValorParame('inscrip_especial_debe') inscrip_especial_debe"' as from mt_parame"
'response.Write(Str)
'response.End()

Set Rst = Session("Conn").Execute(Str)

If not Rst.eof then 
	Inscrip_especial_debe = Rst("inscrip_especial_debe")
	Session("Inscrip_especial_debe")=  Rst("inscrip_especial_debe")
else
	Inscrip_especial_debe = "N"
	Session("Inscrip_especial_debe")="N"
End if 
'Rst.Close()


Strcombo = "select coalesce(a.arancelxasignatura,'N') as arancelxasignatura ,b.combo "
strcombo = strcombo & " from mt_carrer a, mt_alumno b where "
strcombo = strcombo & " a.codcarr = b.codcarpr"
strcombo = strcombo & " and b.codcli = '" & trim(CodCli) & "'"
if bcl_ado(strcombo, rstcombo) then
	if trim(valnulo(rstcombo("arancelxasignatura"),str_))="S" then
		CredMaxCombo = valnulo(rstcombo("combo"),num_)
		if CredMaxCombo=0 then
			CredMaxCombo = 5
		end if
	end if
end if

GetCreditos CodPestud, CredMin, CredMax, CredMinOtros, CredMaxOtros

if CodCli="20071ADER0224" then
'	response.Write(CredMaxCombo)
'	response.End()
end if

Ano = AnoAcad()
Periodo = PeriodoAcad()
   			   
strRamosDebe = "SELECT b.codramo, b.nombre, a.inscrito, a.codsecc, a.ramoequiv, b.credito, "
strRamosDebe = strRamosDebe & " a.ramoequiv_i, a.codsecc_i,  coalesce (a.preinscrito,'N')as preinscrito, d.nivel,a.codcli,a.ano,a.periodo,"
'strRamosDebe = strRamosDebe & " a.prioridad,coalesce(d.tipo,'Z') , coalesce(a.SOLICITUDESPECIAL,'N') AS SOLICITUDESPECIAL, coalesce(b.hor_teo,'0') AS hor_teo "
strRamosDebe = strRamosDebe & " a.prioridad,coalesce(d.tipo,'Z') ,  coalesce(b.hor_teo,'0') AS hor_teo "
strRamosDebe = strRamosDebe & " FROM ra_carga a with (nolock), "
strRamosDebe = strRamosDebe & " ra_ramo b with (nolock), mt_alumno c with (nolock), ra_curric d with (nolock)"
strRamosDebe = strRamosDebe & " WHERE a.codramo = b.codramo "
strRamosDebe = strRamosDebe & " and a.codcli ='" & CodCli & "' "
strRamosDebe = strRamosDebe & " and a.ano = '" & ano & "' "
strRamosDebe = strRamosDebe & " and a.periodo ='" & periodo & "' "
strRamosDebe = strRamosDebe & " and a.codcli=c.codcli"
strRamosDebe = strRamosDebe & " and a.prioridad = 0"
strRamosDebe = strRamosDebe & " and c.codpestud*=d.codpestud"
strRamosDebe = strRamosDebe & " and a.codramo*=d.codramo"
strRamosDebe = strRamosDebe & " And  a.INSCRITO <> 'S' "
strRamosDebe = strRamosDebe & " order by a.prioridad,coalesce(d.tipo,'Z')"
'response.Write(strRamosDebe)
'strRamosDebe = strRamosDebe & " UNION SELECT b.codramo, b.nombre, a.inscrito, a.codsecc, a.ramoequiv, b.credito, "
'strRamosDebe = strRamosDebe & " a.ramoequiv_i, a.codsecc_i, coalesce (a.preinscrito,'N')as preinscrito, d.nivel,a.codcli,a.ano,a.periodo,"

'strRamosDebe = strRamosDebe & " a.prioridad,coalesce(d.tipo,'Z'), coalesce(a.SOLICITUDESPECIAL,'N') AS SOLICITUDESPECIAL, coalesce(b.hor_teo,'0') AS hor_teo "

'strRamosDebe = strRamosDebe & " a.prioridad,coalesce(d.tipo,'Z'), coalesce(b.hor_teo,'0') AS hor_teo "
'strRamosDebe = strRamosDebe & " FROM ra_cargaactividad a with (nolock), "
'strRamosDebe = strRamosDebe & " ra_ramo b with (nolock), mt_alumno c with (nolock), ra_curric d with (nolock)"
'strRamosDebe = strRamosDebe & " WHERE a.codramo = b.codramo "
'strRamosDebe = strRamosDebe & " and a.codcli ='" & CodCli & "' "
'strRamosDebe = strRamosDebe & " and a.ano = '" & ano & "' "
'strRamosDebe = strRamosDebe & " and a.periodo ='" & periodo & "' "
'strRamosDebe = strRamosDebe & " and a.codcli=c.codcli"
'strRamosDebe = strRamosDebe & " and a.prioridad = 0"
'strRamosDebe = strRamosDebe & " and c.codpestud*=d.codpestud"
'strRamosDebe = strRamosDebe & " and a.codramo*=d.codramo"
'strRamosDebe = strRamosDebe & " order by a.prioridad,coalesce(d.tipo,'Z'),d.nivel desc"
'RESPONSE.WRITE(strRamosDebe)
'RESPONSE.End()
strRamosPuede  = "SELECT b.codramo, b.nombre, a.inscrito, a.codsecc, a.ramoequiv, b.credito, "
strRamosPuede  = strRamosPuede & " a.ramoequiv_i, a.codsecc_i, a.preinscrito, d.nivel,a.codcli,a.ano,a.periodo,"
strRamosPuede  = strRamosPuede & " a.prioridad,coalesce(d.tipo,'Z') as tipo, coalesce(b.hor_teo,'0') AS hor_teo "
strRamosPuede  = strRamosPuede & " FROM ra_carga a with (nolock), "
strRamosPuede  = strRamosPuede & " ra_ramo b with (nolock), mt_alumno c with (nolock), ra_curric d with (nolock)"
strRamosPuede  = strRamosPuede & " WHERE a.codramo = b.codramo "
strRamosPuede  = strRamosPuede & " and a.codcli ='" & CodCli & "' "
strRamosPuede  = strRamosPuede & " and a.ano = '" & ano & "' "
strRamosPuede  = strRamosPuede & " and a.periodo ='" & periodo & "' "
strRamosPuede  = strRamosPuede & " and a.codcli=c.codcli"
strRamosPuede  = strRamosPuede & " and a.prioridad = 1"
strRamosPuede  = strRamosPuede & " and c.codpestud*=d.codpestud"
strRamosPuede  = strRamosPuede & " and a.codramo*=d.codramo"
strRamosPuede  = strRamosPuede & " And  a.INSCRITO <> 'S' "
strRamosPuede  = strRamosPuede & " order by a.prioridad,coalesce(d.tipo,'Z')"

'strRamosPuede  = strRamosPuede & " UNION SELECT b.codramo, b.nombre, a.inscrito, a.codsecc, a.ramoequiv, b.credito, "
'strRamosPuede  = strRamosPuede & " a.ramoequiv_i, a.codsecc_i, a.preinscrito, d.nivel,a.codcli,a.ano,a.periodo,"
'strRamosPuede  = strRamosPuede & " a.prioridad,coalesce(d.tipo,'Z') as tipo, coalesce(b.hor_teo,'0') AS hor_teo"
'strRamosPuede  = strRamosPuede & " FROM ra_carga a with (nolock), "
'strRamosPuede  = strRamosPuede & " ra_ramo b with (nolock), mt_alumno c with (nolock), ra_curric d with (nolock)"
'strRamosPuede  = strRamosPuede & " WHERE a.codramo = b.codramo "
'strRamosPuede  = strRamosPuede & " and a.codcli ='" & CodCli & "' "
'strRamosPuede  = strRamosPuede & " and a.ano = '" & ano & "' "
'strRamosPuede  = strRamosPuede & " and a.periodo ='" & periodo & "' "
'strRamosPuede  = strRamosPuede & " and a.codcli=c.codcli"
'strRamosPuede  = strRamosPuede & " and a.prioridad = 1"
'strRamosPuede  = strRamosPuede & " and c.codpestud*=d.codpestud"
'strRamosPuede  = strRamosPuede & " and a.codramo*=d.codramo"
'strRamosPuede  = strRamosPuede & " order by a.prioridad,coalesce(d.tipo,'Z')"

'response.Write strRamosDebe
'response.Write("+++++++++++++++++++++++++++++++++++++++++++")
'response.Write strRamosPuede
'response.End()
strCred = "SP_TRAE_DATOS_ALUMNOS_OBSERVADOS '"& session("codcli") & "'"
if bcl_ado(strCred, rstCred) then
	CRMAXALUMNOBS = rstCred("CRMAXALUMNOBS")
	PROMEDIOALUMNO = rstCred("PROMEDIO")
	NOTAALUMNOBS = rstCred("NOTAALUMNOBS")
end if 

RamosDebe.Open strRamosDebe, Session("Conn")
RamosPuede.Open strRamosPuede, Session("Conn")

%>

<%
'response.Write("NOSE PUEDE CONTINUAR HASTA AGREGAR CAMPO SOLICITUDESPECIAL A RA_CARGA")
'response.End()
if RamosDebe.EOF and RamosPuede.EOF then
     'response.Redirect("sinramos.htm")
	 response.Redirect("sinramos.asp")
end if
%>

<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=iso-8859-1">
<link rel="stylesheet" href="css/tablas.css" type="text/css">
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

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<!--
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#FFFFFF" onLoad="MM_preloadImages('../../../../Desktop%20Folder/admalumnosx/imag/botones/todas-las-asig-on.gif','../../../../Desktop%20Folder/admalumnosx/imag/botones/seleccionar-on.gif','../../../../Desktop%20Folder/admalumnosx/imag/botones/terminar-inscrip-on.gif','imag/botones/todas-las-asig-on.gif','imag/botones/seleccionar-on.gif','imag/botones/seccione-horarios-on.gif',
'imag/botones/otras-asig-on.gif','imag/botones/terminar-inscrip-on.gif');levanta1();">
ESTO ME DA ERROR--!>

<body onLoad="MM_preloadImages('imag/botones/terminar%20inscripcion-on.gif','imag/botones/otras-asig-on.gif')">
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=RamosDebe name=RamosDebe>
</OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=RamosPuede name=RamosPuede></OBJECT>

<Script Language="JavaScript">
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
</Script>
	<script>
		//function levanta(){
		//pop=window.open('informateaqui_login.htm','popup','toolbar=no,menubar=no,scrollbars=yes,resizable=no,width=550,height=450');
		
		//}
	</script>
<meta http-equiv="Content-Type" content="text/html;">
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
<body bgcolor="#ffffff" leftmargin="0" rightmargin="0" topmargin="0" marginwidth="0" marginheight="0"  > 
<div align="left">
<table border="0" cellpadding="0" cellspacing="0"  align="left" >
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td valign="top"  >
			<% CargarTop1()%><% SubMenu()%>
			<form name="form1" method="post" action="asignatura-seccion.asp"><input type = "hidden" name = "Ramos" value = "">
				<table border="0" cellpadding="0" cellspacing="15" align="left" width="584">
				  <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
				  <tr> 
					<td colspan="3"><p><span style="font-size: 14px"><p id ="lblTituloInscripcion" style="font-size: 25px" class="text-menu">Inscripci&oacute;n Normal de Asignaturas</p></span></p></td>
					<td width="248">&nbsp;</td>
				  </tr>
				  <tr valign="bottom"> 
					<td colspan="3" height="19" id="tdtrbtns"><p>
                    <font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font color="#000000" size="-7"><a href="javascript:EnviaRamos()" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image4','','Imagenes/botones/rev-secc-horarios-on.gif',1)"><img src="Imagenes/botones/rev-secc-horarios-of.gif" width="246" height="45" border="0" name="Image4"></a></font></font></font>
                    <font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>
                    <font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font color="#000000" size="-7"><a href="javascript:confirma()" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image2','','Imagenes/botones/terminar-inscrip-on.gif',1)"><img src="Imagenes/botones/terminar-inscrip-of.gif" width="178" height="45" border="0" name="Image2"></a></font></font></font>
                    <font face="Arial, Helvetica, sans-serif" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font>
                    <%if OCULTABOTONOTRASASIGNATURAS_PA = "SI" then%>
                    <font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Arial, Helvetica, sans-serif" size="2"></font></font></font></font></font></font></font>
                    <%else%>
                    <font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Arial, Helvetica, sans-serif" size="2"><a href="javascript:EnviaTodosRamos()" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','Imagenes/botones/otras-asignaturas-on.gif',1)"><img src="Imagenes/botones/otras-asignaturas-of.gif" width="178" height="45" name="Image1" border="0"></a></font></font></font></font></font></font></font>
                    <%end if%>
                    </p>
				    </td>
				  </tr>
				  <tr valign="top"> 
					<td colspan="3"> 
                    <table width="750" cellspacing="1" cellpadding="0" height="59" border="0" bordercolor="#FFFFFF">
						<td height="16" colspan="5"> 
						  <div align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="2" color="#000000"><strong id="lblDebeAsig">Debe 
							  inscribir las siguientes asignaturas:</strong></font></font></div>						</td>
						</tr>
						<tr background="Imagenes/fdo-cabecera-cel22.jpg"> 
						  <td width="48" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
							<div align="center" background="Imagenes/fdo-cabecera-cel22.jpg"><b><font color="#FFFFFF"><font size="1" face="Geneva, Arial, Helvetica, san-serif">C&oacute;digo</font></font></b></div>						  </td>
						  <td width="314" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
							<div align="center"><b><font color="#FFFFFF"><font id="lblCabeceraAsignaturaDebe"size="1" face="Geneva, Arial, Helvetica, san-serif">Asignatura</font></font></b></div>						  </td>
						  <td width="83" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
							<div align="center"><b><font color="#FFFFFF"><font size="1" face="Geneva, Arial, Helvetica, san-serif">secci&oacute;n</font></font></b></div>						  </td>
						  <td width="197" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
							<div align="center"><b><font color="#FFFFFF"><font size="1" face="Geneva, Arial, Helvetica, san-serif">Horario</font></font></b></div>						  </td>
						  <td width="102" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
							<div align="center"><b><font color="#FFFFFF"><font size="1" face="Geneva, Arial, Helvetica, san-serif">Solicitado</font></font></b></div>						  </td>
  						  <td width="40" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
							<div align="center"><b><font color="#FFFFFF"><font size="1" face="Geneva, Arial, Helvetica, san-serif">Hrs</font></font></b></div>						  </td>
						</tr>
						<% 
						
					  If RamosDebe.Eof Then
						   TodoslosDebes = true
						   TodoslosDebesNivel = true
						   sumCreditos = 0
					  %>
						<%
					  Else
					   
					   TodoslosDebes = true
					   TodoslosDebesNivel = true
					   
					   While Not RamosDebe.Eof
						 
						 if Ucase(RamosDebe("preinscrito")) = "S"  then
						   
						   if (Ucase(RamosDebe("preinscrito")) = "S") then
							  CodSecc = RamosDebe("codsecc_i")
							  Horario = GetHor2(RamosDebe("ramoequiv_i"), CodSede, RamosDebe("codsecc_i"), Ano, Periodo)							  
						   else
							  CodSecc = RamosDebe("codsecc")
							  Horario = GetHor2(RamosDebe("ramoequiv"), CodSede, RamosDebe("codsecc"), Ano, Periodo)
						   end if   
						   strCred = "select codramo from ra_curric "
						   strCred = strCred & " s where codpestud = '" & trim(CodPestud) & "' "
						   strCred = strCred & " AND  codramo = '" & trim(RamosDebe("codramo")) & "' "
						   
						   if bcl_ado(strCred, rstCred) then
							   SumCreditos = SumCreditos + ValNulo(RamosDebe("credito"), NUM_)
						   else
							   SumOtrosCreditos = SumOtrosCreditos + ValNulo(RamosDebe("credito"), NUM_)		
						   end if
						   
						   
						 else
						   	CodSecc = ""
						   	Horario = ""
						   '  ha inscrito todos los debes
						   	'TodoslosDebesNivel = false
							
							'response.Write("weno hasta aqui todo va bien")
							'response.End()
						
							'COMENTADO PORQUE NO EXISTE CAMPO EN RA_CARGAACTIVIDADE
							'RamosDebe("SOLICITUDESPECIAL") = "N"
						   	'if Session("Inscrip_especial_debe") = "S" then 
							'  	if Ucase(RamosDebe("SOLICITUDESPECIAL")) ="N" and RamosDebe("nivel") = 1  then
							TodoslosDebes = false
							'	   TodoslosDebesNivel = false
							'	end if
							'else
							'	TodoslosDebes = false
							'end if 
							
							
							
							
							
							
						 end if
					   %>
						<tr bgcolor="#DBECF2"> 
						  <td width="48" height="8" align="center"><font size="1" face="Geneva, Arial, Helvetica, san-serif"><%=trim(RamosDebe("CodRamo"))%></font></td>
						  <td width="314" height="0" align="center"><font size="1" face="Geneva, Arial, Helvetica, san-serif"><%=RamosDebe("Nombre") & "(" & RamosDebe("Credito") & ")" %></font></td>
						  <td width="83" height="0" align="center"><font size="1" face="Geneva, Arial, Helvetica, san-serif"><%=CodSecc%></font></td>
						  <td height="0" align="center" bgcolor="#DBECF2"><font size="1" face="Geneva, Arial, Helvetica, san-serif"><%=Horario%></font></td>
						  <td width="102" height="0" valign="top"> 
							<div align="center"> <font size="1" face="Geneva, Arial, Helvetica, san-serif"> 
							  	 <% if RamosDebe("nivel") = 1 then 
								 		if Ucase(RamosDebe("preinscrito")) = "S" then ' OR Ucase(RamosDebe("SOLICITUDESPECIAL")) = "S"  then
								 %>
							  				<input type="checkbox" name="<%=RamosDebe(0)%>" value="checkbox" checked disabled>
                              				<input type = "hidden" name = "<%=RamosDebe(0)%>" value = "checkbox">
                               			<%else %>
							   				<input type="checkbox" name="<%=RamosDebe(0)%>" value="checkbox" disabled>
                              				<input type = "hidden" name = "<%=RamosDebe(0)%>" value = "checkbox">
                                 <%		
								 		end if 	
								  else
								 		if Ucase(RamosDebe("preinscrito")) = "S" then ' OR Ucase(RamosDebe("SOLICITUDESPECIAL")) = "S" then
								        %>
							  				<input type="checkbox" name="<%=RamosDebe(0)%>" value="checkbox" checked disabled>                              			
                               			<%else %>
							   				<input type="checkbox" name="<%=RamosDebe(0)%>" value="checkbox" disabled>                              				
                                 
							 	 <% 	end if
								 end if
								  %>
							  </font> </div>						  </td>
                              <td height="0" align="center" bgcolor="#DBECF2"><font size="1" face="Geneva, Arial, Helvetica, san-serif"><%=RamosDebe("hor_teo")%></font></td>
						</tr>
						<%
						 'RESPONSE.WRITE(RamosDebe("CODRAMO"))
						 'RESPONSE.End()
						 RamosDebe.MoveNext
					   Wend
					 End If %>
					  </table>
					  <p></p>
					  <div align="center"></div>
					  <table width="750" cellspacing="1" cellpadding="0" height="30" border="0" bordercolor="#FFFFFF">
						<td height="20" colspan="5"> 
						  <div align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="2" color="#000000"><strong id="lblPuedeAsig">Puede 
							  inscribir las siguientes asignaturas:</strong></font></font></div>						</td>
						</tr>
						<tr background="Imagenes/fdo-cabecera-cel22.jpg"> 
						  <td width="48" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
							<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">C&oacute;digo</font></b></font></div>						  </td>
						  <td width="314" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
							<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1" id="lblCabeceraAsignaturaPuede" >Asignatura</font></b></font></div>						  </td>
						  <td width="83" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
							<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">secci&oacute;n</font></b></font></div>						  </td>
						  <td width="197" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
							<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Horario</font></b></font></div>						  </td>
						  <td width="102" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
							<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Solicitado</font></b></font></div>						  </td>
 							<td width="40" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
							<div align="center"><b><font color="#FFFFFF"><font size="1" face="Geneva, Arial, Helvetica, san-serif">Hrs</font></font></b></div>	                            
						</tr>
						<% 
					  If RamosPuede.Eof Then
					  %>
						<%
					  Else
					   While Not RamosPuede.Eof
						 if Ucase(RamosPuede("preinscrito")) = "S" then
						   if Ucase(RamosPuede("preinscrito")) = "S" then
								 CodSecc = RamosPuede("codsecc_i")
								 RamoEquiv = RamosPuede("ramoequiv_i")
								 Horario = GetHor2(RamosPuede("ramoequiv_i"), CodSede, RamosPuede("codsecc_i"), Ano, Periodo)
						   else
							 CodSecc = RamosPuede("codsecc")
							 RamoEquiv = RamosPuede("ramoequiv")
							 Horario = GetHor2(RamosPuede("ramoequiv"), CodSede, RamosPuede("codsecc"), Ano, Periodo)
						   end if
							   strCred = "select codramo from ra_curric "
							   strCred = strCred & " s where codpestud = '" & trim(CodPestud) & "' "
							   strCred = strCred & " AND codramo = '" & trim(RamosPuede("codramo")) & "' "
							  ' RESPONSE.WRITE(strCred)
							   'RESPONSE.End()
							   if bcl_ado(strCred, rstCred) then
								   SumCreditos = SumCreditos + ValNulo(RamosPuede("credito"), NUM_)
							   else
								   SumOtrosCreditos = SumOtrosCreditos + ValNulo(RamosPuede("credito"), NUM_)		
							   end if
						 else
						   CodSecc = ""
						   Horario = ""
						 end if
					   %>
						<tr bgcolor="#DBECF2"> 
						  <td width="48" height="8" align="center"><font face="Verdana" size="1"><%=trim(RamosPuede(0))%></font></td>
						  <td width="314" height="0" align="center"><font face="Verdana" size="1"><%=RamosPuede(1) & "(" & RamosPuede("Credito") & ")"%></font></td>
						  <td width="83" height="0" align="center"><font face="Verdana" size="1"><%=CodSecc%></font></td>
						  <td height="0" align="center"><font face="Verdana" size="1"><%=Horario%></font></td>
						  <td width="102" height="0" valign="top"> 
							<div align="center"> <font face="Verdana" size="1"> 
							  <%
							   if Ucase(RamosPuede("preinscrito")) = "S" then
							%>
							  <input type="checkbox" name="<%=RamosPuede(0)%>" value="checkbox" checked disabled>
							  <% else %>
							  <input type="checkbox" name="<%=RamosPuede(0)%>" value="checkbox" disabled>
							  <% end if %>
							  </font> </div>						  </td>
                              <td height="0" align="center" bgcolor="#DBECF2"><font size="1" face="Geneva, Arial, Helvetica, san-serif"><%=RamosPuede("hor_teo")%></font></td>
						</tr>
						<%
						 RamosPuede.MoveNext
					   Wend
					 End If %>
					  </table>					</td>
                      
				  </tr>
				  <tr valign="top"> 
                  <td></td>
			 	 </tr>
				  <tr valign="top"> 
					<td> 
					  </td>
				  </tr>
				  <tr valign="top"> 
					<td colspan="3">&nbsp;</td>
				  </tr>
				</table>
			 </form>
	      </tr>
		</table>
	  </td>
    </tr>
  </table>
	</td>
  </tr>
</table>
<% 
if CodCli="20071ADER0224" then
	'response.Write(CredMaxCombo)
	'response.Write(SumCreditos)
	'response.End()
end if

%>
  <font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Arial, Helvetica, sans-serif" size="2"></font><font color="#000000" size="-7"></font><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font face="Verdana, Arial, Helvetica, sans-serif" color="#666666"><font color="#000000" size="-7"></font></font></font></font></font><font face="Arial, Helvetica, sans-serif" size="2"></font></font></font></font></font></font></font></div>
</body>
<script languaje = "javascript">
function visualiza() {
  var x=window.open ("Text-inscripcionasignatura.htm","ResumenInscripción","width=400,height=400,resizable=yes,toolbar=yes");
 }

function EnviaRamos()
{
var i;
  
  document.form1.Ramos.value = "";
  for (i=1; i< document.form1.elements.length; i++)
  {
     if (document.form1.elements[i].type == "checkbox")
      {
        //alert(document.form1.elements[i].name + " " );
        document.form1.Ramos.value +=  document.form1.elements[i].name + ";";
       }
  }
  //alert(document.form1.Ramos.value);
  document.form1.action = "asignatura-seccion.asp"
  document.form1.submit();
} 

function EnviaTodosRamos()
{
  document.form1.action = "solicitud-seccion.asp?P=S"
  //document.form1.action = "solicitud.htm"
 // alert("Hola");
  document.form1.submit();
} 

function confirma()
{

<%

'falta agregar validación de créditos de otros ramos, 
'separando ramos malla con ramos fuera de malla (optativos u otros ramos)

 
  EnFalta = (not TodoslosDebes) or (SumCreditos < CredMin) or (SumCreditos > CredMax) or (SumOtrosCreditos > CredMaxOtros) or (SumCreditos > CredMaxCombo)_ 
  			 or (CRMAXALUMNOBS > 0 AND SumCreditos > CRMAXALUMNOBS AND PROMEDIOALUMNO <= NOTAALUMNOBS)'mejora uarm
	
  if enFalta then 
    if (not TodoslosDebes) then %>
	//alert(<%=todoslosdebes%>);
	//alert(<%=TodoslosDebesNivel%>);
	
      //alert("Debe inscribir TODOS los Ramos DEBE");
	   //window.location = "confirmacion.asp";
	   
	   	  	 <%	if inscrip_especial_debe = "S" then %>
			  <% if (Not TodoslosDebes) and (Not TodoslosDebesNivel) then 
			 	'if  Session("InscripcionDebeRealizada") = "N" then
			   %> 
			        alert("Debe inscribir TODOS los Ramos DEBE ");
					var i;
					  
					  document.form1.Ramos.value = "";
					  for (i=1; i< document.form1.elements.length; i++)
					  {
						 if (document.form1.elements[i].type == "hidden")
						  {
							//alert(document.form1.elements[i].name + " " );
							document.form1.Ramos.value +=  document.form1.elements[i].name + ";";
						   }
					  }
						alert("- Debe Solicitar Obligatoriamente Todos los ramos DEBE y luego debe volver a terminar la inscripcion. \r Espere un momento para que ingrese a la siguiente pagina... ");
					  //alert("Debe solicitar obligatoriamente todos los ramos DEBE. \r Espere un momento para que ingrese a la siguiente pagina ");
					  //document.form1.action = "asignatura-seccion-especial-nivel.asp";
					  document.form1.action = "solicitud-seccion-especial-nivel.asp";
					  document.form1.submit();
          			  window.location = "solicitud-seccion-especial-nivel.asp";
					  return;
				<%else%>
						//alert("Debe solicitar a traves de otras asignaturas las asignaturas pendientes");
					  
					 //if (confirm(" Ud. inscribirá definitivamente las asignaturas ...\n  ¿Está seguro? "))
   					//{
       				//		window.location = "confirmacion.asp";
   					//}
			  <%end if %> 
		<%end if %>
		alert("Debe inscribir TODOS los Ramos DEBE"); 
<%  end if
    if (SumCreditos < CredMin) then%> 
	//alert("SumCreditos"+<%=SumCreditos%>+"CredMin"+<%=CredMin%>)	
      alert("Debe inscribir al menos <%=CredMin%> cr\u00e9ditos \n\n Realice una solicitud especial de Inscripci\u00f3n caso de no tener m\u00e1s cr\u00e9ditos que tomar" );  
<%  end if
    if (SumCreditos > CredMax) then %>
      alert("Debe inscribir a lo m\u00e1s <%=CredMax%> Asignaturas de Malla.\n\n Si quieres inscribir una asignatura m\u00e1s de malla, solicitala en Opci\u00f3n Solicitud Especial de Inscripci\u00f3n de Asignaturas " );  
<%  end if
    if (SumCreditos > CredMaxCombo) then %> 
      alert("Debe inscribir a lo m\u00e1s <%=CredMaxCombo%> Asignaturas." );  
<%  end if
    if (SumOtrosCreditos > CredMaxOtros) then %>
      alert("Debe inscribir a lo m\u00e1s <%=CredMaxOtros%> asignaturas optativas " );  
<%  end if

    if (CRMAXALUMNOBS > 0 AND SumCreditos > CRMAXALUMNOBS AND PROMEDIOALUMNO <= NOTAALUMNOBS) then %>
		alert("Debe inscribir a lo m\u00e1s <%=CRMAXALUMNOBS%> cr\u00e9ditos de asignaturas de Malla, ya que tiene seleccionado <%=SumCreditos%> cr\u00e9ditos." );  
<%  end if

%>    
<%
  else  
%>
  <% 'if (Not TodoslosDebes) and  (not TodoslosDebesNivel) then 
 	'if  Session("InscripcionDebeRealizada") = "N" then
   %> 
	//alert("Debe inscribir TODOS los Ramos DEBE preuba todos"); 
   if (confirm(" Ud. inscribir\u00e1 definitivamente las asignaturas ...\n  �Est\u00e1 seguro? "))
   {
	   var elem = document.getElementById("tdtrbtns");
	   elem.style.display = "none";
	   <%Audita 466,"Acepta Terminar Inscripci�n"%>
       window.location = "confirmacion.asp";
   }
   <%'else%>
   	//alert("Debe inscribir TODOS los Ramocdcdcdcdcds DEBE preuba todos");
   
<%	'end if 
  end if
%>   
}

</script>

<%ObjetosLocalizacion("inscrip-asigna.asp")%>
</html>
<%
RamosDebe.Close()
RamosPuede.Close()
%>
<!--#INCLUDE file="include/desconexion.inc" -->