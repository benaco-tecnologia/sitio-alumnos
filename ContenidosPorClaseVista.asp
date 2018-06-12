<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
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
<link rel="stylesheet" href="css/tex-normales.css" type="text/css">
<body  leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#FFFFFF">

<OBJECT RUNAT=server PROGID=ADODB.Recordset id=RamosSolici name=RamosSolici>
</OBJECT> 
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=RamosPuede name=RamosPuede>
</OBJECT> 

<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if
%>
<%
Dim strContenidos, strRamosPuede, Rut, CodCli
CodCli = Session("CodCli")
CodSede = Session("CodSede")
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if

ano =AnoAcad()
periodo=PeriodoAcad()
codramo = request.Form("CodigoRamo")
codsecc = request.Form("CodigoSeccion")
codcarr = request.Form("CodigoCarrera")



STRnom_ramo = "select CODRAMO + ' - ' + NOMBRE as ramo_nom FROM RA_RAMO WHERE CODRAMO = '" & codramo & "'"
if bcl_ado(STRnom_ramo,rstRAMO) then
	nom_ramo =valnulo(rstRAMO("ramo_nom"),"")
end if

STRcodcarr = "select codcarpr  FROM mt_alumno WHERE codcli = '" & CodCli & "'"
if bcl_ado(STRcodcarr,rstcarrera) then
	codcarr =valnulo(rstcarrera("codcarpr"),"")
end if

strANOP="select ANO,PERIODO from MT_CARRERAPERIODO WHERE CODCARR='"& session("codcarr") &"' AND codpestud='"& session("codpestud") &"' "

if bcl_ado(strANOP,rstANOP) then
	ano =valnulo(rstANOP("ANO"),NUM_)
	periodo=valnulo(rstANOP("PERIODO"),NUM_)
end if

if EstaHabilitadaNW (470)="S" then 
	if GetPermisoNW(470) ="N" then
		response.Redirect("MensajesBloqueos.asp")		
	end if
else
	response.Redirect("MensajeBloqueoHabilita.asp")
end if

seccio = "select distinct codsecc  FROM ra_nota WHERE codcli = '" & CodCli & "' and ramoequiv = '" &  codramo & "'  and ano = " & ano & " and periodo = " & periodo

codsecc = 0

if bcl_ado(seccio,rstseccio) then
	codsecc = rstseccio("codsecc")
end if

STRprofesor = ""
STRprofesor = "select distinct codprof  FROM ra_seccio WHERE  codramo = '" &  codramo & "'  and ano = " & ano & " and periodo = " & periodo & " and codsecc = " & codsecc & " "


if bcl_ado(STRprofesor,rstprof) then
	profesor =valnulo(rstprof("codprof"),"")
end if

STRprofesor = ""
STRprofesor = "select nombres + ' ' + ap_pater + ' ' + ap_mater as nom_profe FROM ra_profes WHERE  codprof = '" &  profesor & "'" 

if bcl_ado(STRprofesor,rstprof_nom) then
	profesor =valnulo(rstprof_nom("nom_profe"),"")
end if



strContenidos = "SP_GRILLA_CONTCLASE_HISTORICA_ALUMNOS " & ano & "," & periodo & ",'" & codramo & "','" & codcarr & "'," & codsecc 

		     
set Contenidos = Session("Conn").execute (strContenidos)


strRamosSolici = "SELECT b.codramo, b.nombre, a.estado, a.codsecc, a.Ramoequiv, c.nombre as NombreReal, a.CodSeccOri " & _
                " FROM ra_solici a, ra_ramo b, ra_ramo c " & _
                " WHERE a.ramoequiv = b.codramo and " & _
                " a.ramoequiv *= c.codramo and " & _
                " a.codcli ='" & CodCli & "' and " & _
                " a.ano = '" & ano & "' and " & _
                " a.PerInscrip = " & SESSION("PER_ID") & " and " & _
                " a.periodo = '" & periodo & "' "

'agregar el año y periodo a esta query
'RESPONSE.WRITE(strRamosSolici)
'RESPONSE.End

RamosSolici.Open strRamosSolici, Session("Conn")
dim paso 
paso=0
do while not RamosSolici.eof and paso=0
    if RamosSolici("codsecc")<>RamosSolici("CodSeccori") then
	 paso=1	  
    end if
   RamosSolici.movenext
loop
RamosSolici.close

strRamosSolici = "SELECT b.codramo, b.nombre, a.estado, a.codsecc, a.Ramoequiv, c.nombre as NombreReal, CodSeccOri " & _
                " FROM ra_solici a, ra_ramo b, ra_ramo c " & _
                " WHERE a.ramoequiv = b.codramo and " & _
                " a.ramoequiv *= c.codramo and " & _
                " a.codcli ='" & CodCli & "' and " & _
                " a.ano = '" & ano & "' and " & _
                " a.PerInscrip = " & SESSION("PER_ID") & " and " & _
                " a.periodo = '" & periodo & "'  "
'agregar el año y periodo a esta query
'RESPONSE.WRITE(strRamosSolici)
'RESPONSE.End
RamosSolici.Open strRamosSolici, Session("Conn")
	
%>
<Script Language="JavaScript">
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
</Script><html>
<head> 
<title>Portal de Alumnos en L&iacute;nea</title>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1"> 
</head>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript"  SRC="include/tooltip.js"></SCRIPT>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
<%' if not InscritoNormal() then %>
  <script>
  //   alert("No ha inscrito aún vía solicitud Normal ...");
  </script>   
<%'  end if%>
<table border="0" cellpadding="0" cellspacing="0"  align="left">
  <tr>
	<td>
		<table width="862" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right" ><% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td width="795" valign="top">
			<% CargarTop1()%><% SubMenu()%>	
				<!--<div style="width:825PX; height:480px; overflow:scroll;" align="center">-->
                
				 
					<table border="0" cellpadding="0" cellspacing="15" width="757">
					  <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
					  <tr valign="top" > 
						<td ><img name="fder_r2_c1" src="imagenes/titulos/T-contenidos_clases.gif" width="300" height="38" border="0"></tr>
					  <tr valign="top"> 
						<td colspan="3" height="80"> 
						  <table width="801" cellspacing="1" cellpadding="0" height="59" border="0" bordercolor="#FFFFFF">
							  <td height="20" colspan="11"> 
                              
                              
                              
                              <table width="810" border="0" cellpadding="2" cellspacing="2">
                    <tr>
                      <td width="133" height="25" background="Imagenes/fdo-cabecera-cel.jpg" class="text-cabecera-celda" ><strong class="text-cabecera-celda">ASIGNATURA : </strong></td>
                      <td width="435" height="25" bgcolor="#9FDBE8" class="tex-totales-celda"><font face="Arial, Helvetica, sans-serif" class="tex-totales-celda"><%=nom_ramo%></font></td>
                    </tr>
                    <tr>
                      <td height="25" background="Imagenes/fdo-cabecera-cel.jpg" bgcolor="#D7EAFD" class="text-cabecera-celda" ><strong class="text-cabecera-celda">SECCI&Oacute;N : </strong></td>
                      <td height="25" bgcolor="#9FDBE8" class="tex-totales-celda"><%=codsecc%></td>
                    </tr>
                    <tr>
                      <td height="25" background="Imagenes/fdo-cabecera-cel.jpg" bgcolor="#D7EAFD" class="text-cabecera-celda" ><strong class="text-cabecera-celda">PROFESOR : </strong></td>
                      <td height="25" bgcolor="#9FDBE8" class="tex-totales-celda"><%=profesor%></td>
                    </tr>                    
                 </table>
                              
                             <br> 
                              
                              
                              </td>
							</tr>
							<tr background="imagenes/fdo-cabecera-cel22.jpg"> 
							  <td width="200" height="30"  background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center" ><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Contenido</font></b></font></div></td>
							  <td width="200" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Descripci&oacute;n</font></b></font></div></td>
							  <td width="150" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Avance real</font></b></font></div></td>
                              <td width="150" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Avance esperado</font></b></font></div></td>
                              <td width="100" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Hora real</font></b></font></div></td>
                              <td width="150" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Hora Planificada</font></b></font></div></td>

							 
							</tr>
							
							<% 
						   While Not Contenidos.Eof
						  %>
							
						   
							<tr bgcolor="#DBECF2" height="25"> 
							 <td width="200" height="30" align="left" ><font face="Arial, Helvetica, sans-serif" size="1"><%=Contenidos("Contenido")%></font></td>
							 <td width="200" height="30" align="left"><font face="Arial, Helvetica, sans-serif" size="1"><%=Contenidos("Descripcion")%></font></td>
                             <td width="150" height="30" align="center" ><font face="Arial, Helvetica, sans-serif" size="1"><%=Contenidos("AVANCE_REAL")%></font></td>
							 <td width="150" height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=Contenidos("AVANCE_ESPERADO")%></font></td>
                              <td width="100" height="30" align="center" ><font face="Arial, Helvetica, sans-serif" size="1"><%=Contenidos("HORA_REAL")%></font></td>
							  <td width="150" height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=Contenidos("HORA_PLANIFICADA")%></font></td>
                
							 
							 
							</tr>
							<%
							 Contenidos.MoveNext
						   Wend %>
				
						  </table>
						  <br>
						 
					    <p></p></td>
					  </tr>
	</table>
    <table>
           <tr>
    <td width="83" align="left"><A onMouseOver="MM_swapImage('Image7','','Imagenes/botones/volver-on.gif',1)" onmouseout=MM_swapImgRestore() href="javascript: window.top.location.href ='ContenidosSelClase.asp'"><IMG src="Imagenes/botones/volver-of.gif" border="0"  name="Image7"></A></td>
  </tr>
  </table>
	</table>
	</td>
  </tr>
</table>


</body>
<script languaje = "javascript">
function ConfirmaInscrip()
{
  document.form1.action = "confirmacion.asp"
 // alert("Hola");
  document.form1.submit();
} 
function visualiza() {
  var x=window.open ("tex-resumenInscrip.htm","ResumenInscripción","width=500,height=400,resizable=yes,toolbar=yes");
 }

function imprime()
{
  parent.focus();
  parent.print(); 
  //window.print();
  //parent.Horario.focus();
  //parent.Horario.print();
}

</script>
</html>
<%
Contenidos.Close()
RamosSolici.Close()
%>
<!--#INCLUDE file="include/desconexion.inc" -->
