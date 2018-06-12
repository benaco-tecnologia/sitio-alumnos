<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE file="analytics.asp" -->

<%
strParame="SELECT coalesce(dbo.Fn_ValorParame('USAANALYTICS'),'NO')Parame"
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
	USAANALYTICS=rstParame("Parame")
end if

if USAANALYTICS="SI" then
	analitycs() 
end if 
%>

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
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=RamosDebe name=RamosDebe>
</OBJECT> 
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
Dim strRamosDebe, strRamosPuede, Rut, CodCli
CodCli = Session("CodCli")
CodSede = Session("CodSede")
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if

ano =AnoAcad()
periodo=PeriodoAcad()
codcarr =  "select CODCARPR from mt_alumno where codcli = '" & CodCli & "'"
if bcl_ado(codcarr,rstCODCARR) then
	codcarr =valnulo(rstCODCARR("CODCARPR"),"")
end if



strANOP="select ANO,PERIODO from MT_CARRERAPERIODO WHERE CODCARR='"& session("codcarr") &"' AND codpestud='"& session("codpestud") &"' "

if bcl_ado(strANOP,rstANOP) then
	ano =valnulo(rstANOP("ANO"),NUM_)
	periodo=valnulo(rstANOP("PERIODO"),NUM_)
end if

if EstaHabilitadaNW (470)="S" then 
	if GetPermisoNW(470) ="N" then
		response.Redirect("MensajesBloqueos.asp")	
	else
		strParame="SELECT dbo.Fn_ValorParame('BLOQUEAPAENCUESTAS')Parame"
		set rstParame= Session("Conn").Execute(strParame)		
		if not rstParame.eof then
				BLOQUEAPAENCUESTAS=rstParame("Parame")
			else
				BLOQUEAPAENCUESTAS="" 
		end if
		
		if BLOQUEAPAENCUESTAS="SI" then
			'valida si contesta o no la encuesta
			if TomadeRamo(session("Codcarr"),session("Codcli"),session("ano_ed"),session("periodo_ed"))=0 then
				session("MensajeBloqueosVarios") ="Para acceder a esta opción debe responder sus Encuestas Pendientes."
				response.Redirect("MensajeBloqueo.asp")
			end if 
		end if 
		
	end if
else
	response.Redirect("MensajeBloqueoHabilita.asp")
end if

Audita 470,"Ingresa a Resumen de inscripcion"


'sql = ""
'sql = sql & " Select distinct top 1 periodo from ra_carga "
'sql = sql & " where codcli ='" & Session("CodCli") & "' "
'sql = sql & " and ano = '" & Ano & "' "
'sql = sql & " order by periodo desc "
'response.Write(sql)

'if bcl_ado(sql, rstsql2) then 
'	Periodo = rstsql2("Periodo")
'else 
'	Periodo = Session("Periodo")
'end if 

'response.Write("<br>")
'response.Write(periodo)
'response.End()
strParame="SELECT dbo.Fn_ValorParame('MUESTRA_OPORTUNIDAD')Parame"
if BCL_ADO(strParame, rstParame) then
	MUESTRA_OPORTUNIDAD=trim(valnulo(rstParame("Parame"),STR_))
end if

'Periodo = Session("Periodo")
'desde la tabla mt_parame obtener el año y periodo
strRamosDebe = "sp_listaramosalumnoportal '" & CodCli & "','" & ano & "','" & periodo & "'"

'response.write(strRamosDebe)
'response.End()			     
		  
				' Order By a.prioridad "
'agregar el año y periodo a esta query
RamosDebe.Open strRamosDebe, Session("Conn")

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

fechaCalendarioinicio=""
fechaCalendariofin=""

strCalendar="MT_CARRERAPERIODO_BUSCAPERIODO2 '"& session("codcarr") &"'"

if bcl_ado(strCalendar,RstCalendar) then
   strFechaCalendar ="SP_FechaClases "& RstCalendar("num_max_periodo") &", "& RstCalendar("ano") &", "& RstCalendar("periodo") &", '"& session("codcarr") &"'"  

   if bcl_ado(strFechaCalendar,RstFechaCalendar) then
		if isnull(RstFechaCalendar("fecha")) then
			fechaCalendarioinicio=""
		else
			fechaCalendarioinicio=RstFechaCalendar("fecha")
		end if 
		if isnull(RstFechaCalendar("fecha1")) then
			fechaCalendariofin=""
		else
			fechaCalendariofin=RstFechaCalendar("fecha1")
		end if 
   end if 
end if 

	
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
							<tr background="imagenes/fdo-cabecera-cel22.jpg"> 
							  <td width="63" height="30"  background="imagenes/fdo-cabecera-cel22.jpg"> <div align="left" ><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">C&oacute;digo</font></b></font></div></td>
							  <td width="144" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="left"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font id="lblAsignatura" size="1">Asignatura</font></b></font></div></td>
							  <td width="50" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Secci&oacute;n</font></b></font></div></td>
                           
                              <td width="50" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Ver</font></b></font></div></td>
                            
							 
							</tr>
							<% 
						  If RamosDebe.Eof Then
						  %>
							<%
						  Else
						   While Not RamosDebe.Eof

							 if Ucase(RamosDebe("inscrito")) = "S" then
							   CodSecc = RamosDebe("CodSecc")
							   'Horario = GetHor2(RamosDebe("RamoEquiv"), CodSede, RamosDebe("CodSecc"), Ano, Periodo)
							   StrSqlH = "Select dbo.HORARIO_RAMO_SECCION_SEDE('" &RamosDebe("RamoEquiv") & "','" &RamosDebe("CodSecc") & "','" &Ano & "','" &Periodo & "','"& session("codsede") &"') as Horario" 
							   	'response.Write(StrSqlH)

							   if bcl_ado(StrSqlH, SqlH) then 
							   Horario = SqlH("Horario")
							   end if
							 else
							   CodSecc = ""
							   Horario = ""
							 end if
							 
							if isnull(RamosDebe("fechainicial")) then
								fechainicio=fechaCalendarioinicio
							else
								fechainicio=RamosDebe("fechainicial")
							end if 
							if isnull(RamosDebe("fechafinal")) then
								fechafin=fechaCalendariofin
							else
								fechafin=RamosDebe("fechafinal")
							end if 
							 
						   %>
							<tr bgcolor="#DBECF2" height="25"> 
							  <td width="63" height="30" align="left" ><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("CodRamo")%></font></td>
							  <td width="400" height="30" align="left"><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("Nombre")%></font></td>
							  <td width="50" height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("codsecc")%></font></td>
                        
               
               
               <td align="center">
               <form method="POST" name="envio" action = "ContenidosPorClaseVista.asp">
               <input type="image" name="imageField" src="imagenes/asignaturas.gif">
                <input type = "hidden" name = "CodigoRamo" value = "<%=RamosDebe("CodRamo")%>">
                <input type = "hidden" name = "CodigoSeccion" value = "<%=RamosDebe("codsecc")%>">
                </form>

               </td>
        
               
               </tr>
							<%
							 RamosDebe.MoveNext
						   Wend
						 End If %>
						  </table>
						  <br>
						 
					    <p></p></td>
					  </tr>
	</table>
   
    <table>
           <tr>
    <td width="83" align="left"><A onMouseOver="MM_swapImage('Image7','','Imagenes/botones/volver-on.gif',1)" onmouseout=MM_swapImgRestore() href="javascript: window.top.location.href ='menu_consultas.asp'"><IMG src="Imagenes/botones/volver-of.gif" border="0"  name="Image7"></A></td>
  </tr>
  </table></table>
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
<%ObjetosLocalizacion("ContenidosSelClase.asp")%>
</html>
<%
RamosDebe.Close()
RamosSolici.Close()
%>
<!--#INCLUDE file="include/desconexion.inc" -->
