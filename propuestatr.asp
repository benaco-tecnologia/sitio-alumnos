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

strRamosDebe = "sp_listacarga_alumnoportal '" & CodCli & "','" & ano & "','" & periodo & "'"
RamosDebe.Open strRamosDebe, Session("Conn")


	
%>
<Script Language="JavaScript">
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
</Script><html>
<head> 
<title><%=Session("NombrePestana")%></title>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1"> 
</head>

<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript"  SRC="include/tooltip.js"></SCRIPT>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<%if Instr(Request.ServerVariables("HTTP_REFERER"), "resultado.asp") <> 0 then %>
	<script>
    window.onload=function(){
    var pos=window.name || 0;
    window.scrollTo(0,pos);
    }
	</script>
<%end if%>
<script>
	window.onunload=function(){
	window.name=self.pageYOffset || (document.documentElement.scrollTop+document.body.scrollTop);
	}
</script>
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
<%' if not InscritoNormal() then %>
  <script>
  //   alert("No ha inscrito aún vía solicitud Normal ...");
  </script>   
<%'  end if%>
<table width="1049" border="0"  align="left" cellpadding="0" cellspacing="0">
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
                
				  <form name="form1" method="post" action="asignatura-seccion.asp">
					<input type = "hidden" name = "Ramos" value = "">
					<table border="0" cellpadding="0" cellspacing="15" width="962">
					  <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
					  <tr valign="top" > 
						<td ><p><span style="font-size: 14px"><p id ="lblTituloInscripcion" style="font-size: 25px" class="text-menu">Propuesta de inscripci&oacute;n asignaturas</p></span></p></tr>
						<tr valign="top" > 
                        
						<td ><a href="javascript:imprime()"><img src="imag/f-der/impresion.gif" width="36" height="36" border="0" align="left"></a></tr>
					  <tr valign="top">
					  	<td height="1" colspan=3><span style="font-size: 13px; font-weight: bold;"><b style="font-weight: bold; color: #0099CC; font-size: 13px;"><font face="Verdana, Arial, Helvetica, sans-serif">Propuesta</font></b></span><span style="font-size: 13px; font-weight: bold; font-family: Verdana, Arial, Helvetica, sans-serif;"><b class="Tit-celdas"><font face="Verdana, Arial, Helvetica, sans-serif"> 
						  de inscripci&oacute;n de : &nbsp;&nbsp;&nbsp;</font></b></span><span class="Tit-celdas"><font class="Tit-celdas"><%=GetNombrealumno(Codcli)%> </font></span></td>
				      </tr>
					  <tr valign="bottom"> 
						<td width="727" class="text-normal-celdas"><font class="text-normal-celdas">Fecha:&nbsp;&nbsp;<%=date()%></font></span></b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
						
					  <tr valign="top"> 
						<td colspan="3" height="80"> 
						  <table width="556" cellspacing="1" cellpadding="0" height="59" border="0" bordercolor="#FFFFFF">
							<td height="20" colspan="11"> <div align="left"> 
								
							  </div></td>
							</tr>
							<tr background="imagenes/fdo-cabecera-cel22.jpg"> 
							  <td width="210" height="30"  background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center" ><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">C&oacute;digo Asignatura</font></b></font></div></td>
							  <td width="269" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Nombre Asignatura</font></b></font></div></td>
							</tr>
							<% 
						  If RamosDebe.Eof Then
						  %>
							<%
						  Else
						   While Not RamosDebe.Eof
%>
							<tr bgcolor="#DBECF2" height="25"> 
							  <td width="210" height="30" align="center" ><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("codramo")%></font></td>
							  <td width="269" height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("Nombre")%></font></td				 
							></tr>
							<%
							 RamosDebe.MoveNext
						   Wend
						 End If %>               
						  </table>
						  <br>
					    <p></p></td>
					  </tr>
      <tr>
        <td align="center" colspan="2">
          
         </td>
      </tr>
	</table>
    <table>
           <tr>
    <td width="83" align="left"><A onMouseOver="MM_swapImage('Image7','','Imagenes/botones/volver-on.gif',1)" onmouseout=MM_swapImgRestore() href="javascript: window.top.location.href ='menu_tomaderamos.asp'"><IMG src="Imagenes/botones/volver-of.gif" border="0"  name="Image7"></A></td>
  </tr>
  </table>
	</form></table>
	</td>
  </tr>
</table>


</body>
<script languaje = "javascript">

function ConfirmaModi()
{
	if (confirm("¿Est\u00e1 seguro de modificar su inscripci\u00f3n? "))
	   { 
		   window.location = "modificainscripcion.asp";
	   }
}

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
<%ObjetosLocalizacion("resultado.asp")%>
</html>
<%
RamosDebe.Close()
%>
<!--#INCLUDE file="include/desconexion.inc" -->
