<%
  'Response.Expires = -1 
  'Response.Buffer = false
    Response.Expires = -1
    Session.Timeout = 999
    Response.Buffer = false
    Server.ScriptTimeout = 99999
%>
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%
	if Session("CodCli") = "" then
	  Response.Redirect("saltoinicio.htm")
	end if

	if EstaHabilitadaNW (475)="S" then 
		if GetPermisoNW(475) ="N" then
			response.Redirect("MensajesBloqueos.asp")		
		end if
	else
		response.Redirect("MensajeBloqueoHabilita.asp")
	end if 
	
    Audita 475,"Ingresa a Evaluacion Docente" 
	
	strRut= trim(request("txtRut"))
	strDig = trim(request("txtDig"))
	strNombre= trim(request("txtNombre"))
	strCarrera = trim(request("txtCarrera"))
	strAno = trim(request("txtAno"))
	strPeriodo= trim(request("txtPeriodo"))
	strFecha = trim(request("txtFecha"))
	Accion= request.form("Accion")
	
	if Accion="Grabar" then     ' Actualiza Confirmacion  
	   sql="Update mt_Alumno set EncDoc='S',Ano_Ed='" & session("Ano_ed") & "',Periodo_Ed='" & session("Periodo_Ed") & "'"
	   Sql = Sql & " where Codcli='" & Session("CodCli") & "'"
	   Sql = Sql & " and estacad='VIGENTE'"
	   
	   Session("Conn").execute(Sql)
	   %><script>
	     //parent.parent.mainFrame.location.href="adm-acad.asp"
		 //parent.parent.leftFrame.location.href="f-izq.asp"
 		 	window.top.location.href ='menu_encuestas.asp'
		 </script>
	   <%
	end	if
 	if Accion="Omitir" then     ' Actualiza Confirmacion  
	   sql="Update mt_Alumno set EncDoc='S',Ano_Ed='" & session("Ano_ed") & "',Periodo_Ed='" & session("Periodo_Ed") & "'"
	   Sql = Sql & " where Codcli='" & Session("CodCli") & "'"
	   Sql = Sql & " and estacad='VIGENTE'"
	   Session("Conn").execute(Sql)
	   %><script>
	     //parent.parent.mainFrame.location.href="adm-acad.asp"
		 //parent.parent.leftFrame.location.href="f-izq.asp"
		window.top.location.href ='menu_encuestas.asp'
		 </script>
	   <%
	end	if
    StrSql = "select a.codcli,a.codcarpr,a.ano_mat,a.periodo_mat,a.fec_mat,a.ano_mat,a.periodo_mat,"
    StrSql = StrSql & " a.Rut,b.dig  from mt_alumno a,mt_client b "
	StrSql = StrSql & " where a.Codcli='" & Session("CodCli") & "'"
	StrSql = StrSql & " and a.estacad='VIGENTE' and a.rut= b.codcli"
	
	'set rs=Session("Conn").execute(StrSql)
'	if not rs.eof then	   
	if bcl_ado(strsql,rs) then
	   strNombre= GetNombreAlumno(rs("codcli"))	   
	   strCarrera = GetNombreCarrera(rs("codcarpr"))
	   strAno = session("Ano_Ed")
	   strPeriodo= session("Periodo_Ed")
	   strFecha = rs("fec_mat")
	   'session("Ano")=rs("ano_mat")
	   'session("Periodo")=rs("periodo_mat")	   	 
	   strRut=rs("Rut")  
	   strDig=rs("Dig")
	   session("CodCarr")=rs("codcarpr")
	   %>
				 <script>				   	  
				    parent.bottomFrame.location.href="pie.htm"
					//parent.parent.bottomFrame1.location.href="instrucciones.htm"
					//parent.parent.bottomFrame2.location.href="pregunta-opc.asp";					
				 </script>
				 <%
	else
	   strRut= trim(request("txtRut"))
	   strDig = trim(request("txtDig"))
	   strNombre= ""
	   strCarrera = ""
	   strAno = ""
   	   strPeriodo= ""
	   strFecha = ""
	end if
	
  'function ValidaRamosPendientes()     
	'  ValidaRamosPendientes=1
	'  sql="Select * from ra_encuestados where Ano=" & session("Ano") & ""
    '  sql=sql & " and Periodo='" & session("Periodo") & "' and CodCli='" & session("CodCli") & "'"
    '  sql=sql & " and Evaluado='N'"  
	'  set rst=Session("Conn").execute(sql)		 
	'  if not rst.eof then
	'     ValidaRamosPendientes=0
	'  else
	'  	 ValidaRamosPendientes=1
	'  end if    

  'end function

 
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

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="ed.css" type="text/css">
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" bgcolor="#FFFFFF" text="#000000" onLoad="MM_preloadImages('ima/bot/confirmar-on.gif','ima/bot/salir-on.gif','ima/bot/ayuda-on.gif','ima/bot/no-deseo-of.gif')">
<table border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="left">
		  <tr>
		  	<td valign="top"><% CargarTop1()%><% SubMenu()%>
			<form name="TheForm" method="post" action="cabecera.asp">
			<input type="hidden" name="Accion" value="<%=Accion%>">
			  <table width="800" border="0" cellspacing="" cellpadding="0">
				<tr align="center"> 
				  <td width="182" height="10"> <div align="left"><br>
					</div></td>
				 <% if Session("HABILITAENCUESTA")=0 then%>
				  <td width="182"><a href="javascript:Omitir();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image4','','Imagenes/Botones/no-deseo-realiz-enc-on.gif',1)"><img src="Imagenes/Botones/no-deseo-realiz-enc-of.gif" name="Image4" width="246" height="45" border="0"></a></td>
				  <%end if%>
				  <td width="17" height="10"> <div align="center"></div></td>
				  <td width="80" height="10"> <a  href="javascript:Grabar('<%=ValidaRamosPendientes(session("CodCli"),session("Ano_ed"),session("Periodo_ed"))%>');" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','Imagenes/botones/confirmar-on.jpg',1)"><img src="Imagenes/botones/confirmar-of.jpg" width="178" height="45" name="Image1" border="0" ></a></td>
				  <td width="15" height="10">&nbsp;</td>
				  <td width="80" height="10"> <div align="center"> <a href="javascript:Salir();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image2','','Imagenes/botones/salir-on.gif',1)"><img src="Imagenes/botones/salir-of.gif" width="178" height="45" name="Image2" border="0"></a></div></td>
				  <td width="15" height="10">&nbsp;</td>
				  <td width="80" height="10"> <div align="center"> <a href="javascript:Intruc();" onMouseOver="MM_swapImage('Image3','','Imagenes/botones/ayuda-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="Imagenes/botones/ayuda-of.gif" width="178" height="45" name="Image3" border="0"></a></div></td>
				  <td width="103" height="10"> <div align="center"></div></td>
				</tr>
			  </table>
			<table width="800" border="0" cellspacing="15" cellpadding="0" height="80">
			  <tr valign="bottom"> 
				<td width="380" height="80" class="sub-titulo"> 
				  <table width="400" border="0" cellspacing="0" cellpadding="0" height="80">
					<tr valign="bottom"> 
					  <td colspan="2" height="20"> 
						<p align="left"><span id="lblTituloEncuesta" style="font-size: 25px"  class="text-menu">Antecedentes del Alumno</span> 
						</p>
						<!--<p><span style="font-size: 14px"><p id ="lblTituloInscripcion" style="font-size: 25px" class="text-menu">Inscripci&oacute;n de Asignaturas</p></span></p>-->
					  </td>
					</tr>
					<tr valign="bottom"> 
					  <td width="52" height="20" class="Tit-celdas"> 
						<p align="left" ><font id="lblCI" class="Tit-celdas">CI</font></p>
					  </td>
					  <td width="312" height="20"> 
						<div align="left"><span class="tex"> 
							<input type="text" name="txtRut" maxlength="10" size="15" value="<%=strRut%>" disabled="true">
                            <%IF trim(ucase(session("TIPOVALIDACIONRUT"))) ="PERUANA" then%>
							<%else%>
	                            <input type="text" name="<%=strDig%>" maxlength="1" size="2" value="<%=strDig%>" disabled="true">
							<%end if %>							
						  </span></div>
					  </td>
					</tr>
					<tr valign="bottom"> 
					  <td width="52" height="20" class="Tit-celdas"> 
						<div align="left"><span class="Tit-celdas">Nombre</span></div>
					  </td>
					  <td width="312" height="20"> 
						<div align="left"><span class="tex"> 
							<input type="text" name="txtNombre" maxlength="50" size="63" value="<%=strNombre%>" disabled="true">
						  </span></div>
					  </td>
					</tr>
					<tr valign="bottom"> 
					  <td width="52" height="20" class="Tit-celdas"> 
						<div align="left"><span class="Tit-celdas">Carrera</span></div>
					  </td>
					  <td width="312" height="20"> 
						<div align="left"><span class="tex"> 
							<input type="text" name="txtCarrera" maxlength="50" size="63" value="<%=strCarrera%>" disabled="true">
						  </span></div>
					  </td>
					</tr>
				  </table>
				</td>
				<td class="sub-titulo" width="380" height="80">
               <table width="400" border="0" cellspacing="0" cellpadding="0" height="80">
					<tr valign="bottom"> 
					  <td colspan="2" height="20"> 
						<p align="left">&nbsp;</p>
					  </td>
					</tr>
					<tr valign="bottom"> 
					  <td width="145" height="20" class="Tit-celdas"> 
						<p align="left" ><font  class="Tit-celdas">A&ntilde;o Encuesta</font></p>
					  </td>
					  <td width="255" height="20"> 
						<div align="left"><span class="tex"> 
							<input type="text" name="txtAno"   maxlength="4" size="10" value="<%=strAno%>" disabled="true">
						</span></div>
					  </td>
					</tr>
					<tr valign="bottom"> 
					  <td width="145" height="20" class="Tit-celdas"> 
						<div align="left"><span id="lblPeriodoEnc" class="Tit-celdas">Periodo Encuesta</span></div>
					  </td>
					  <td width="255" height="20"> 
						<div align="left"><span class="tex"> 
							<input type="text" name="txtPeriodo" maxlength="1" size="10" value="<%=strPeriodo%>" disabled="true">
						  </span></div>
					  </td>
					</tr>
					<tr valign="bottom"> 
					  <td width="145" height="20" class="Tit-celdas"> 
						<div align="left">Fecha Matriculado</div>
					  </td>
					  <td width="255" height="20"> 
						<div align="left"><span class="tex"> 
							<input type="text" name="txtFecha" maxlength="10" size="10" value="<%=left(strFecha,10)%>" disabled="true">
						  </span></div>
					  </td>
					</tr>
				  </table>

				</td>
			  </tr>
			</table>
			</form>
			</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>			
</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
<%ObjetosLocalizacion("cabecera.asp")%>
<script language="JavaScript">

function Grabar(valor)
{ 
 if (valor==0) {
   alert("Debe Completar la Encuesta para Confirmar");
   //alert(valor);
   return;
 }

  document.TheForm.Accion.value = "Grabar";	
  document.TheForm.submit();
}
function Salir()
{ 
//window.location.href="salir.asp"
window.top.location.href="menu_encuestas.asp"
}
function Intruc()
{ 
  var x=window.open ("Instrucciones.htm","HelpWindow","width=400,height=250,resizable=yes,scrollbars=yes,menubar=no,status=no");
}

function visualiza() {
  var x=window.open ("instrucciones.htm","Pauta de Evaluación Docente","width=400,height=400,resizable=yes,toolbar=yes");
 }
function Omitir()
{ 
   alert("Usted no contestarà la encuesta...!!");
   document.TheForm.Accion.value = "Omitir";	
   document.TheForm.submit();
}
</script>