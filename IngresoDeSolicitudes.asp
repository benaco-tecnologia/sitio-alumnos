<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->  
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE file="analytics.asp" -->
<script type="text/javascript">
function checkear(n)
    {
		var docto='Docto'+n
		var uploadedFile = document.getElementById(docto);
		var fileSize = uploadedFile.files[0].size;
		var nombrearchivo= uploadedFile.value;
		
		var ext_validas = /\.(jpg|gif|doc|docx|xls|xlsx|ppt|pptx|pdf)$/i.test(nombrearchivo);
		
		if(ext_validas == false)
		{
			alert("Seleccione un archivo con formato jpg, gif, doc, docx, xls, xlsx, ppt, pptx, pdf.");
			document.getElementById(docto).value ='';
			return;
		}
		
		if(nombrearchivo.indexOf("'") != -1 || nombrearchivo.indexOf("á") != -1 || nombrearchivo.indexOf("é") != -1 || nombrearchivo.indexOf("í") != -1 || nombrearchivo.indexOf("ó") != -1 || nombrearchivo.indexOf("ú") != -1)
		{ 
			alert("El nombre del archivo no puede llevar caracteres especiales como comillas o acentos.");
			document.getElementById(docto).value ='';
			return;
		} 

		/*if(nombrearchivo.indexOf("'") != -1){ 
			alert("El nombre del archivo no puede llevar caracteres especiales como comillas o acentos.");
			document.getElementById(docto).value ='';
			return;
					
		}*/
		
		if (fileSize>41943040)//40 MB		
		{
			var peso=Math.round(fileSize/1024);			
			alert('No puede subir este archivo ya que supera los 10 Mb.');
			document.getElementById(docto).value ='';
		} 		
    }
	
	function validaForm()
	{
		var txtSolicitud = document.getElementById('txtSolicitud').value;
		var txtMotivo = document.getElementById('txtMotivo').value;
		var mensaje = "";
        var EsCorrecto = true;
		
		if(txtSolicitud=="0") {
			if(EsCorrecto==false) {
				mensaje = mensaje + "\n- El campo Tipo es obligatorio.";
			}
			else {
				mensaje = "- El campo Tipo es obligatorio.";
				EsCorrecto = false;
			}
		}
		
		if(txtMotivo=="") {
			if(EsCorrecto==false) {
				mensaje = mensaje + "\n- El campo Motivo es obligatorio.";
			}
			else {
				mensaje = "- El campo Motivo es obligatorio.";
				EsCorrecto = false;
			}
		}
		
		if (EsCorrecto == false) {
			alert(mensaje);
			return false;
		}
		else {
			return true;
		}  
			
	}
</script>

<%
strParame="SELECT coalesce(dbo.Fn_ValorParame('USAANALYTICS'),'NO')Parame"
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
	USAANALYTICS=rstParame("Parame")
end if

if USAANALYTICS="SI" then
	analitycs() 
end if 

if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if

if EstaHabilitadaNW (898)="S" then 
	if GetPermisoNW(898) ="N" then
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

Audita 898,"Ingresa a Ingreso de solicitudes"


strF="select CONVERT(VARCHAR(10),GETDATE(),105) Fecha"
set rstF= Session("Conn").Execute(strF)		
if not rstF.eof then
	Fecha=rstF("Fecha")
end if

TipoDeSolicitud= Request.QueryString("S")
Codsede= Request.QueryString("SD")

if TipoDeSolicitud <> "" and TipoDeSolicitud <> "0" then
	
	'Verificamos si posee flujo
	strPF ="pa08_RA_DETRES_sel @TIPOSOLI= '"& TipoDeSolicitud &"', @CODROL = '"& Session("perfilNW") &"',@SECUENCIA ='1' "
	set rstPF= Session("Conn").Execute(strPF)
		
	if not rstPF.eof then

		strTP="SP_TRAE_DATOS_SOLICI_PA "& TipoDeSolicitud &","& Session("perfilNW") &""
	
		set rstTP= Session("Conn").Execute(strTP)		
		
		if not rstTP.eof then
			USACARRERADES = rstTP("CHCARRERA")	
			USACHSEDEDES  = rstTP("CHSEDE")	
		else
			USACARRERADES ="NO"
			USACHSEDEDES ="NO"
		end if
		
		TienePermisoSolicitud="SI"
		
	else
		TienePermisoSolicitud ="NO"
	end if
	
end if 
%>
<html>
<head><link rel="shortcut icon" type="image/x-icon" href="imagenes/2_r1_c2.jpg">
<title><%=Session("NombrePestana")%></title>
<link href="css/tex-normales.css" rel="stylesheet" type="text/css">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1"> 
<style type="text/css">
<!--
.Estilo19 {
	color: #FFFFFF;
	font-weight: bold;
}
.Estilo20 {color: #FFFFFF; font-weight: bold; font-size: 16px; }
-->
</style>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<link href="css/textos.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.Estilo21 {color: #000000}
.Estilo22 {font-size: 10px}
-->
</style>
<script language="JavaScript" type="text/javascript">


function CargarReporte()
{
	var opciones="toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=no, resizable=no, width=1024, height=768, top=85, left=125";
window.open("cuponera.asp","",opciones);
}


</script>
</head>
<body>
<table border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
	<td>
		<table width="794" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right"><% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td width="523" valign="top" >
			<% CargarTop1()%><% SubMenu()%> 
            	  <form method='post' enctype="multipart/form-data" name='frm' id='frm' action="solicitudes_graba.asp" onSubmit="">
                  <table width='100%'  border='0' cellpadding='5' cellspacing='1' bgcolor='#FFFFFF'>
                  
                  <tr valign="top" > 
						<td colspan="4" >
                        <span style="font-size: 14px"><p style="font-size: 25px" class="text-menu">Ingreso de solicitudes</p></span>
                        </td>				                        
                  </tr>
                  <%if TienePermisoSolicitud ="NO" then%>
                  <tr valign="top" > 
                  		<td colspan="4" >
                        <font color="#CC0033" size="2" face="Arial, Helvetica, sans-serif">*Su perfil no es parte del flujo de la solicitud seleccionada.</font>
                        </td>
                  </tr>
                  <%end if%>
									<tr valign="middle" class='caption'>
                                      
									  <td width="19%" height="30" align='right' background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda">*Tipo :</td>
                                      
                                     
									  <td height="30"  colspan='1' align="left" class="calResaltado">
                                      
                                      <select style="width:463px" name="txtSolicitud"  id="txtSolicitud" onChange="javascript:Actualizar();">
									  <option value='0'>Selecione una solicitud</option>					
                                      <%
									  strSol ="sp_traeSolicitudesPA "& Session("perfilNW")& ""
									  Set rstSol = Session("Conn").Execute(strSol)
									  While Not rstSol.Eof
									  
									  if TipoDeSolicitud = rstSol(0) and TienePermisoSolicitud ="SI" then%>
                                      	  <option value='<%=rstSol(0)%>' selected><%=rstSol(1)%></option>
                                      <%else%>
                                      	  <option value='<%=rstSol(0)%>'><%=rstSol(1)%></option>
                                      <%end if
									  
									  rstSol.movenext
									  wend 
									  %>
                                      
									  </select>
                                      </td>
                                      
									</tr>
									<tr valign="middle" class='caption'>
									  <td height="30" align='right' background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda">Número:</td>
									  <td height="30" colspan='2' align="left">
									  <input  style="width:463px" name="txtNumero" type="text" class="calCeldaResaltado" id="txtNumero" size="55" disabled="disabled">                                 
                                      </td>
									  </tr>
									<tr valign="middle" class='caption'>
									  <td height="30" align='right' background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda">Fecha:</td>
									  <td height="30" colspan='1' align="left">
									  <input style="width:463px" name="txtFechaM" type="text" class="calCeldaResaltado" id="txtFechaM" size="55" value="<%=Fecha%>" disabled="disabled">
                                      <input type="hidden" name="txtFecha" id="txtFecha" value="<%=Fecha%>" />
                                      </td>
									  </tr>
                                      
                                      
                                      <%if USACHSEDEDES  ="SI" then%> 
                                     <tr valign="middle" class='caption'>
									  <td width="19%" height="30" align='right' background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda">Sede de destino :</td>
									  <td height="30"  colspan='1' align="left" class="calResaltado">
                                      <select  style="width:463px" name="txtCodsedeDes"  id="txtCodsedeDes" onChange="javascript:ActualizarMasSede();">
									  <option value='0'>Selecione una Sede</option>					
                                      <%
									  strSede ="SELECT CODSEDE,CODSEDE +' - '+DBO.ACENTO_ESPECIAL(DESCRIPCION)DESCRIPCION FROM RA_SEDE"
									  Set rstSede = Session("Conn").Execute(strSede)
									  While Not rstSede.Eof
									  
									  if Codsede = rstSede(0) and TienePermisoSolicitud ="SI" then%>
                                      	<option value='<%=rstSede(0)%>' selected><%=rstSede(1)%></option>
                                      <%else%>
                                      	<option value='<%=rstSede(0)%>'><%=rstSede(1)%></option>
                                      <%end if%>
                                    
                                      <%
									  rstSede.movenext
									  wend 
									  %>
                                      
									  </select>
                                      </td>
									</tr> 
                                    <%end if%> 
                                    
                                    <%if USACARRERADES  ="SI" then%> 
                                    <tr valign="middle" class='caption'>
									  <td width="19%" height="30" align='right' background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda">Carrera de destino :</td>
									  <td height="30"  colspan='1' align="left" class="calResaltado">
                                      <select  style="width:463px" name="txtCarreraDes"  id="txtCarreraDes">
									  <option value='0' Selected>Selecione una Carrera</option>					
                                      <%
									  strCarr ="SELECT CODCARR,CODCARR +' - '+DBO.ACENTO_ESPECIAL(NOMBRE_C)NOMBRE_C FROM MT_CARRER WHERE SEDE ='"& Codsede &"'"
									  Set rstCarr = Session("Conn").Execute(strCarr)
									  While Not rstCarr.Eof
									  %>
                                      <option value='<%=rstCarr(0)%>'><%=rstCarr(1)%></option>
                                      <%
									  rstCarr.movenext
									  wend 
									  %>
                                      
									  </select>
                                      </td>
									</tr>
                                    <%end if%>
                                    
                                      <tr valign="middle" class='caption'>
									  <td height="30" align='right' background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda">*Motivo:</td>
									  <td height="30" colspan='1' align="left">
									  <input style="width:463px" name="txtMotivo" type="text" class="calCeldaResaltado" id="txtMotivo" size="55" >                                      </td>
									  </tr>
                                      
                                      <tr valign="middle" class='caption'>
                                      <td height="30" align='right' background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda">Glosa:</td>
                                      <td colspan='2' rowspan="3" align="left"><textarea name="txtGlosa" cols="55" rows="5" wrap="virtual" class="calCeldaResaltado" id="txtGlosa"></textarea></td>
									  </tr>
                                      <tr valign="middle" class='caption'>
                                      </tr>
                                       <tr valign="middle" class='caption'>
                                      </tr>
                                      
									<tr valign="middle" class='caption'>
									  <td height="30" align='right' background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda">Adjunto(1):</td>
									  <td height="30" colspan='1' align="left"><input onChange="checkear(1);" name="Docto1" type="file" class="calCeldaResaltado" id="Docto1" size="50" maxlength="50"></td>
									  </tr>
									<tr valign="middle" class='caption'>
									  <td height="30" align='right' background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda">Adjunto(2):</td>
									  <td height="30" colspan='1' align="left"><input onChange="checkear(2);" name="Docto2" type="file" class="calCeldaResaltado" id="Docto2" size="50" maxlength="50"></td>
									  </tr>
									<tr valign="middle" class='caption'>
									  <td height="30" align='right' background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda">Adjunto(3):</td>
									  <td height="30" colspan='1' align="left"><input onChange="checkear(3);"  name="Docto3" type="file" class="calCeldaResaltado" id="Docto3" size="50" maxlength="50"></td>
									  </tr>		
                                      			
								</table>
                                </br>
              <table>
              
                       <tr>
                            <td width="83" align="left"><A onMouseOver="MM_swapImage('Image7','','Imagenes/botones/volver-on.gif',1)" onmouseout=MM_swapImgRestore() 				                 href="javascript: window.top.location.href ='menu_consultas.asp'"><IMG src="Imagenes/botones/volver-of.gif" border="0"  name="Image7">                 </A>
                            </td>
                            <td width="83" align="left"><A onMouseOver="MM_swapImage('Image123','','Imagenes/botones/grabar-on.gif',1)" onmouseout=MM_swapImgRestore() 				                 href="javascript: document.frm.submit();" onClick="return validaForm();"><IMG src="Imagenes/botones/grabar-of.gif" border="0"  name="Image123">                 </A>
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

<script>
function Actualizar() {
  var txtSolicitud = document.getElementById('txtSolicitud').value;
  window.top.location="IngresoDeSolicitudes.asp?S=" + txtSolicitud +"";
}

function ActualizarMasSede() {
  var txtSolicitud = document.getElementById('txtSolicitud').value;
  var txtCodsedeDes = document.getElementById('txtCodsedeDes').value;
  window.top.location="IngresoDeSolicitudes.asp?S=" + txtSolicitud +"&SD=" + txtCodsedeDes +" ";
}
</script>

<% if Request.QueryString("M") = "OK" then%>
<script>alert('Se ha registrado correctamente la solicitud ingresada.');</script>
<%end if%>

<!--#INCLUDE file="include/desconexion.inc" -->