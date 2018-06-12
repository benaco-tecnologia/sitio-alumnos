<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%
Dim CambioClave
CambioClave = Request("CambioClave")
variable=request("sw") 
'response.write(variable)
'response.end
if variable = 1 then
   session("Entrada")= 1
end if   


strParameCO="SELECT coalesce(dbo.Fn_ValorParame('ENVIACORREOCAMBIOPASSPA'),'')Parame"
set rstParameCO= Session("Conn").Execute(strParameCO)		
if not rstParameCO.eof then
		ENVIACORREOCAMBIOPASSPA = rstParameCO("Parame")
	else
		ENVIACORREOCAMBIOPASSPA="" 
end if

if ENVIACORREOCAMBIOPASSPA = "SI" then	
	'logica envio de correo
	EnviaMail = 1
	
	strSol ="SP_TRAE_DATOS_SOL_PA '"& session("codcli") &"'"
	if bcl_ado(strSol,Rstsol)  then
		SolNombre = Rstsol("NOMBRE")
		SolMail = Rstsol("MAIL")
		 
		if SolMail = "SINMAIL" then
			EnviaMail = 0
		end if	
		
		strVM="SELECT dbo.fn_validamail('"& SolMail &"')mail"
		set rstVM= Session("Conn").Execute(strVM)
			
		if not rstVM.eof then
			MalBueno=rstVM("mail")
			if MalBueno = 0 then
				EnviaMail = 0
			end if 
		end if 
		
		Body =  "Estimado(a) " & SolNombre &" "& SolPaterno & " <BR><BR> Su clave de acceso al portal de Alumnos se ha cambiado correctamente. <BR> Saludos."	
		
		if EnviaMail = 1 then
			call enviargmail(SolMail,Body)
		end if			
	end if 
	
end if 

function enviargmail(destino,texto) 

Const cdoSendUsingMethod = "http://schemas.microsoft.com/cdo/configuration/sendusing"
Const cdoSendUsingPort = 2
Const cdoSMTPServer = "http://schemas.microsoft.com/cdo/configuration/smtpserver"
Const cdoSMTPServerPort = "http://schemas.microsoft.com/cdo/configuration/smtpserverport"
Const cdoSMTPConnectionTimeout = "http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout"
Const cdoSMTPAuthenticate = "http://schemas.microsoft.com/cdo/configuration/smtpauthenticate"
Const cdoBasic = 1
Const cdoSendUserName = "http://schemas.microsoft.com/cdo/configuration/sendusername"
Const cdoSendPassword = "http://schemas.microsoft.com/cdo/configuration/sendpassword"
Const cdoSMTPUseSSL = "http://schemas.microsoft.com/cdo/configuration/smtpusessl"
Dim objConfig 
Dim objMessage 
Dim Fields  

Set objConfig = Server.CreateObject("CDO.Configuration")
Set Fields = objConfig.Fields


serverMail=Application("server")
portMail=Application("port")
userMail=Application("user")
passMail=Application("pass")
sslMail=Application("ssl")  

With Fields
.Item(cdoSendUsingMethod) = cdoSendUsingPort
.Item(cdoSMTPServer) = serverMail
.Item(cdoSMTPServerPort) = portMail
.Item(cdoSMTPConnectionTimeout) = 10
.Item(cdoSMTPAuthenticate) = cdoBasic 
.Item(cdoSendUserName) = userMail
.Item(cdoSendPassword) = passMail
.Item(cdoSMTPUseSSL) = sslMail 
.Update
End With 
Set objMessage = Server.CreateObject("CDO.Message")
Set objMessage.Configuration = objConfig
With objMessage
.To = destino
.From = userMail
.Subject = "Recupera clave portal alumnos"
.HTMLBody = texto 
.Send 
End With

If Err=0 Then 
	'response.Write "Se envia correo"
Else
	'Response.Write "<html><body><h1>The following error occured when sending</h1>Error (" & Err & ") :" & Err.Description & "</body></html>"
End If
Set Fields = Nothing
Set objMessage = Nothing
Set objConfig = Nothing
End function

%>
<html>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript"  SRC="include/validarut.js"></SCRIPT>
<script language="JavaScript">
function Validar(campo1, campo2, campo3){
if (campo1 != "" && campo2 != "" && campo3 != ""){
  if (ValidaCampo(campo1) && ValidaCampo(campo2))
      {
	   if (ComparaClaves(campo1,campo2))
	   {
	       window.document.claves.action = "CambioClaveSql.asp";
		   window.document.claves.submit();
	   }
	   else
	      return false;
     }
	 else
	     return false;  
} 
else {
   alert("Por favor ingrese todos los datos solicitados.");
   return false;
  }
}

</script>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title>Portal de Alumnos en L&iacute;nea</title>
<table border="0" cellpadding="0" cellspacing="0"  align="left">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right" ><% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td valign="top">
			<% CargarTop1()%><% SubMenu()%>
			<table width="100%" border="0" cellspacing="0" cellpadding="20" height="550" bgcolor="#FFFFFF">
			  <tr> 
			<td valign="top" align="left">
			   	<div align="center">
             		<table border="0" cellpadding="0" cellspacing="0" height="53" width="584" align="left">
					<tr align="left"> 
					  <td width="336">&nbsp;</td>
					  <td width="248">&nbsp;</td>
					</tr>
					<tr align="left"> 
					  <td width="336"><img src="imagenes/titulos/T-cambio-password-3.gif" width="299" height="66"></td>
					  <td width="248">&nbsp;</td>
					</tr>
					<tr align="left"> 
					  <td width="336">&nbsp;</td>
					  <td width="248">&nbsp;</td>
					</tr>
					<tr align="left" valign="top">
					  <td colspan="2" height="0"><a href="salir.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('_r11_c21','','Imagenes/botones/A-_r12_c2_f2.jpg',1)" ><img src="Imagenes/botones/A-_r12_c2.jpg" name="_r11_c21" width="162" height="21" border="0" align="left" id="_r11_c21"></a>					  </td>
					</tr>
			     </table>
			  </div>
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
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
