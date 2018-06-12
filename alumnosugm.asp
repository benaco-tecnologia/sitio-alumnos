
<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title>Portal de Alumnos en L&iacute;nea </title>
<body >
<table width="1024" border="0" align="left" cellpadding="0" cellspacing="0" background="Imagenes/ugm/fondo.jpg" style="background-repeat:no-repeat;">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0">
		  <tr>
			<td width="0" valign="top" align="left" ><table width='198' height='136' border='0' cellpadding='0' cellspacing='0' aling ='right'><tr valing = 'top'><td valign='top'><img name='A_r1_c1' src='Imagenes/ugm/2_r1_c1.png' width='18'height='136' border='0' alt=''></td><td valign='top' ><img name='A_r1_c2' src='Imagenes/ugm/2_r1_c2.png' width='162' height='136' border='0' alt=''></td><td valign='top'><img name='A_r1_c3' src='Imagenes/ugm/2_r1_c3.png' width='18' height='136' border='0' alt=''></td></tr></table>
    <script language="JavaScript" type="text/javascript" src="include/validarut.js"></script>
   
<style type="text/css">
body {
    overflow:hidden;
}
</style>
<!--<script language="JavaScript" type="text/javascript" src="include/validarut.js"></script>-->

<script language="JavaScript" type="text/javascript">
<!--
function setFocus() {
  var loginForm = document.getElementById("login");
  if (loginForm) {
	loginForm["logrut"].focus();
  }
}
function Ingresar()
{
	document.login.action="validaclave.asp";
	document.login.submit();
}
function RecuperaClave()
{
 var texto = document.getElementById("logrut").value; 
	
 var tmpstr = "";
 

  for ( i=0; i < texto.length ; i++ )
    if ( texto.charAt(i) != ' ' && texto.charAt(i) != '.' && texto.charAt(i) != '-' )
      tmpstr = tmpstr + texto.charAt(i);
  texto = tmpstr;
  largo = texto.length;


  tmpstr = "";
  for ( i=0; texto.charAt(i) == '0' ; i++ );
  for (; i < texto.length ; i++ )
     tmpstr = tmpstr + texto.charAt(i);
  texto = tmpstr;
  largo = texto.length;

  //alert(largo)

  if ( largo < 2 )
  {
    alert("Debe ingresar el RUT completo.");
    window.document.login.logrut.focus();
    window.document.login.logrut.select();
    return false;
  }


  for (i=0; i < largo ; i++ )
  {
    if ( texto.charAt(i) !="0" && texto.charAt(i) != "1" && texto.charAt(i) !="2" && texto.charAt(i) != "3" && texto.charAt(i) != "4" && texto.charAt(i) !="5" && texto.charAt(i) != "6" && texto.charAt(i) != "7" && texto.charAt(i) !="8" && texto.charAt(i) != "9" && texto.charAt(i) !="k" && texto.charAt(i) != "K" )
    {
      alert("El RUT ingresado no es válido.");
      window.document.login.logrut.focus(); 
      window.document.login.logrut.select();
      return false;
    }
  }
  window.location ="Alumnosugm.asp?rc=s&rut="+texto;
  //location.href="RecuperaClave.asp?rut="&texto;
  
}
<!--
function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
setFocus();
 var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
   var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
   if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

//-->

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}
//-->
</script>

<script languaje="javascript">
    function Enviar3(pag, valor, validaencuesta) {
        if (validaencuesta == 'SI') {
            if (valor == 0) {
                alert("Para acceder a esta opci\u00f3n debe responder sus Encuestas Pendientes");
                return;
            }
            top.location.href = pag;
        }
        else {
            top.location.href = pag;
        }
    }

function isNumber(n) {
  return !isNaN(parseFloat(n)) && isFinite(n);
}
function myTrim(x) {
    return x.replace(/^\s+|\s+$/gm,'');
}
function validarRUC(valor) {
    var valor = myTrim(valor) 	
    if (isNumber(valor)) {
        if (valor.length == 8) {
            suma = 0
            for (i = 0; i < valor.length - 1; i++) {
                digito = valor.charAt(i) - '0';
                if (i == 0) suma += (digito * 2)
                else suma += (digito * (valor.length - i))
            }
            resto = suma % 11;
            if (resto == 1) resto = 11;
            if (resto + (valor.charAt(valor.length - 1) - '0') == 11) {
                return true				
            }
        } else if (valor.length == 11) {
            suma = 0
            x = 6
            for (i = 0; i < valor.length - 1; i++) {
                if (i == 4) x = 8
                digito = valor.charAt(i) - '0';
                x--
                if (i == 0) suma += (digito * x)
                else suma += (digito * x)
            }
            resto = suma % 11;
            resto = 11 - resto
            if (resto >= 10) resto = resto - 10;
            if (resto == valor.charAt(valor.length - 1) - '0') {
                return true
            }
        }
    }
	alert('El Ruc ingresado es incorrecto.');
	window.document.login.logrut.value ='';
    return false;
}

</script>

<style type="text/css">
    <!
    -- body
    {
        margin-left: 0px;
        margin-top: 0px; 
    }
    -- ></style>
    
 <%
rc = Request.QueryString("rc")
rutRC = Request.QueryString("rut")

if rc = "s" then

strSql = "SELECT dbo.Decrypt(us_password) AS clave,pe_nombrecompleto AS nombre,COALESCE(pe_email,'') as mail FROM ca_usuarios ,tg_personas " &_
"WHERE ca_usuarios.id_persona = tg_personas.id_persona and " &_
"us_consuser = '" & rutRC & "' "

Set rscl = Session("Conn").execute(strSql)

	if rscl.eof then
		Response.Write("<script language='javascript'>alert('Lo siento el Rut que ha Ingresado no se encuentra en nuestros registros, vuelva a Intentar.'); window.location.href='"  + pagina + "';</script>")
		response.End()
	else
		
		if rscl("mail") = "" then
			Response.Write("<script language='javascript'>alert('Usted no tiene un correo electrónico asociado a su cuenta, Comuníquese con la unidad académica para actualizar sus datos.'); window.location.href='"  + pagina + "';</script>")
			response.End()
		end if		
		
		strParame="SELECT dbo.Fn_ValorParame('PORTADA_PROPIA_PORTAL_ALUMNO')Parame"
		set rstParame= Session("Conn").Execute(strParame)
		if not rstParame.eof then
			pagina=rstParame("Parame")
		else
			pagina=""
		end if
		
		dim Sql , Body
		Body =  chr(34) & "Estimado(a) " & rscl("nombre") & " su clave de acceso al portal de Alumnos es: '" & rscl("clave") &"' (omitir comillas)" & chr(34) 	
		
		call enviargmail(rscl("mail"),Body)  	
		
		session("RecuperaClaveMail") = rscl("mail")
		
		if(pagina <> "" and not isNull(pagina)) then
			Response.Write("<script language='javascript'>alert('Su clave de acceso al Portal de Alumnos ha sido enviada a su correo electr\u00f3nico.'); window.location.href='"  + pagina + "';</script>")
			response.End() 
		else
			response.Redirect("RecuperaClave.asp")
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
<link href="css/tablas.css" rel="stylesheet" type="text/css">
<body onLoad="MM_preloadImages('Imagenes/ugm/-_r11_c2_f2.html')">
    
    <table border="0" cellpadding="0" cellspacing="0" align="right">
        <tr>
            <td colspan="3" valign="top">
                <table width="198" border="0" cellspacing="0" cellpadding="0" height="559">
                    <tr valign="top">
                        <td width="18" height="449">&nbsp;</td>
                        <td width="162" height="449">
                            <table width="162" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                    <td>
                                        <a href="#" onMouseOut="MM_swapImgRestore()" onClick="return RecuperaClave();" onMouseOver="MM_swapImage('_r9_c2','','Imagenes/ugm/Botones/-_r9_c2_f2.png',1);">
                                            <img name="_r9_c2" src="Imagenes/ugm/botones/-_r9_c2.png" width="162" height="27" border="0"></a>
                                    </td>
                                </tr>
                                <tr>
                                    <td background="Imagenes/ugm/-_r10_c2.png" valign="bottom" height="139">
                                        <table border="0" cellspacing="0" cellpadding="0" width="130" align="center" height="79">
                                            <form name="login" method="post"  onSubmit="javascript:return ValidaDatos();">
                                            <input type="hidden" name="rut" value="">
                                            <input type="hidden" name="pin" value="">
                                            <tr>
                                                <td height="30" width="130">
                                                    <div align="center">
                                                   
                                                        <!--<input name="logrut" type="text" class="casillas-form" id="logrut" size="13" onChange="javascript:checkRutField(this.value);"
                                                            value="" maxlength="12">-->
                                                        <input name="logrut" type="text" class="casillas-form" id="logrut" size="13" onChange="javascript:checkRutField(this.value);" value="" maxlength="12">  
                                                    </div>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td height="10" width="130">
                                                    <div align="center">
                                                    </div>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td height="39" width="130">
                                                    <div align="center">
                                                        <input name="logclave" type="password" value="" class="casillas-form" id="logclave"
                                                            size="13" maxlength="30">
                                                    </div>
                                                </td>
                                            </tr>
                                            </form>
                                        </table>
                                    </td>
                                </tr>
                                <tr>
                                    <td align="center">
                                        <a target="_top" href="javascript:Ingresar();return ValidaDatos();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('_r11_c21','','Imagenes/ugm/botones/A-_r12_c2_f2.png',1)" onFocus="javascript:Ingresar();return ValidaDatos();" >
                                            <img src="Imagenes/ugm/botones/A-_r12_c2.png" name="_r11_c21" width="162" height="21" border="0" id="_r11_c21" >
                                            </a>
                                    </td>
                                </tr>
                                <tr>
                                    <td valign="top" width="162" height="23">&nbsp;</td>
                                </tr>
                                <!--<tr><td align="center">
                    <a href="#" id="needone" onClick="return RecuperaClave();">Recuperar clave</a></td>
                    </tr> -->
                                <td valign="top" width="162" height="149"> </td>
                            </table>
                        </td>
                        <td height="449" width="18" valign="top"></td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <img src="Imagenes/ugm/spacer.gif" width="1" height="28" border="0">
    </td>
		    <td valign="top" align="left" ><table width='704' height='0' border='0' cellpadding='0' cellspacing='0'><td valign='top'><img name='A_r1_c4' src='Imagenes/ugm/A_r1_c4.png' width='826' height='80' border='0' alt=''></td></table>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
-->
</style> 

<table border="0" cellpadding="0" cellspacing="0" width="702">
  <tr> 
    <td width="702" height="480" colspan="5" valign="top"><img name="_r3_c4" src="Imagenes/ugm/A-_r3_c4.png" width="826" height="480" border="0"></td>
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

