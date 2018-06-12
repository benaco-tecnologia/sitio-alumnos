<%if session("MensajeLogin")<>"" then%>
<script language="JavaScript" TYPE="text/javascript">
	alert('<%=session("MensajeLogin")%>');
</script>
<%end if%> 
<html>
<%
	Session("RutCli")=""
	Session("RutCliente")= ""
	Session("MiClave")=""
	Session("ClaveEncriptada")=""
	Session("MiValor")=0
	Session("NomAlum") = ""
	Session("RutAlum") = ""
	Session("Rut") = ""
	Session("CodCli") = ""
	session("Codsede")=""
	session("Codcarr")=""
	Session("codpestud")=""
	Session("anoalumno") = ""
	Session("Jornada") = ""
	Session("Nivel") = ""
    Session("Logueado") = ""
	Session("BloqueoA") = ""
	Session("BloqueoF") = ""
	Session("estado") = ""
	Session("EstadoMatriculado") = ""
	session("Cerrada") = ""
	Session("GetCarr") = ""
	session("CarreraAlumno") = ""
	session("NumeroMaximoPrueba") = ""
	session("MensajeLogin")=""
%>
<%
dim objMyDLL,liveStr
rc = Request.QueryString("rc")
rutRC = Request.QueryString("rut")

if rc = "s" then

strSql = "SELECT dbo.Decrypt(us_password) AS clave,pe_nombrecompleto AS nombre,COALESCE(pe_email,'') as mail FROM ca_usuarios ,tg_personas"
StrSql = strSql & " WHERE ca_usuarios.id_persona = tg_personas.id_persona and  us_consuser = '" & rutRC & "' "

Set rscl = Session("Conn").execute(strSql)
	 if rscl.eof then
	 	'Response.Redirect("MensajeNoExiste.asp")
		session("MensajeLogin")="Lo siento, el Rut que ha Ingresado no se encuentra en nuestros registros....!!!\nVuelva a Intentar"
		Response.Redirect "alumnosfinning.asp" 
		Response.End()
	 else
	 
	 if rscl("mail") = "" then
		'Response.Redirect("MailError.asp")
		session("MensajeLogin")="Usted no tiene una cuenta de correo asociado a su cuenta, Comuníquese con la unidad académica para actualizar sus datos."
		Response.Redirect "alumnosfinning.asp" 
		Response.End()
	 end if		


	dim Sql , Body
	Body =  chr(34) & "Estimado(a) " & rscl("nombre") & " su clave de acceso al portal de Alumnos es: '" & rscl("clave") &"' (omitir comillas)" & chr(34) 	
	
	Sql = "Exec Pr_RecuperaClave " & Body & ", " &  chr(34) &  rscl("mail") & chr(34) & ", " &  chr(34) &  Application("server") & chr(34) &", " &  chr(34) &  Application("port") & chr(34) & ", " &  chr(34) &  Application("ssl") & chr(34)& ", " &  chr(34) &  Application("from") & chr(34)& ", " &  chr(34) &  Application("user") & chr(34)& ", " &  chr(34) &  Application("pass") & chr(34)& ", " &  chr(34) &  Application("cred") & chr(34)
			
		Session("Conn").Execute(Sql)
		
		session("RecuperaClaveMail") = rscl("mail")
		glosa="Su clave de acceso al Portal de Alumnos ha sido enviada al E-Mail: " 
		largoNombre=InStr(session("RecuperaClaveMail"), "@")					  
		dominio=Mid(session("RecuperaClaveMail") ,largoNombre,len(session("RecuperaClaveMail")))			
		PNombre=Mid(session("RecuperaClaveMail") ,1,(largoNombre-1)/2)	
	
		for i=1 to (largoNombre-1)/2
			x=x &"*"
		next
		MensajeFinalCorreo= glosa & "" & PNombre & "" & x & "" & dominio
		session("MensajeLogin")=MensajeFinalCorreo
		Response.Redirect "alumnosfinning.asp" 
		Response.End()
		
	end if
end if
%>
<script language="JavaScript">
<!--
function setFocus() {
  var loginForm = document.getElementById("login");
  if (loginForm) {
	loginForm["logrut"].focus();
  }
}
function Ingresar()
{
	document.login.action="ValidaClavefinning.asp";	
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
  window.location ="alumnosfinning.asp?rc=s&rut="+texto;
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
<head>
<link rel="stylesheet" href="estilo.css" type="text/css" media="screen">
</head>
<body>
	<div id="contenedor_global">
		<div id="cabecera">
        <div id="titulo">Centro de Formación<br>
Técnica <span class="fi">Finning </span></div>
<div id="linea"><img src="images/linea.jpg" width="10" height="55"></div>
        <div id="bajada">Bienvenido al portal de alumnos en línea<br>
        <span style="font-size:10px;">A través de &eacute;ste, puedes realizar múltiples operaciones relacionadas con tu carrera.</span></div>
        
        </div>
	
	  <div id="zona_1">
			<div id="forms"><p>Login</p>
			<form name="login" id="login" method="post" >
			  <p><input id="logrut" name="logrut" type="text" value="" onChange="javascript:checkRutField(this.value);" style="width: 170px;	color:#fff;	background:#6e6c6b;	border:1px solid #6e6c6b;"></p>
              <p><input  id="logclave" name="logclave" type="password" value="" style="width: 170px;	color:#fff;	background:#6e6c6b;	border:1px solid #6e6c6b;"></p> 
              <input type="hidden" name="rut" value="">
              <input type="hidden" name="pin" value="">
              
              
          </form>
          </div>
          <div id="chek">
           <div id="recuperar"><a  id="recuperar" href="#"onClick="return RecuperaClave();" >recuperar contraseña</a></div>
        <div id="enviar"><a target="_top" id="enviar" href="#" onFocus="return ValidaDatos();"onClick="return ValidaDatos();" >Ingresar</a></div>
        </div>
          
        </div>
       
	  <div id="zona_info">
        <div id="foto"><img src="images/foto_alumnos.jpg" width="233" height="369"></div>
        <div id="texto">
        <p><ul>
          <li>Ingresa tu rut y contrase&ntilde;a para acceder al portal de Alumnos en l&iacute;nea.</li>
          <br>
    <li>Si es la primera vez   que ingresa a la aplicaci&oacute;n, usted debe ocupar como contrase&ntilde;a su rut   sin d&iacute;gito verficador (ej: 17635487) y luego cambiar la contrase&ntilde;a por una nueva.&nbsp;</li>
    <br>
    <li>Si olvidaste tu contraseña, haz clic en "recuperar contraseña" y te la reenviaremos a tu correo ingresado en el sistema.</li><br>
    </ul></p>
        </div>
        </div>
		<div id="pie">
			
		</div>
	</div>
</body>
</html>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">

function ValidaDatos() {

	rut = document.login.logrut.value;
	clave = document.login.logclave.value;

  var tmpstr = "";
  for ( i=0; rut.charAt(i) == '0' ; i++ );
  for (; i < rut.length ; i++ )
     tmpstr = tmpstr + rut.charAt(i);
  rut = tmpstr;
  document.login.logrut.value=rut;

	if ( rut.length == 0 || clave.length == 0  ) {
	    alert( "Ingrese los datos requeridos.");
	    cont_click = 0;
	    return false;
	} else {
		if (!(checkFields( rut, clave )))
		{
			cont_click=0;
			return false;
		}
		else
		{
			Ingresar();
			return true;	
		}	
	}

}


function checkFields( rut, clave )
{

  var tmpstr = "";
  var SoloRut="";
  var RutNum;
  window.document.login.pin.value =   window.document.login.logclave.value;

  for ( i=0; i < rut.length ; i++ )
    if ( rut.charAt(i) != ' ' && rut.charAt(i) != '.' && rut.charAt(i) != '-' )
      tmpstr = tmpstr + rut.charAt(i);
  rut = tmpstr;

  if ( !checkRutField(rut) )
    return false;

  if ( !checkDV( rut ) )
    return false;

  if ( !checkPinField(rut) )
      return false;

  window.document.login.rut.value = rut;

  //window.document.login.logclave.value = "";
  //window.document.login.logrut.value="";
  SoloRut=rut.substring(0,rut.length-1);

  RutNum=SoloRut
 document.login.rut.value = document.login.rut.value.toUpperCase();
 //document.passemp.rut.value = document.passemp.rut.value.toUpperCase();

  return true;
}

function checkCDV( dvr )

{
  dv = dvr + "";
  if ( dv != '0' && dv != '1' && dv != '2' && dv != '3' && dv != '4' && dv != '5' && dv != '6' && dv != '7' && dv != '8' && dv != '9' && dv != 'k'  && dv != 'K')
  {
    alert("El dígito verificador ingresado no es válido.");
    window.document.login.logrut.focus();
    window.document.login.logrut.select();
    return false;
  }
  return true;
}

function checkDV( crut )

{
  largo = crut.length;
  if ( largo < 2 )
  {
    alert("Por favor ingrese un RUT válido.");
    window.document.login.logrut.focus();
    window.document.login.logrut.select();
    return false;
  }

  if ( largo > 2 )
    rut = crut.substring(0, largo - 1);
  else
    rut = crut.charAt(0);
  dv = crut.charAt(largo-1);
  checkCDV( dv );

  if ( rut == null || dv == null )
      return 0;

  var dvr = '0';

  suma = 0;
  mul  = 2;

  for (i= rut.length -1 ; i >= 0; i--)
  {
    suma = suma + rut.charAt(i) * mul;
    if (mul == 7)
      mul = 2;
    else
      mul++;
  }


  res = suma % 11;
  if (res==1)
    dvr = 'k';
  else if (res==0)
    dvr = '0';
  else
  {
    dvi = 11-res;
    dvr = dvi + "";
  }

//alert(dvr);
//alert(dv.toLowerCase());

  if ( dvr != dv.toLowerCase() )
  {
    alert("El RUT ingresado es incorrecto.");
    window.document.login.logrut.focus();
    window.document.login.logrut.value = "";
    return false;
  }

      return true;
}

function checkPinField()
{

 if ( window.document.login.logclave.value.length < 3 )
  {
    alert("La clave debe poseer un largo mínimo de 4 digitos.");
    window.document.login.logclave.focus();
    window.document.login.logclave.select();
    return false;
  }
 if (ValidaCampo(window.document.login.logclave.value))
  return true;
 else
  return false;
}	

function checkRutField(texto)
{

  var tmpstr = "";
  for ( i=0; i < texto.length ; i++ )
    if ( texto.charAt(i) != ' ' && texto.charAt(i) != '.' && texto.charAt(i) != '-' )
      tmpstr = tmpstr + texto.charAt(i);
  texto = tmpstr;
  largo = texto.length;

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


  var invertido = "";

  for ( i=(largo-1),j=0; i>=0; i--,j++ )
    invertido = invertido + texto.charAt(i);


  var dtexto = "";

  dtexto = dtexto + invertido.charAt(0);
  dtexto = dtexto + '-';
  cnt = 0;

  for ( i=1,j=2; i<largo; i++,j++ )
  {
    if ( cnt == 3 )
    {
      dtexto = dtexto + '.';
      j++;
      dtexto = dtexto + invertido.charAt(i);
      cnt = 1;
    }
    else
    {
      dtexto = dtexto + invertido.charAt(i);
      cnt++;
    }
  }

  invertido = "";

  for ( i=(dtexto.length-1),j=0; i>=0; i--,j++ )
    invertido = invertido + dtexto.charAt(i);

  window.document.login.logrut.value = invertido;

  if ( checkDV(texto) )
  {
    return true;
   }	
	
  return false;
}
function ValidaCampo(campo){
  var enter = "\n"
  var caracteres = "abcdefghijklmnopqrstuvwxyzñ1234567890 ABCDEFGHIJKLMNOPQRSTUVWXYZÑáéíóúÁÉÍÓÚ" + String.fromCharCode(13) 

  var contador = 0
  for (var i=0; i < campo.length; i++) {
    ubicacion = campo.substring(i, i + 1)
    if (caracteres.indexOf(ubicacion) != -1) {
      contador++
    } else {
	  alert("Verifique los caracteres ingresados en su clave.")
      return false
    }
  }
   return true;
}
</script>