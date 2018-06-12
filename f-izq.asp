<!--#INCLUDE FILE="include/funciones.inc" -->
 
<%
function CargarMenu()
Session("FromNet") = 0

Session("RQRut")=""
Session("RQpass")=""
Session("RQCodcarr")=""
Session("Destino")=""

If Not Request.Form("user") Is Nothing And Request.Form.Count > 0 Then
    
	Session("RQRut") = desencripta_portal(Request.Form("user"))	
	Session("RQpass") = desencripta_portal(Request.Form("pass"))
	Session("RQCodcarr") = desencripta_portal(Request.Form("codcarr"))
	Session("Destino") = desencripta_portal(Request.Form("redirect")) 
	
	Session("FromNet") = 1
	response.Redirect("ValidaClave.asp")
end if

%>


 <%strParame="SELECT dbo.Fn_ValorParame('TIPOVALIDACIONRUT')Parame"
 
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
		TIPOVALIDACIONRUT=rstParame("Parame")
	else
		TIPOVALIDACIONRUT="" 
end if


session("TIPOVALIDACIONRUT")=TIPOVALIDACIONRUT 

if TIPOVALIDACIONRUT ="PERUANA" then 
	funcionJSR="javascript:validarRUC(this.value);"
else
	funcionJSR="javascript:checkRutField(this.value);"%>
    <script language="JavaScript" type="text/javascript" src="include/validarut.js"></script>
 <%end if%>  

<!--<script language="JavaScript" type="text/javascript" src="include/validarut.js"></script>-->

<script language="JavaScript" type="text/javascript">
<!--


function seteaFocus() {
	var obj = document.getElementById("ImageLogin");
	if (obj){
	   obj.click();   
	}
}

function setFocus() {
  var loginForm = document.getElementById("login");
  if (loginForm) {
	loginForm["logrut"].focus();
  }
}
function Ingresar()
{
	document.login.action="ValidaClave.asp";
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
  window.location ="Alumnos.asp?rc=s&rut="+texto;
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
	//alert('El Ruc ingresado es incorrecto.');
	//window.document.login.logrut.value ='';
    //return false;
	return true;
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
<link href="css/tablas.css" rel="stylesheet" type="text/css">
    <%
	

dim objMyDLL,liveStr
'set objMyDLL = server.createObject("LiveEdu.LiveEduClass")
rc = Request.QueryString("rc")
rutRC = Request.QueryString("rut")

if rc = "s" then

strSql = "SELECT dbo.Decrypt(us_password) AS clave,pe_nombrecompleto AS nombre,COALESCE(pe_email,'') as mail FROM ca_usuarios ,tg_personas " &_
"WHERE ca_usuarios.id_persona = tg_personas.id_persona and " &_
"us_consuser = '" & rutRC & "' "

Set rscl = Session("Conn").execute(strSql)

	 if rscl.eof then
	 	Response.Redirect("MensajeNoExiste.asp")
	 else
	 
	 if rscl("mail") = "" then
		Response.Redirect("MailError.asp")
	end if		
	
	strVM="SELECT dbo.fn_validamail('"& rscl("mail") &"')mail"
	set rstVM= Session("Conn").Execute(strVM)
		
	if not rstVM.eof then
		MalBueno=rstVM("mail")
		if MalBueno = 0 then
			Response.Redirect("MailError.asp")
		end if 
	end if 
	
	dim Sql , Body
	Body =  chr(34) & "Estimado(a) " & rscl("nombre") & " su clave de acceso al portal de Alumnos es: '" & rscl("clave") &"' (omitir comillas)" & chr(34) 	
	
	'Sql = "Exec Pr_RecuperaClave " & Body & ", " &  chr(34) &  rscl("mail") & chr(34) & ", " &  chr(34) &  Application("server") & chr(34) &", " &  chr(34) &  Application("port") & chr(34) & ", " &  chr(34) &  Application("ssl") & chr(34)& ", " &  chr(34) &  Application("from") & chr(34)& ", " &  chr(34) &  Application("user") & chr(34)& ", " &  chr(34) &  Application("pass") & chr(34)& ", " &  chr(34) &  Application("cred") & chr(34)	
	'Session("Conn").Execute(Sql)	
	
	StrSqlA ="INSERT INTO RA_AUDITA( USERNAME,CODCLI,OPCION,FECHA,TRANSACCION) " 
	StrSqlA =StrSqlA & "VALUES  (1,NULL,471,GETDATE(),'Recupera de clave Portal alumno con rut :"& rutRC &"')"
	'response.Write(StrSqlA)
	Session("Conn").Execute(StrSqlA)
	
	call enviargmail(rscl("mail"),Body)  	
	
	session("RecuperaClaveMail") = rscl("mail")
	
	dim strParame,pagina
	strParame="SELECT dbo.Fn_ValorParame('PERSONALIZAPA')Parame"
	set rstParame= Session("Conn").Execute(strParame)
	
	if not rstParame.eof then
		salirnuevo=rstParame("Parame")		
	else
		salirnuevo=""
	end if 
  	
	'if salirnuevo="SI" then
	''	pagina="alumnosfinning.asp"
	'else
	    strParame="SELECT dbo.Fn_ValorParame('PORTADA_PROPIA_PORTAL_ALUMNO')Parame"
	    set rstParame= Session("Conn").Execute(strParame)
	    if not rstParame.eof then
		        salirnuevo=rstParame("Parame")
	    else
		        salirnuevo=""
	    end if 
	    if(salirnuevo <> "" and not isNull(salirnuevo)) then
	            pagina=rstParame("Parame")
                Response.Write("<script language='javascript'>alert('Su clave de acceso al Portal de Alumnos ha sido enviada a su correo electr\u00f3nico.'); window.location.href='"  + pagina + "';</script>")
                response.End() 
	    else
	             response.Redirect("RecuperaClave.asp")
	    end if 
		
	
	
	
	
	
	
		
	end if
end if


    %>
    <table border="0" cellpadding="0" cellspacing="0" align="right">
        <tr>
            <td colspan="3" valign="top">
                <table width="198" border="0" cellspacing="0" cellpadding="0" height="559">
                    <tr valign="top">
                        <td width="18" height="449">
                            <img src="imagenes/-_r4_c1.jpg" name="_r4_c1" width="18" height="276" border="0">
                        </td>
                        <td width="162" height="449">
                            <table width="162" border="0" cellspacing="0" cellpadding="0">
                                <tr>
                                    <td>
                                        <a href="#" onMouseOut="MM_swapImgRestore()"  onMouseOver="MM_swapImage('_r9_c2','','imagenes/Botones/-_r9_c2_f2.jpg',1);"> 
                                            <input type="image" name="_r9_c2" onClick="return RecuperaClave();" id="_r9_c2" src="imagenes/botones/-_r9_c2.jpg" width="162" height="27" border="0"></a>
                                    </td>
                                </tr>
                                <tr>
                                    <td background="imagenes/-_r10_c2.jpg" valign="bottom" height="139">
                                        <table border="0" cellspacing="0" cellpadding="0" width="130" align="center" height="79">
                                            <form name="login" method="post"  onSubmit="javascript:return ValidaDatos();">
                                            <input type="hidden" name="rut" value="">
                                            <input type="hidden" name="pin" value="">
                                            <tr>
                                                <td height="30" width="130">
                                                    <div align="center">
                                                   
                                                        <!--<input name="logrut" type="text" class="casillas-form" id="logrut" size="13" onChange="javascript:checkRutField(this.value);"
                                                            value="" maxlength="12">-->
                                                        <input name="logrut" type="text" class="casillas-form" id="logrut" size="13" onChange="<%=funcionJSR%>" value="" maxlength="12">  
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
                                        <a target="_top" href="#"  onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('_r11_c21','','Imagenes/botones/A-_r12_c2_f2.jpg',1)">
                                            <input type="image" onClick="javascript:Ingresar();return ValidaDatos();"src="Imagenes/botones/A-_r12_c2.jpg" name="_r11_c21" width="162" height="21" border="0" id="_r11_c21" >
                                            </a>
                                    </td>
                                </tr>
                                <tr>
                                    <td valign="top">
                                        <img name="_r12_c2" src="imagenes/-_r12_c2.jpg" width="162" height="23" border="0">
                                    </td>
                                </tr>
                                <!--<tr><td align="center">
                    <a href="#" id="needone" onClick="return RecuperaClave();">Recuperar clave</a></td>
                    </tr> -->
                                <td valign="top">
                                    <img src="Imagenes/A-_r13_c2.jpg" width="162" height="149">
                                </td>
                            </table>
                        </td>
                        <td height="449" width="18" valign="top">
                            <img name="_r4_c3" src="imagenes/-_r4_c3.jpg" width="18" height="277" border="0">
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
    <img src="imagenes/spacer.gif" width="1" height="28" border="0">
    <%
end function
    %>
    <%
function CargarMenu2()
    %>
    <%
dim carrera,estado,bloqueoF,bloqueoB,Nivel,Permiso,MensajeP,bloqueoA
dim Opcion1,Opcion2,Opcion3,Opcion4,Opcion5,Opcion6,Opcion7,Opcion8,Opcion9,Opcion10
dim permiso1,permiso2,permiso3,permiso4,permiso5,permiso6,permiso7,permiso8,permiso9,permiso10
dim habilitada1,habilitada2,habilitada3,habilitada4,habilitada5,habilitada6,habilitada7,habilitada8,habilitada9,habilitada10
dim url1,url2,url3,url4,url5,url6,url7,url8,url9,url10,url11

'response.Write("Hola")
'response.End
'response.write(session("RutAlum"))

carrera=Session("codcarr")
estado=session("estado")
EstadoMatriculado=session("EstadoMatriculado")
bloqueoF=session("BloqueoF")
bloqueoB=session("BloqueoB")
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
strcc= "select coalesce(dbo.Fn_ValorParame('USACONCENTNOTASPA_CENFOTUR'),'')USACONCENTNOTASPA_CENFOTUR"
set rstcc= Session("Conn").Execute(strcc)
if not rstcc.eof then
	USACONCENTNOTASPA_CENFOTUR = rstcc("USACONCENTNOTASPA_CENFOTUR")
else
	USACONCENTNOTASPA_CENFOTUR = ""
end if

strParame="SELECT dbo.Fn_ValorParame('ACTIVALIBRESITE')Parame,dbo.Fn_ValorParame('SERVICIOLIBRESITE')SERVICIOLIBRESITE,dbo.Fn_ValorParame('URLLIBRESITE')URLLIBRESITE,dbo.Fn_ValorParame('ACTIVAELIBROS')ACTIVAELIBROS,dbo.Fn_ValorParame('ACTIVAREVISTASELECT')ACTIVAREVISTASELECT,dbo.Fn_ValorParame('ACTIVAPEARSON')ACTIVAPEARSON,dbo.Fn_ValorParame('ACTIVABOTONRCAPA')ACTIVABOTONRCAPA,dbo.Fn_ValorParame('URLBOTONRCAPA')URLBOTONRCAPA"
set rstParame= Session("Conn").Execute(strParame)

if not rstParame.eof then
	libresite = rstParame("Parame")
	SERVICIOLIBRESITE = rstParame("SERVICIOLIBRESITE")
	URLLIBRESITE = rstParame("URLLIBRESITE")
	ACTIVAELIBROS = rstParame("ACTIVAELIBROS")
	ACTIVAREVISTASELECT = rstParame("ACTIVAREVISTASELECT")
	ACTIVAPEARSON = rstParame("ACTIVAPEARSON") 
	ACTIVABOTONRCAPA = rstParame("ACTIVABOTONRCAPA") 
	URLBOTONRCAPA = rstParame("URLBOTONRCAPA") 
else
	libresite = ""
	SERVICIOLIBRESITE = ""
	URLLIBRESITE = ""
	ACTIVAELIBROS = ""
	ACTIVAREVISTASELECT =""
	ACTIVAPEARSON = ""
	ACTIVABOTONRCAPA =""
	URLBOTONRCAPA = ""
end if 

strParame2="SELECT dbo.Fn_ValorParame('BLOQUEAPAENCUESTAS')Parame"
set rstParame2= Session("Conn").Execute(strParame2)

if not rstParame2.eof then
	BLOQUEAPAENCUESTAS=rstParame2("Parame")
else
	BLOQUEAPAENCUESTAS="" 
end if

Opcion1="INT_001"
Opcion2="INT_002"
Opcion3="INT_003"
Opcion4="470"
Opcion5="467"
Opcion6="INT_006"
Opcion7="INT_007"
Opcion8="471"
Opcion9="477"
Opcion10="INT_010"
Opcion11="INT_011"

Nivel=0
GetPeriodoActivo
dim enlaces(12)


strParamePOP="SELECT dbo.Fn_ValorParame('ACTIVATRAUTO')ACTIVATRAUTO"
set rstParamePOP= Session("Conn").Execute(strParamePOP)

if not rstParamePOP.eof then
	ACTIVATRAUTO = rstParamePOP("ACTIVATRAUTO")
else
	ACTIVATRAUTO =""
end if 

StrA = "select Acepto from SIS_REG_INSCRIPCION WHERE codcli= '" & trim(Session("CodCli")) & "' and ano= '" & trim(Session("ano")) & "' AND periodo = '" & trim(Session("periodo")) & "'"
if bcl_ado(StrA, rstA) then
  if trim(valnulo(rstA("Acepto"),str_))="SI" then
  	AceptoCarga = "SI"
  else
  	AceptoCarga = "NO"
  end if   
else 
  AceptoCarga = "SI"
end if

strAN="select nuevo from mt_alumno where codcli='"& session("codcli") &"'"
if bcl_ado(strAN,rstAN) then
	if ucase(trim(valnulo(rstAN("nuevo"),str_)))="S" then
		ACTIVATRAUTO ="NO"
	end if 
end if

if session("logueado")="1" then
	   if Session("perfilNW") = 0 then
	   permiso1 = 0
	   else
	   permiso1 = 1
	   end if
	  
'	   permiso1=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion1,nivel,EstadoMatriculado,bloqueoA)
	   Nivel=1
	   'permiso2=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion2,nivel,EstadoMatriculado,bloqueoA)
	
	   'permiso3=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion3,nivel,EstadoMatriculado,bloqueoA)
	    permiso4=1'GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion4,nivel,EstadoMatriculado,bloqueoA)
	    permiso5=1'GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion5,nivel,EstadoMatriculado,bloqueoA)
	   'permiso6=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion6,nivel,EstadoMatriculado,bloqueoA)
	    'permiso7=1'GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion7,nivel,EstadoMatriculado,bloqueoA)
	    permiso8=1'GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion8,nivel,EstadoMatriculado,bloqueoA)
	    permiso9=1'GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion9,nivel,EstadoMatriculado,bloqueoA)
	   'permiso10=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion10,nivel,EstadoMatriculado,bloqueoA)
	   'permiso11=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion11,nivel,EstadoMatriculado,bloqueoA)
	   
	   habilitada1=EstaHabilitada (Opcion1)
	   habilitada2=EstaHabilitada (Opcion2)
	   habilitada3=EstaHabilitada (Opcion3)
	   
	   'habilitada4=EstaHabilitada (Opcion4)
	   url4=GetPermisoNW (Opcion4)
	   habilitada4=EstaHabilitadaNW (Opcion4)
	   if url4 = "N" then
	   permiso4 = 0
	   else
	   permiso4 = 1
	   end if
	   
	   
	   'habilitada5=EstaHabilitadaNW (Opcion5)
	   url5=GetPermisoNW (Opcion5)
	   habilitada5=EstaHabilitadaNW (Opcion5)
	   if url5 = "N" then
	   permiso5 = 0
	   else
	   permiso5 = 1
	   end if
	   'response.Write(url5)
	   'response.End()
	   
	   habilitada6=EstaHabilitada (Opcion6)
	   habilitada7=EstaHabilitada (Opcion7)
	   'habilitada8=EstaHabilitada (Opcion8)
	   	   
	   url8=GetPermisoNW (Opcion8)
	   habilitada8=EstaHabilitadaNW (Opcion8)
	   if url8 = "N" then
	   permiso8 = 0
	   else
	   permiso8 = 1
	   end if
	   
	   url9=GetPermisoNW (Opcion9)
	   habilitada9=EstaHabilitadaNW (Opcion9)
	   if url9 = "N" then
	   permiso9 = 0
	   else
	   permiso9 = 1
	   end if
	   
	    habilitada10=EstaHabilitada (Opcion10)
	   habilitada11=EstaHabilitada (Opcion11)

	  if Permiso1=1 then
		  'enlaces(0)="adm-acad.asp"
		  enlaces(0)="menu_tomaderamos.asp"
	  else
		  enlaces(0)="MensajesBloqueos.asp" 
	  end if
	
    if trim(Session("Logueado")) <> "" then
		if trim(habilitada2)="S" then
		  if Permiso2=1 then
			  enlaces(1)="inscrip-asigna.htm" 
			else
			  enlaces(1)="MensajesBloqueos.asp" 
		  end if
		else
			  enlaces(1)= "MensajeBloqueoHabilita.asp"
		end if
	else	
		  enlaces(1)= "MensajeBloqueoNoLog.asp"
	end if
	if trim(habilitada3)="S"  then
	  if Permiso3=1 then
		enlaces(2)="solicitud.htm"
	  else
		enlaces(2)="MensajesBloqueos.asp" 
	  end if
	else
		enlaces(2)= "MensajeBloqueoHabilita.asp"
	end if

	if trim(habilitada4)="S"  then
	  if Permiso4=1 then
	  	
		if ACTIVATRAUTO = "SI" and SESSION("PER_ID") <> "-1" then
			if AceptoCarga = "SI" then
				enlaces(3)="resultado.asp"
			else
				enlaces(3)="resultadotr.asp"		
			end if
		else
			enlaces(3)=url4	  		  
		end if
			
		'if ACTIVATRAUTO = "SI" then
		'	if AceptoCarga = "SI" then
		'		enlaces(3)="resultado.asp"
		'	else
		'		enlaces(3)="NoAceptoCarga.asp"		
		'	end if	
		'else
		'	enlaces(3)="resultado.asp" 	  		  
		'end if 	
			
		else
		enlaces(3)="MensajesBloqueos.asp" 
	  end if
	 else
		enlaces(3)= "MensajeBloqueoHabilita.asp"
	end if

	if trim(habilitada5)="S"  then
	  if Permiso5=1 then
		enlaces(4)=url5'"malla-curri.asp"
	  else
		enlaces(4)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(4)= "MensajeBloqueoHabilita.asp"
	end if

	if trim(habilitada6)="S"  then
	 if Permiso6=1 then
		enlaces(5)="ficha.asp"  
	  else
		enlaces(5)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(5)= "MensajeBloqueoHabilita.asp"
	end if

	if trim(habilitada7)="S"    then
	  if Permiso7=1 then
		enlaces(6)="certificado.asp"
		else
		enlaces(6)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(6)= "MensajeBloqueoHabilita.asp"
	end if

	if trim(habilitada8)="S"  then  
	  if Permiso8=1 then
		enlaces(7)=url8'"cambioclave.asp"
		else
		enlaces(7)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(7)= "MensajeBloqueoHabilita.asp"
	end if

  'TENGO DUDA EN ESTE****************************************************************
  	enlaces(8)="javascript:Terminar()"
	if trim(habilitada9)="S"    then
	 
	   if Permiso9=1 then
		enlaces(9)=url9'"concent-notas.asp"
	   else
		enlaces(9)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(9)= "MensajeBloqueoHabilita.asp"
	end if
	
	

	if trim(habilitada10)="S"    then
	  if Permiso10=1 then
	  enlaces(11)="Actualizador.asp"
	   else
	  enlaces(11)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(11)= "MensajeBloqueoHabilita.asp"
	end if
	
	if trim(habilitada11)="S"    then
	  if Permiso11=1 then
	  enlaces(12)="encuesta_ei.asp"
	   else
	  enlaces(12)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(12)= "MensajeBloqueoHabilita.asp"
	end if	
  
  '********************************************TENGO DUDA EN ESTO****************************
  
	if trim(habilitada2)="S"    then
		  if Permiso2=1 then
			  if InscritoNormal() then
				'enlaces(1)="resultado.htm" 
				enlaces(1)="resultado.asp" 
			  end if
		  else
				enlaces(1)="MensajesBloqueos.asp" 
		 end if		  
	 else
		  enlaces(1)= "MensajeBloqueoHabilita.asp"
	end if
				  
  	if trim(habilitada3)="S"    then
		  if Permiso3=1 then
			  if InscritoEspecial() then
				'enlaces(2)="resultado.htm"
				enlaces(2)="resultado.asp" 
			  end if
		   else
  				enlaces(2)="MensajesBloqueos.asp" 
		 end if
	 else
		  enlaces(2)= "MensajeBloqueoHabilita.asp"
	end if		  
	
		 if SESSION("PER_ID") = "-1" then
			'enlaces(1)="sinperiodo.htm" 
			enlaces(1)="sinperiodo.asp" 
			'enlaces(2)="sinperiodo.htm" 
			enlaces(2)="sinperiodo.asp" 
		 end if
else

	if session("SW")=1 then	
	  	  'enlaces(x)="adm-log.htm"
		  enlaces(0)="alumnos.asp"
	  	  enlaces(1)="alumnos.asp"
	  	  enlaces(2)="alumnos.asp"
	  	  enlaces(3)="alumnos.asp"
	  	  enlaces(4)="alumnos.asp"
	  	  enlaces(5)="alumnos.asp"
	  	  enlaces(6)="alumnos.asp"
	  	  enlaces(7)="alumnos.asp"
	  	  enlaces(8)="alumnos.asp"
	  	  enlaces(9)="alumnos.asp"
	  	  enlaces(10)="alumnos.asp"
	  	  enlaces(11)="alumnos.asp"
		  enlaces(12)="alumnos.asp"
	else
	  enlaces(0)="MensajeBloqueoNoLog.asp"
	  enlaces(1)="MensajeBloqueoNoLog.asp"
	  enlaces(2)="MensajeBloqueoNoLog.asp"
	  enlaces(3)="MensajeBloqueoNoLog.asp"
	  enlaces(4)="MensajeBloqueoNoLog.asp"
	  enlaces(5)="MensajeBloqueoNoLog.asp"  
	  enlaces(6)="MensajeBloqueoNoLog.asp"
	  enlaces(7)="MensajeBloqueoNoLog.asp"
	  'enlaces(8)="adm-log.htm"
	  enlaces(8)="alumnos.asp"
	  enlaces(9)="MensajeBloqueoNoLog.asp"
	  enlaces(10)="MensajeBloqueoNoLog.asp"
	  enlaces(11)="MensajeBloqueoNoLog.asp"
	  enlaces(12)="MensajeBloqueoNoLog.asp"
	  
	 end if
end if
'response.Write("esto"+enlaces(4))
'response.End()


    %>
    <%
Session("FechaIng")=time()
    %>
<script language="JavaScript" type="text/javascript" src="include/validarut.js"></script> 
    <script language="JavaScript">
<!--

<!--
function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
 var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
   var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
   if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
//-->

function Terminar()
{
   if (confirm(String.fromCharCode(191)+"Desea abandonar el Portal de Alumnos?") )
   {
      window.location = "salir.asp";
        //window.top.location = "alumn-udd.asp";
   }
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

    <script language="javascript" type="text/javascript">
        /*
        * A JavaScript implementation of the Secure Hash Algorithm, SHA-1, as defined
        * in FIPS PUB 180-1
        * Version 2.1a Copyright Paul Johnston 2000 - 2002.
        * Other contributors: Greg Holt, Andrew Kepert, Ydnar, Lostinet
        * Distributed under the BSD License
        * See http://pajhome.org.uk/crypt/md5 for details.
        */

        /*
        * Configurable variables. You may need to tweak these to be compatible with
        * the server-side, but the defaults work in most cases.
        */
        var hexcase = 0;  /* hex output format. 0 - lowercase; 1 - uppercase        */
        var b64pad = "="; /* base-64 pad character. "=" for strict RFC compliance   */
        var chrsz = 8;  /* bits per input character. 8 - ASCII; 16 - Unicode      */

        /*
        * These are the functions you'll usually want to call
        * They take string arguments and return either hex or base-64 encoded strings
        */
        function hex_sha1(s) { return binb2hex(core_sha1(str2binb(s), s.length * chrsz)); }
        function b64_sha1(s) { return binb2b64(core_sha1(str2binb(s), s.length * chrsz)); }
        function str_sha1(s) { return binb2str(core_sha1(str2binb(s), s.length * chrsz)); }
        function hex_hmac_sha1(key, data) { return binb2hex(core_hmac_sha1(key, data)); }
        function b64_hmac_sha1(key, data) { return binb2b64(core_hmac_sha1(key, data)); }
        function str_hmac_sha1(key, data) { return binb2str(core_hmac_sha1(key, data)); }

        /*
        * Perform a simple self-test to see if the VM is working
        */
        function sha1_vm_test() {
            return hex_sha1("abc") == "a9993e364706816aba3e25717850c26c9cd0d89d";
        }

        /*
        * Calculate the SHA-1 of an array of big-endian words, and a bit length
        */
        function core_sha1(x, len) {
            /* append padding */
            x[len >> 5] |= 0x80 << (24 - len % 32);
            x[((len + 64 >> 9) << 4) + 15] = len;

            var w = Array(80);
            var a = 1732584193;
            var b = -271733879;
            var c = -1732584194;
            var d = 271733878;
            var e = -1009589776;

            for (var i = 0; i < x.length; i += 16) {
                var olda = a;
                var oldb = b;
                var oldc = c;
                var oldd = d;
                var olde = e;

                for (var j = 0; j < 80; j++) {
                    if (j < 16) w[j] = x[i + j];
                    else w[j] = rol(w[j - 3] ^ w[j - 8] ^ w[j - 14] ^ w[j - 16], 1);
                    var t = safe_add(safe_add(rol(a, 5), sha1_ft(j, b, c, d)),
                       safe_add(safe_add(e, w[j]), sha1_kt(j)));
                    e = d;
                    d = c;
                    c = rol(b, 30);
                    b = a;
                    a = t;
                }

                a = safe_add(a, olda);
                b = safe_add(b, oldb);
                c = safe_add(c, oldc);
                d = safe_add(d, oldd);
                e = safe_add(e, olde);
            }
            return Array(a, b, c, d, e);

        }

        /*
        * Perform the appropriate triplet combination function for the current
        * iteration
        */
        function sha1_ft(t, b, c, d) {
            if (t < 20) return (b & c) | ((~b) & d);
            if (t < 40) return b ^ c ^ d;
            if (t < 60) return (b & c) | (b & d) | (c & d);
            return b ^ c ^ d;
        }

        /*
        * Determine the appropriate additive constant for the current iteration
        */
        function sha1_kt(t) {
            return (t < 20) ? 1518500249 : (t < 40) ? 1859775393 :
         (t < 60) ? -1894007588 : -899497514;
        }

        /*
        * Calculate the HMAC-SHA1 of a key and some data
        */
        function core_hmac_sha1(key, data) {
            var bkey = str2binb(key);
            if (bkey.length > 16) bkey = core_sha1(bkey, key.length * chrsz);

            var ipad = Array(16), opad = Array(16);
            for (var i = 0; i < 16; i++) {
                ipad[i] = bkey[i] ^ 0x36363636;
                opad[i] = bkey[i] ^ 0x5C5C5C5C;
            }

            var hash = core_sha1(ipad.concat(str2binb(data)), 512 + data.length * chrsz);
            return core_sha1(opad.concat(hash), 512 + 160);
        }

        /*
        * Add integers, wrapping at 2^32. This uses 16-bit operations internally
        * to work around bugs in some JS interpreters.
        */
        function safe_add(x, y) {
            var lsw = (x & 0xFFFF) + (y & 0xFFFF);
            var msw = (x >> 16) + (y >> 16) + (lsw >> 16);
            return (msw << 16) | (lsw & 0xFFFF);
        }

        /*
        * Bitwise rotate a 32-bit number to the left.
        */
        function rol(num, cnt) {
            return (num << cnt) | (num >>> (32 - cnt));
        }

        /*
        * Convert an 8-bit or 16-bit string to an array of big-endian words
        * In 8-bit function, characters >255 have their hi-byte silently ignored.
        */
        function str2binb(str) {
            var bin = Array();
            var mask = (1 << chrsz) - 1;
            for (var i = 0; i < str.length * chrsz; i += chrsz)
                bin[i >> 5] |= (str.charCodeAt(i / chrsz) & mask) << (32 - chrsz - i % 32);
            return bin;
        }

        /*
        * Convert an array of big-endian words to a string
        */
        function binb2str(bin) {
            var str = "";
            var mask = (1 << chrsz) - 1;
            for (var i = 0; i < bin.length * 32; i += chrsz)
                str += String.fromCharCode((bin[i >> 5] >>> (32 - chrsz - i % 32)) & mask);
            return str;
        }

        /*
        * Convert an array of big-endian words to a hex string.
        */
        function binb2hex(binarray) {
            var hex_tab = hexcase ? "0123456789ABCDEF" : "0123456789abcdef";
            var str = "";
            for (var i = 0; i < binarray.length * 4; i++) {
                str += hex_tab.charAt((binarray[i >> 2] >> ((3 - i % 4) * 8 + 4)) & 0xF) +
           hex_tab.charAt((binarray[i >> 2] >> ((3 - i % 4) * 8)) & 0xF);
            }
            return str;
        }

        /*
        * Convert an array of big-endian words to a base-64 string
        */
        function binb2b64(binarray) {
            var tab = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwx  yz0123456789+/";
            var str = "";
            for (var i = 0; i < binarray.length * 4; i += 3) {
                var triplet = (((binarray[i >> 2] >> 8 * (3 - i % 4)) & 0xFF) << 16)
                | (((binarray[i + 1 >> 2] >> 8 * (3 - (i + 1) % 4)) & 0xFF) << 8)
                | ((binarray[i + 2 >> 2] >> 8 * (3 - (i + 2) % 4)) & 0xFF);
                for (var j = 0; j < 4; j++) {
                    if (i * 8 + j * 6 > binarray.length * 32) str += b64pad;
                    else str += tab.charAt((triplet >> 6 * (3 - j)) & 0x3F);
                }
            }
            return str;
        }

    function enviarlibresitio()
    {
        var fecha, matricula, servicio, hash, nombreusuario, tipodeusuario, carrera, url, final;

        //matricula = "librisite@librisite.com";
        matricula = "<%=Session("NomAlum")%>";//username

         var actual = new Date();


         fecha = Date.UTC(actual.getFullYear(), actual.getMonth(), actual.getDay(), actual.getHours(), actual.getMinutes(), actual.getSeconds())
         //alert(fecha);
    servicio = "<%=SERVICIOLIBRESITE%>";
    hash = hex_sha1(matricula + "|" + servicio + "|" + fecha);
    //document.writeln(hash);

    nombreusuario = "<%=Session("RutAlum")%>";
    tipodeusuario = "Estudiante";
    carrera = "<%=Session("codcarr")%>";
    url = "<%=URLLIBRESITE%>";
 
    //echo "<a href=".$url."usuario_valida.php?matricula=".$matricula."&fecha=".$fecha."&carrera=".$carrera."&tipodeusuario=".$tipodeusuario."&nombreusuario=".$nombreusuario."&libroId=1&hash=".$hash.">link</a>";

    final = url + "usuario_valida.php?matricula=" + matricula + "&fecha=" + fecha + "&carrera=" + carrera + "&tipodeusuario=" + tipodeusuario + "&nombreusuario=" + nombreusuario + "&libroId=&hash=" + hash;
    //alert(final);

        //window.open(url+"usuario_valida.php?matricula="+matricula+"&fecha="+fecha+"&carrera="+carrera+"&tipodeusuario="+tipodeusuario+"&nombreusuario="+nombreusuario+"&libroId=1&hash="+hash);

        // window.open("http://ccs.libricentro.com/usuario_valida.php?matricula=hcastillo@librisite.com&fecha=1310488508&carrera=Ing.Telecom&tipodeusuario=Alumno&nombreusuario=hcastillo&libroId=&hash=368c92f423e99aa5ab6f6cde4222ea6c529ae9cb");

    window.open(final); 
    }
    </script>

    <table width="198" border="0" cellspacing="0" cellpadding="0" height="559" align="right">
        </tr>
        <%if libresite="SI" then%>
        <td width="18" height="559" valign="top">
            <img name="_r4_c1" src="imagenes/-_r4_c1.jpg" width="18" height="430" border="0">
        </td>
        <%else%>
        <td width="18" height="559" valign="top">
            <img name="_r4_c1" src="imagenes/-_r4_c1.jpg" width="18" height="385" border="0">
        </td>
        <%end if%>
        <td width="162" height="559" valign="top">
            <table width="162" border="0" cellspacing="0" cellpadding="0">
                </tr>
                <tr>
                    <td>
                        <a target="_top" href="javascript:Enviar3('<%=enlaces(4)%>','<%=TomadeRamo(session("Codcarr"),session("Codcli"),session("ano_ed"),session("periodo_ed"))%>','<%=BLOQUEAPAENCUESTAS%>');"
                            onmouseout="MM_swapImgRestore();" onmouseover="MM_swapImage('A_r6_c2','','imagenes/botones/A-_r6_c2_f2.jpg',1);">
                            <img name="A_r6_c2" src="imagenes/botones/A-_r6_c2.jpg" width="162" height="27" border="0"
                                alt="">
                    </td>
                </tr>
                <tr>
                    <td>
                        <a target="_top" href="javascript:Enviar3('<%=enlaces(3)%>','<%=TomadeRamo(session("Codcarr"),session("Codcli"),session("ano_ed"),session("periodo_ed"))%>','<%=BLOQUEAPAENCUESTAS%>');"
                            onmouseout="MM_swapImgRestore();" onmouseover="MM_swapImage('A_r7_c2','','imagenes/botones/A-_r7_c2_f2.jpg',1);">
                            <img name="A_r7_c2" src="imagenes/botones/A-_r7_c2.jpg" width="162" height="27" border="0"
                                alt="">
                    </td>
                </tr>
                <tr>
                    <td>
                    	
                        <%if USACONCENTNOTASPA_CENFOTUR = "SI" then%>
                        <a target="_top" href="javascript:Enviar3('<%=enlaces(9)%>','<%=TomadeRamo(session("Codcarr"),session("Codcli"),session("ano_ed"),session("periodo_ed"))%>','<%=BLOQUEAPAENCUESTAS%>');"onmouseout="MM_swapImgRestore();" onmouseover="MM_swapImage('A_r8_c2','','imagenes/botones/cdnpp2.jpg',1);">
                            <img name="A_r8_c2" src="imagenes/botones/cdnpp1.jpg" width="162" height="27" border="0" alt="">
                        <%else%>
                        <a target="_top" href="javascript:Enviar3('<%=enlaces(9)%>','<%=TomadeRamo(session("Codcarr"),session("Codcli"),session("ano_ed"),session("periodo_ed"))%>','<%=BLOQUEAPAENCUESTAS%>');"onmouseout="MM_swapImgRestore();" onmouseover="MM_swapImage('A_r8_c2','','imagenes/botones/cdn2.jpg',1);">
                            <img name="A_r8_c2" src="imagenes/botones/cdn1.jpg" width="162" height="27" border="0" alt="">
                        <%end if%>
                  
                    </td>
                </tr>
                <tr>
                    <td>
                    
                    	
                        <a target="_top" href="javascript:Enviar3('<%=enlaces(7)%>','<%=TomadeRamo(session("Codcarr"),session("Codcli"),session("ano_ed"),session("periodo_ed"))%>','<%=BLOQUEAPAENCUESTAS%>');"
                            onmouseout="MM_swapImgRestore();" onmouseover="MM_swapImage('A_r10_c2','','imagenes/botones/A-_r10_c2_f2.jpg',1);">
                            <img name="A_r10_c2" src="imagenes/botones/A-_r10_c2.jpg" width="162" height="27"
                                border="0" alt="">
                    </td>
                </tr>
                <td background="imagenes/-_r10_c22.jpg" valign="bottom" height="139">
                    <table border="0" cellspacing="0" cellpadding="0" width="142" align="center">
                        <tr>
                            <td width="130" height="20" align="center">
                                <div align="center" class="tex-totales-celda">
                                    <label id="lblNombreAlum">Nombre del Alumno:</label>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td height="20" align="center" bgcolor="#DBECF2">
                                <div align="center" class="text-normal-celdas">
                                    <div align="center" class="text-normal-celdas">
                                        <%=Session("NomAlum")%></div>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td width="130" height="20" align="center" class="text-normal-celdas">
                                <div align="center" class="tex-totales-celda">
                                    <label id="lblRun">R.U.N.:</label>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td height="20" align="center" bgcolor="#DBECF2">
                                <div align="center" class="text-normal-celdas">
                                    <div align="center">
                                        <%
				  'response.write("rutalu->"+Session("RutAlum")+"rx->"+session("RX") )
				  
				  	IF trim(ucase(session("TIPOVALIDACIONRUT"))) ="PERUANA" then
						response.write(session("Rut"))
					else
						If Session("RutAlum")= empty Then 
							 If response.write(session("RX"))<> Empty Then
							  'response.write("rx vacio")
								response.write(session("RX"))
								'response.Write("17.763.235-k")
							else
							   response.write(session("RutAlum")) 'response.Write("17.763.235-k")
							End if
						 else 
						 response.write(session("RutAlum"))
							'response.write(session("RutAlum"))
							'response.Write("17.763.235-k")
						 End If  
					End If%>
                                    </div>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td height="20" class="text-normal-celdas">
                                <div align="center" class="tex-totales-celda">
                                    N&uacute;mero de Matr&iacute;cula:</div>
                            </td>
                        </tr>
                        <tr>
                            <td height="20" align="center" bgcolor="#DBECF2">
                                <div align="center" class="text-normal-celdas">
                                    <img src="imagenes/10x10.gif" width="10" height="10"><%=Session("CodCli")%></div>
                            </td>
                        </tr>
                        <tr>
                            <td height="10">
                                <div align="center">
                                    <img src="imagenes/10x10.gif" width="10" height="10"></div>
                            </td>
                        </tr>
                    </table>
                </td>
                </tr>
                <!--Mejora botonera dinamica-->
                <%rstBotonera ="sp_traeFwlinkPortales 'MENU_IZQ','PORTAL ALUMNO'"
				Set rstBotonera = session("conn").Execute(rstBotonera)
				While Not rstBotonera.Eof
                %>
                </tr>
                    <td align="center">
                        <a href="<%=rstBotonera(2)%>" target="<%=rstBotonera(5)%>">
                            <img alt="" id="<%=rstBotonera(0)%>" src="<%=rstBotonera(3)%>" onMouseOver="MM_swapImage('<%=rstBotonera(0)%>','','<%=rstBotonera(1)%>',1)"
                            onmouseout="MM_swapImgRestore()" width="162" border="0" />
                        </a>
                    </td>
                </tr>
                <tr>
                    <td height="4">
                        <div align="center">
                        	<img src="imagenes/10x10.gif" width="10" height="4">
                        </div>
                    </td>
                </tr>
                <%
				rstBotonera.MoveNext
				Wend%>
                <!--Fin Mejora botonera dinamica-->
                
                
                    <tr>
                        <td width="205" align="center">
                            <a href="javascript:;Terminar()">
                                <img src="imagenes/botones/A-cerrar-sesion-of.jpg" id="Image1N" onMouseOver="MM_swapImage('Image1N','','imagenes/botones/A-cerrar-sesion-on.jpg',1)"
                                    onmouseout="MM_swapImgRestore()" width="162" height="21" border="0"></a></div>
                        </td>
                        <td>
                    </tr>
                    <tr>
                        <td valign="top">
                            <img name="_r12_c2" src="imagenes/-_r12_c2.jpg" width="162" height="23" border="0">
                        </td>
                    </tr>
                    <td valign="top">
                        <img src="Imagenes/A-_r13_c2.jpg" width="162" height="149">
                    </td>
            </table>
        </td>
        <%if libresite="SI" then%>
        <td height="559" width="18" valign="top">
            <img name="_r4_c3" src="imagenes/-_r4_c3.jpg" width="18" height="430" border="0">
        </td>
        <%else%>
        <td height="559" width="18" valign="top">
            <img name="_r4_c3" src="imagenes/-_r4_c3.jpg" width="18" height="385" border="0">
        </td>
        <%end if%>
        </tr>
    </table>
    <%ObjetosLocalizacion("f-izq.asp")%>
    <%
End function

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
