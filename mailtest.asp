<%
function CargarMenu()

%>

<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript"  SRC="include/validarut.js"></SCRIPT>
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
</script><style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
-->
</style>

<link href="css/tablas.css" rel="stylesheet" type="text/css">

<%
rc = Request.QueryString("rc")
rutRC = Request.QueryString("rut")

if rc = "s" then

strSql = "SELECT dbo.Decrypt(us_password) AS clave,pe_nombrecompleto AS nombre,COALESCE(pe_email,'') as mail FROM ca_usuarios ,tg_personas " &_
"WHERE ca_usuarios.id_persona = tg_personas.id_persona and " &_
"us_consuser = '" & rutRC & "' "

'response.Write(strSql)
'response.End()

Set rscl = Session("Conn").execute(strSql)


	 if rscl.eof then
	 	Response.Redirect("MensajeNoExiste.asp")
	 else
	 
	 if rscl("mail") = "" then
			Response.Redirect("MailError.asp")
		end if		


	dim Sql , Body
	Body =  chr(34) & "Estimado(a) " & rscl("nombre") & " su clave de acceso al portal de Alumnos es: " & rscl("clave") &"" & chr(34) 	
	
	Sql = "Exec Pr_RecuperaClave_prueba " & Body & ", " &  chr(34) &  rscl("mail") & chr(34)
	
	response.Write(Sql)

	response.write("ejecuta SP")
	'Session("Conn").Execute(Sql)
    response.write("EJECUTADO")
	response.End()	 
'		Set iMsg = CreateObject("CDO.Message")'
'		Set iConf = CreateObject("CDO.Configuration")
'		Set Flds = iConf.Fields
'		
'		' send one copy with Google SMTP server (with autentication)
'		schema = "http://schemas.microsoft.com/cdo/configuration/"
'		Flds.Item(schema & "sendusing") = 2
'		Flds.Item(schema & "smtpserver") = "mail.fullcom.cl" '"smtp.live.com" 
'		Flds.Item(schema & "smtpserverport") = 25
'		Flds.Item(schema & "smtpauthenticate") = 1
'		Flds.Item(schema & "sendusername") = "jcardenas@bettersoft.cl" '"recuperaclaveidma@hotmail.cl"
'		Flds.Item(schema & "sendpassword") =  "udpjcar" '"notedigo774"
'		Flds.Item(schema & "smtpusessl") = 0 '1
'		Flds.Update
'	
'		With iMsg
'		.To = rscl("mail")' "jcardenas@bettersoft.cl"
'		.From = "recuperacorreo@idma.cl" '"recuperaclaveidma@hotmail.cl"
'		.Subject = "Portal de Alumnos IDMA"
'		.HTMLBody = "Estimado " & rscl("nombre") & " su clave de acceso al portal de Alumnos es: '" & rscl("clave") &"' (omitir comillas)"
'		.Sender = "CFT IDMA"
'		.Organization = "CFT IDMA"
'		.ReplyTo = "recuperaclaveidma@hotmail.cl"
'		Set .Configuration = iConf
'		SendEmailGmail = .Send
'		End With
		
'		set iMsg = nothing
'		set iConf = nothing
'		set Flds = nothing
'		
		session("RecuperaClaveMail") = rscl("mail")
		response.Redirect("RecuperaClave.asp")
		
	end if
end if
%>
<table border="0" cellpadding="0" cellspacing="0" align="right">
  <tr>
    <td colspan="3" valign="top">
	<table width="198" border="0" cellspacing="0" cellpadding="0" height="559">
        <tr valign="top">
          <td width="18" height="449"><img src="imagenes/-_r4_c1.jpg" name="_r4_c1" width="18" height="276" border="0"></td>
          <td width="162" height="449"><table width="162" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><a href="#"  onMouseOut="MM_swapImgRestore()" onClick="return RecuperaClave();"  onMouseOver="MM_swapImage('_r9_c2','','imagenes/Botones/-_r9_c2_f2.jpg',1);" ><img name="_r9_c2" src="imagenes/botones/-_r9_c2.jpg" width="162" height="27" border="0"></a></td>
              </tr>
              <tr>
                <td background="imagenes/-_r10_c2.jpg" valign="bottom" height="139"><table border="0" cellspacing="0" cellpadding="0" width="130" align="center" height="79">
                      <form name="login" method="post" onSubmit="javascript:return ValidaDatos();">
                      <input type="hidden" name="rut" value="">
                      <input type="hidden" name="pin" value="">
					  <tr>
                        <td height="30" width="130"><div align="center">
                          <input name="logrut" type="text" class="casillas-form" id="logrut" size="13" onChange="javascript:checkRutField(this.value);" value=""  maxlength="12">
</div></td>
                      </tr>
                      <tr>
                        <td height="10" width="130"><div align="center"></div></td>
                      </tr>
                      <tr>
                        <td height="39" width="130"><div align="center">
                            <input name="logclave" type="password" value="" class="casillas-form" id="logclave" size="13" maxlength="9" >
                        </div></td>
                      </tr>
                    </form>
                   
                </table></td>
              </tr>              <tr>
                <td align="center"><a target="_top"  href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('_r11_c21','','Imagenes/botones/A-_r12_c2_f2.jpg',1)" onFocus="javascript:Ingresar();return ValidaDatos();" onClick="javascript:Ingresar();return ValidaDatos();" ><img src="Imagenes/botones/A-_r12_c2.jpg" name="_r11_c21" width="162" height="21" border="0" id="_r11_c21" onClick="javascript:Ingresar();return ValidaDatos();" ></a></td>
              </tr>
              
              <tr>
                <td valign="top"><img name="_r12_c2" src="imagenes/-_r12_c2.jpg" width="162" height="23" border="0"></td>
              </tr> 
              <!--<tr><td align="center">
                    <a href="#" id="needone" onClick="return RecuperaClave();">Recuperar clave</a></td>
                    </tr> -->

			<td valign="top"><img src="Imagenes/A-_r13_c2.jpg" width="162" height="149" ></td>
            
          </table>
          </td>
          <td height="449" width="18" valign="top"><img name="_r4_c3" src="imagenes/-_r4_c3.jpg" width="18" height="277" border="0"></td>
        </tr>
    </table></td>
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


Opcion1="INT_001"
Opcion2="INT_002"
Opcion3="INT_003"
Opcion4="470"
Opcion5="467"
Opcion6="INT_006"
Opcion7="INT_007"
Opcion8="471"
Opcion9="INT_009"
Opcion10="INT_010"
Opcion11="INT_011"

Nivel=0
GetPeriodoActivo
dim enlaces(12)



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
	   'permiso9=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion9,nivel,EstadoMatriculado,bloqueoA)
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
	   
	   
	   habilitada9=EstaHabilitada (Opcion9)
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
		'enlaces(3)="resultado.htm" 
		 enlaces(3)="resultado.asp" 
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
	if trim(habilitada8)="S"    then
	 
	   if Permiso8=1 then
		enlaces(9)="concent-notas.asp"
	   else
		enlaces(9)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(9)= "MensajeBloqueoHabilita.asp"
	end if
	
	if trim(habilitada9)="S"    then
	  if Permiso9=1 then
	  enlaces(10)="ed.htm"
	   else
	  enlaces(10)="MensajesBloqueos.asp" 
	  end if
	 else
		  enlaces(10)= "MensajeBloqueoHabilita.asp"
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
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript"  SRC="include/validarut.js"></SCRIPT>
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

   <table width="198" border="0" cellspacing="0" cellpadding="0" height="559" align="right">
     </tr>
    <td width="18" height="559" valign="top"><img name="_r4_c1" src="imagenes/-_r4_c1.jpg" width="18" height="385" border="0"></td>
    <td width="162" height="559" valign="top">
      <table width="162" border="0" cellspacing="0" cellpadding="0">
        </tr>
        <tr>
          <td><a target="_top" href="<%=enlaces(4)%>" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('A_r6_c2','','imagenes/botones/A-_r6_c2_f2.jpg',1);"><img name="A_r6_c2" src="imagenes/botones/A-_r6_c2.jpg" width="162" height="27" border="0" alt=""></td>
        </tr>
        <tr>
          <td><a target="_top" href="<%=enlaces(3)%>" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('A_r7_c2','','imagenes/botones/A-_r7_c2_f2.jpg',1);"><img name="A_r7_c2" src="imagenes/botones/A-_r7_c2.jpg" width="162" height="27" border="0" alt=""></td>
        </tr>
        <tr>
          <td><a target="_top" href="<%=enlaces(3)%>" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('A_r8_c2','','imagenes/botones/A-_r8_c2_f2.jpg',1);"><img name="A_r8_c2" src="imagenes/botones/A-_r8_c2.jpg" width="162" height="27" border="0" alt=""></td>
        </tr>
        <tr>
          <td><a target="_top" href="<%=enlaces(7)%>" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('A_r10_c2','','imagenes/botones/A-_r10_c2_f2.jpg',1);"><img name="A_r10_c2" src="imagenes/botones/A-_r10_c2.jpg" width="162" height="27" border="0" alt=""></td>
        </tr>
          <td background="imagenes/-_r10_c22.jpg" valign="bottom" height="139">
            <table border="0" cellspacing="0" cellpadding="0" width="142" align="center">
              <tr>
                <td width="130" height="20" align="center" >                  
                <div align="center" class="tex-totales-celda">Nombre del Alumno:</div></td>
              </tr>
              <tr>
                <td height="20" align="center" bgcolor="#DBECF2">
                <div align="center" class="text-normal-celdas">
                  <div align="center" class="text-normal-celdas"><%=Session("NomAlum")%></div>
                </div></td>
              </tr>
              <tr>
                <td width="130" height="20" align="center" class="text-normal-celdas"> <div align="center" class="tex-totales-celda">R.U.N.:</div></td>
              </tr>
              <tr>
                <td height="20" align="center" bgcolor="#DBECF2">
                <div align="center" class="text-normal-celdas">
                  <div align="center">
				  <%
				  'response.write("rutalu->"+Session("RutAlum")+"rx->"+session("RX") )
				  
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
			 		 End If  %>
					</div>
                </div></td>
              </tr>
              <tr>
					<td height="20" class="text-normal-celdas"><div align="center" class="tex-totales-celda">N&uacute;mero de Matr&iacute;cula:</div></td>
              </tr>
              <tr>
                <td height="20" align="center" bgcolor="#DBECF2">
                <div align="center" class="text-normal-celdas"><img src="imagenes/10x10.gif" width="10" height="10"><%=Session("CodCli")%></div></td>
              </tr>
              <tr>
                <td height="10"><div align="center"><img src="imagenes/10x10.gif" width="10" height="10"></div></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td width="205" align="center">
		  <a href="javascript:;Terminar()"><img src="imagenes/botones/A-cerrar-sesion-of.jpg"  id="Image1" onMouseOver="MM_swapImage('Image1','','imagenes/botones/A-cerrar-sesion-on.jpg',1)" onMouseOut="MM_swapImgRestore()" width="162" height="21" border="0"></a></div></td>
	    <td>		</tr>
        <tr>
          <td valign="top"><img name="_r12_c2" src="imagenes/-_r12_c2.jpg" width="162" height="23" border="0"></td>
		</tr>
		<td valign="top"><img src="Imagenes/A-_r13_c2.jpg" width="162" height="149" ></td>
	  </table>
      </td>
    <td height="559" width="18" valign="top"><img name="_r4_c3" src="imagenes/-_r4_c3.jpg" width="18" height="385" border="0"></td>
  </tr>
</table>

<%
End function
%>

