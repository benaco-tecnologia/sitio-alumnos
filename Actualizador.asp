
<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
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
<%
Dim str,Sql,sTRSQL,Sl,mySTR,sT,Strs,MStr
dim rst,Rql,rSTSQL,Rs,myRST,rT,rsts,MRst,rTQ,sqlZ,rqlZ,sqlTP
dim rut,dv,paterno,materno,nombre,estado,fecnac,diractual,comuna,ciudad,fonoact,dirproc,ciuproc,comunaproc,fonoproc,codestcivil
dim mail,Opiniones,Hora,Fecha,Carrera,MailInstitucion,valida
'set rst=Server.CreateObject ("ADODB.Recordset")

if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if
	if EstaHabilitadaNW (478)="S" then 
		if GetPermisoNW(478) ="N" then
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
					session("MensajeBloqueosVarios") ="Para acceder a esta opciÛn debe responder sus Encuestas Pendientes."
					response.Redirect("MensajeBloqueo.asp")	
				end if 
			end if 
			
		end if
	else
		response.Redirect("MensajeBloqueoHabilita.asp")
	end if
	
	strParame="SELECT dbo.Fn_ValorParame('MSJ_SOLICITA_CLAVE')Parame"
			set rstParame= Session("Conn").Execute(strParame)		
			if not rstParame.eof then
					MSJ_SOLICITA_CLAVE=rstParame("Parame")
				else
					MSJ_SOLICITA_CLAVE="" 
			end if
	
Audita 478,"Ingresa a Atualiza datos"

dim CiudadActual
valida = "N"
fecha=Date
hora=Time
Carrera=session("Codcarr")
VALOR=request.Form("VALOR")
if VALOR <> "GRABAR" then
	call llamar
else
	call grabar
	call Correo
	call llamar
end if

function Grabar ()
	strParame2="SELECT dbo.Fn_ValorParame('GRABAPALUMDATOSMAYUSCULA')Parame"
	set rstParame2= Session("Conn").Execute(strParame2)		
	if not rstParame2.eof then
			GRABAPALUMDATOSMAYUSCULA=rstParame2("Parame")
	end if 
	
	if GRABAPALUMDATOSMAYUSCULA ="SI" then 
		rut=LimpiarTexto(request.Form("txtrut"))
		paterno=UCASE(LimpiarTexto(request.Form("txtpaterno")))
		materno=UCASE(LimpiarTexto(request.Form("txtmaterno")))
		nombre=UCASE(LimpiarTexto(request.Form("txtnombres")))
		fecnac=LimpiarTextoMail(replace(request.Form("txtfecha"),"/","-"))
		estado=UCASE(LimpiarTexto(request.Form("cbestado")))
		diractual=UCASE(LimpiarTexto(request.Form("txtdireccionact")))
		comuna=UCASE(Request.Form("txtcomunaact"))
		ciudad=UCASE(LimpiarTexto(request.Form("txtciudadact")))
		fonoact=LimpiarTexto(request.Form("txtfono"))
		dirproc=UCASE(LimpiarTexto(request.Form("txtdireccionpro")))
		ciuproc=UCASE(LimpiarTexto(request.Form("txtciudadpro")))
		comunaproc=UCASE(request.Form("txtcomunapro"))
		fonoproc=LimpiarTexto(request.Form("txtfonoproc"))
		mail=UCASE(LimpiarTextoMail(request.Form("txtmail")))        
		'MailInstitucion=UCASE(LimpiarTexto(request.Form("txtmailinstitucion")))
	
	else
		rut=LimpiarTexto(request.Form("txtrut"))
		paterno=LimpiarTexto(request.Form("txtpaterno"))
		materno=LimpiarTexto(request.Form("txtmaterno"))
		nombre=LimpiarTexto(request.Form("txtnombres"))
		fecnac=LimpiarTextoMail(replace(request.Form("txtfecha"),"/","-"))
		estado=LimpiarTexto(request.Form("cbestado"))
		diractual=LimpiarTexto(request.Form("txtdireccionact"))
		comuna=Request.Form("txtcomunaact")
		ciudad=LimpiarTexto(request.Form("txtciudadact"))
		fonoact=LimpiarTexto(request.Form("txtfono"))
		dirproc=LimpiarTexto(request.Form("txtdireccionpro"))
		ciuproc=LimpiarTexto(request.Form("txtciudadpro"))
		comunaproc=request.Form("txtcomunapro")
		fonoproc=LimpiarTexto(request.Form("txtfonoproc"))
		mail=LimpiarTextoMail(request.Form("txtmail"))    
		'MailInstitucion=LimpiarTexto(request.Form("txtmailinstitucion"))
	end if
	
	
	
	
	mySTR="select * from ra_mantieneactWEB with (nolock) "
	set myRST=Session("Conn").execute(mySTR)
	Sql="Update mt_client set "
	if  myRST("paterno")="S" then
		Sql=Sql & " paterno='"& paterno & "',"
		valida = "S"
	end if
	if  myRST("materno")="S" then
		Sql=Sql & " materno='"& materno & "',"
		valida = "S"
	end if
	if  myRST("nombres")="S" then
		Sql=Sql & " nombre='"& nombre & "',"
		valida = "S"
	end if
	if  myRST("fecha_nacimiento")="S" then
		Sql=Sql & " fecnac='"& trim(fecnac) &"',"
		valida = "S"
	end if
	if  myRST("direccion")="S" then
		Sql=Sql & "diractual='"& trim(diractual) &"',"
		valida = "S"
	end if
	if  myRST("comuna")="S" then
		Sql=Sql & " comuna='"& trim(comuna) &"',"
		valida = "S"
	end if
	if  myRST("ciudad")="S" then
		Sql=Sql & "ciudadact='"& trim(ciudad) &"', "
		valida = "S"
	end if
	if  myRST("estado_civil")="S" then
		Sql=Sql & "CODESTCIVIL='"& trim(estado) &"', "
		valida = "S"
	end if
	if  myRST("fono")="S" then
		Sql=Sql & " fonoact ='"& trim(fonoact) &"',"
		valida = "S"
	end if
	if  myRST("direccion_procedencia")="S" then
		Sql=Sql & " dirproc='"& trim(dirproc) &"',"
		valida = "S"
	end if
	if  myRST("ciudad_procedencia")="S" then
		Sql=Sql & " ciuproc='"& trim(ciuproc) &"',"
		valida = "S"
	end if
	if  myRST("comuna_procedencia")="S" then
		Sql=Sql & " comunapro='"& trim(comunaproc) &"', "
		valida = "S"
	end if
	if  myRST("fono_procedencia")="S" then
		Sql=Sql & " fonoproc='"& trim(fonoproc) &"',"
		valida = "S"
	end if
	if  myRST("e_mail")="S" then
		Sql=Sql & " mail='"& trim(replace(mail,",",".")) &"'"
		valida = "S"
	else
		Sql=Sql & " mail=mail "'"
	end if
	if(valida = "S") then
	
	Sql=Sql &  " where codcli='"& trim(session("rut")) &"'"
	
	set Rql=Session("Conn").execute(Sql)    
	
	Strs="INSERT INTO RA_AUDITA (USERNAME,CODCLI,OPCION,FECHA,TRANSACCION) VALUES "
	Strs= Strs & " ('"& Session("id_usuario") &"','"& session("codcli")&"',478,GETDATE(),' Portal Alumnos - Actualiza Datos de Alumno') "

	if  myRST("e_mail")="S" then
	sqlTP = "update tg_personas set pe_email ='"& trim(replace(mail,",",".")) &"' where pe_rut='"& trim(Session("RutAlum")) &"'"
	Session("Conn").execute(sqlTP)
	end if
	
		
	
	'Strs="insert into rA_audita insert into SIS_REG_ACTUALIZADATOS (RUT) "
	'Strs=Strs & " values ('"& trim(session("rut")) &"') "

	'response.write(Strs)
	'response.End()
	
	
	set rsts=Session("Conn").execute(Strs)
	end if
	'Se debiÚ acrualizar el registro.
	'sTRSQL="Select * from RA_OPINIONES WHERE rut='"& rut &"' and fecha = '"& fecha &"'"
	'set rSTSQL=Session("Conn").execute(sTRSQL)
	 'if rSTSQL.eof then 
	 '	 sl="Insert into RA_OPINIONES (rut,fecha,hora,codcarr,opinion) "
	 '	 sl=sl & " values ('"& rut &"','"& fecha &"','"& Hora &"','"& Carrera &"','"& Opiniones &"') "
	 '	 set rs=Session("Conn").execute(sl)
	 'else
	 ' 	sl="Update RA_OPINIONES set hora='"& Hora &"',opinion='"& Opiniones &"' "
     '	sl=sl & " where rut='"& rut &"' and fecha='"& fecha &"' and codcarr='"& Carrera &"'"
     '		 set rs=Session("Conn").execute(sl)
	 ' end if
	 '**************** Se terminÚ la actualizacion ******************
VALOR=""

response.Redirect("menu_perfil.asp")
end Function
function Llamar ()
	str="select  distinct b.codcli, b.dig, b.paterno,b.materno,b.nombre,c.DESCRIPCION,b.CODESTCIVIL,b.fecnac, "
	str=str & " b.diractual,b.comuna,b.ciudadact,b.fonoact,b.dirproc,b.ciuproc,b.comunapro,b.fonoproc, "
	str=str & " replace(b.mail,',','.') as mail,b.mail_inst "
	str=str & " from mt_alumno a with (nolock), mt_client  b with (nolock) LEFT JOIN MT_ESTCIVIL c on b.CODESTCIVIL = c.CODIGO"
	str=str & " where a.rut=b.codcli and "
	str=str & " a.rut='"& Session("Rut") &"'"
	'response.Write(str)
	'response.end()
	
	set rst=Session("Conn").execute(str)
	rut=rst("codcli")
	session("rut")=rut
	dv=rst("dig")
	paterno=rst("paterno")
	materno=rst("materno")
	nombre=rst("nombre")
	estado=rst("DESCRIPCION")
	codestcivil = rst("CODESTCIVIL")
	fecnac=rst("fecnac")
	'response.write(fecnac)
	'response.End()	
	diractual=rst("diractual")
	comuna=rst("comuna")
	ciudad=rst("ciudadact")
	fonoact=rst("fonoact")
	dirproc=rst("dirproc")
	ciuproc=rst("ciuproc")
	comunaproc=rst("comunapro")
	fonoproc=rst("fonoproc")
	if rst("mail") <> "" then
		mail=replace(rst("mail"), ",",".")
	else
		mail=""
	end if
	if rst("mail_inst") <> "" then
		MailInstitucion=replace(rst("mail_inst"), ",",".")
	else
		MailInstitucion=""
	end if
	VALOR=""
end function
function Correo()
'dim strstr 
'dim rstrst
'strstr ="select habilitacorreo,correotitular,correocopia from mt_parame "
'set rstrst=Session("Conn").execute(strstr)
'if rstrst("habilitacorreo")="S" then
	'if mail <> "" then
'		Opiniones=request.Form("txtopiniones")
'mail=request.Form("txtmail")
'		Set ObjetoCDO = Server.CreateObject("CDONTS.NewMail")
'		'Asignamos las propiedades al objeto
'		ObjetoCDO.From = mail
'		ObjetoCDO.To = rstrst("correotitular")
'		ObjetoCDO.Subject = "Opiniones"
'		ObjetoCDO.Body = Opiniones
'		ObjetoCDO.Cc = rstrst("correocopia")
'		ObjetoCDO.AttachFile  Anexo 
'		Enviamos el e-mail
'		ObjetoCDO.Send
'		Destruimos el objeto
'		Set ObjetoCDO = Nothing
'		
'	end if
'end if
end function 
'PONEMOS RUTINA DE LLENADO DE COMBO DE ESTADOS CIVILES..
sql=""
sql="select codigo,descripcion from MT_ESTCIVIL with (nolock)"
set rTQ=Session("Conn").Execute (sql)
if not rTQ.EOF then
	cbestado=""
	pricarr=true	
	while not rTQ.EOF
		if estado ="" and pricarr then
			'sel=" selected"
			pricarr=false
			estado=rTQ("descripcion") 
		else
			if estado=rTQ("descripcion") then
				sel=" selected"
			else
				sel=""
			end if
		end if
		cbestado= cbestado & "<option value=""" & rTQ("codigo") & """ " & sel & ">"
		cbestado=	cbestado & rTQ("descripcion") & "</option>" & vbcrlf
		rTQ.MoveNext 
	wend
end if



%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title><%=Session("NombrePestana")%></title>

<link href="css/tex-normales.css" rel="stylesheet" type="text/css">
<link href="css/parrafo.css" rel="stylesheet" type="text/css">
<link href="css/ed.css" rel="stylesheet" type="text/css">
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1"> 
<script language="JavaScript" src="include/calendario.js">
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
  function okap()
  {
	alert("Sus Datos han sido Grabados, Vuelva a Ingresar");
	window.top.location="menu_perfil.asp";
  }
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
function F8BuscarComuna()
 {   
 	    var x=window.open ("F8Comuna.asp?Ciudad=" + document.TheForm.txtciudadact.value + "","HelpWindow","width=600,height=350,resizable=yes,scrollbars=yes");
 }
function F8BuscarCiudad()
 {  
    var x=window.open ("F8Ciudad.asp","HelpWindow","width=600,height=350,resizable=yes,scrollbars=yes");
 }
function F8BuscarComunaProcedencia()
 {   
 	    var x=window.open ("F8ComunaProc.asp?Ciudad=" + document.TheForm.txtciudadpro.value + "","HelpWindow","width=600,height=350,resizable=yes,scrollbars=yes");
 }
function F8BuscarCiudadProcedencia()
 {  
    var x=window.open ("F8CiudadProc.asp","HelpWindow","width=600,height=350,resizable=yes,scrollbars=yes");
 }
function EnviarValores (rut,dig,paterno,materno,nombre,estado,fecha,direccion,ciudad,comuna,fono,direccionproc,ciudadproc,comunaproc,fonoproc,mail){;
}
function LlamarLeyenda()
 {  
    var x=window.open ("Obligatorios.htm","HelpWindow","width=400,height=400,resizable=yes,scrollbars=yes");
 }
//-->
</script>
<script>
function ValidaMail(){     // CortesÌa de http://www.ejemplode.com

	if (TheForm.txtmail2.value !='')
	{
		re=/^[_a-z0-9-]+(.[_a-z0-9-]+)*@[a-z0-9-]+(.[a-z0-9-]+)*(.[a-z]{2,3})$/
		if(!re.exec(TheForm.txtmail2.value.toLowerCase()))    {
			TheForm.txtmail2.value ='';			
			alert("Escriba una direcciÛn de correo v·lida en el campo \"DirecciÛn de correo\".");
		}else{
			return true;
		}
	}
	else
		return true;
}

function ValidaMail123() {
if (((TheForm.txtmail.value.indexOf ('@', 0) == -1)||(TheForm.txtmail.value.length < 5)) && TheForm.txtmail.value != '') { 
	TheForm.txtmail.value ='';
	alert("Escriba una direcciÛn de correo v·lida en el campo \"DirecciÛn de correo\"."); 		
	return; 
  }
  return (true); 
  
}
 
function ValidaBlanco(){ 

	if  (document.TheForm.txtfecha.value=="") 	{
			alert("Debe ingresar su Fecha de Nacimiento");
			document.TheForm.elements['txtfecha'].focus();
			return false;
												}
	if  (document.TheForm.txtdireccionact.value=="") 	{
			alert("Debe ingresar tu direccion Actual");
			document.TheForm.elements['txtdireccionact'].focus();
			return false;
												}
	if  (document.TheForm.txtciudadact.value=="") 	{
			alert("Debe ingresar tu Ciudad Actual");
			document.TheForm.elements['txtciudadact'].focus();
			return false;
												}
	if  (document.TheForm.txtcomunaact.value=="") 	{
			alert("Debe ingresar tu Comuna Actual");
			document.TheForm.elements['txtcomunaact'].focus();
			return false;
												}
	if  (document.TheForm.txtfono.value=="") 	{
			alert("Debe ingresar tu Fono Actual");
			document.TheForm.elements['txtfono'].focus();
			return false;
												}	
	if  (document.TheForm.txtmail2.value=="") 	{
			alert("Debe ingresar E-Mail ");
			document.TheForm.elements['txtmail2'].focus();
			return false;
												}	
	return true;																					
																							
}
function Grabar() {

          	document.TheForm.VALOR.value="GRABAR";
   			document.TheForm.submit();
  }


function confirma()
{ 
	if (ValidaBlanco() == true) 
	{ 
		var respuesta=confirm("øDeseas grabar tus datos?");
	
		if (respuesta==true){
			document.TheForm.VALOR.value="GRABAR";
			document.TheForm.submit();
			response.Redirect(self)
			alert("Sus Datos han sido Grabados, Vuelva a Ingresar");
			window.top.location="menu_perfil.asp";
			}	
		else
			alert("Ha cancelado la ActualizaciÛn");
	}	
}

function Actualizar() {
	document.TheForm.submit();
}
</script>
<body bgColor=white>
<table border="0" cellpadding="0" cellspacing="0"  align="left">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right"><% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td valign="top" >
			<% CargarTop1()%><% SubMenu()%>
			<div style="width:825PX; height:600px; overflow:scroll;">
				<p align="left"><img src="Imagenes/titulos/T-actualiza_datos.gif"> </p>
				<p id="lblAntecedentes" class="sub-titulo">Antecedentes del Alumno </p>
				<%sT="select * from ra_mantieneactWEB with (nolock)"
				
				set Rt=Session("Conn").execute(sT)				
				if Rt.eof and RT.bof then
					Rt.close
					
					sTNO="SELECT 'NO'PATERNO ,'NO'MATERNO ,'NO'NOMBRES ,'NO'ESTADO_CIVIL ,'NO'FECHA_NACIMIENTO ,'NO'DIRECCION , "
					sTNO= sTNO & "'NO'CIUDAD ,'NO'COMUNA ,'NO'FONO ,'NO'DIRECCION_PROCEDENCIA ,'NO'CIUDAD_PROCEDENCIA ,'NO'COMUNA_PROCEDENCIA , "
					sTNO= sTNO & "'NO'FONO_PROCEDENCIA ,'NO'E_MAIL ,'NO'ID ,'NO'CELULARACT ,'NO'ANT_LABORALES ,'NO'ANT_EDUCACIONALES "					

					set Rt=Session("Conn").execute(sTNO)	
					
				end if  
				%>
				<form action="" method="post" name="TheForm" id="TheForm" >
				<input type="hidden" name="VALOR" value="<%=VALOR%>">
				  <table width="710" border="0">
					<tr> 
					  <td width="340" class="tex"><font id="lblCI" size="1">CI :</font><br> <input name="txtrut" id="txtrut4" disabled="true" value="<%=rut%>"  size="12" >
						- 
						<input name="txtdv" id="txtdv2"  disabled="true" value="<%=dv%>" size="2" maxlength="1"> 
						<br> </td>
					  <td colspan="3">&nbsp; </td>
					</tr>
					<tr> 
					  <%if rt("paterno")="S" then%>
					  <td class="tex">Apellido Paterno<br> <input name="txtpaterno" id="txtpaterno2"  value="<%=paterno%>" size="30"> 
					  <% else %>
					  <td class="tex">Apellido Paterno<br> <input name="txtpaterno" id="txtpaterno2" disabled="true"  value="<%=paterno%>" size="30"> 
					  <%end if%>
					  </td>
					  <%if rt("materno")="S" then%>
					  <td width="185" class="tex">Apellido Materno<br> <input name="txtmaterno" id="txtmaterno2"  value="<%=materno%>" size="30"> 
					  <%else%>
					  <td width="185" class="tex">Apellido Materno<br> <input name="txtmaterno" id="txtmaterno2" disabled="true" value="<%=materno%>" size="30"> 
					  <%end if%>
					  </td>
					  <td colspan="2" class="tex">&nbsp;</td>
					</tr>
					<tr> 
					  <%if rt("Nombres")="S" then%>
					  <td class="tex">Nombres<br> <input name="txtnombres" id="txtnombres2" value="<%=nombre%>" size="40"> 
					  <%else%>
					  <td class="tex">Nombres<br> <input name="txtnombres" id="txtnombres2" disabled="true" value="<%=nombre%>" size="40"> 
					  <%end if%>
						<br> </td>
					  <td colspan="3">&nbsp;</td>
					</tr>
					<tr> 
					 <%if rt("estado_civil")="S" then%>
					  <td class="tex">Estado Civil<br>
					  <select name="cbestado" id="cbestado" >
						<option value=""> Seleccione </option>
							<%=cbestado%> 
						</select>
						<%else%>
					  <td class="tex">Estado Civil<br>
						<select name="select" id="select3"  disabled="true" >
						  <option value=""> Seleccione </option>
						  <%=cbestado%> </select>
						<%end if%>
					  </td>
					  <td><span class="tex">Fecha Nacimiento</span> 
						<font color="#ff0000" size="4">*</font><br>
						  <%if rt("fecha_nacimiento")="S" then%>
						<input name="txtfecha" id="txtfecha3"  value="<%=fecnac%>" size="12" maxlength="12" >
						 <A href="javascript:putcal('TheForm','txtfecha')"> <IMG style="LEFT: 100px; TOP: 25px" height=20 src="ImagenBoton/Calend.gif" width=20 > 
						</A>
						<%else%>
						<input name="txtfecha" id="txtfecha3"  disabled="true" value="<%=fecnac%>" size="12" maxlength="12" >
						<%end if%>
				      </td>
					  <td width="87" class="tex"><br> </td>
					  <td width="180">&nbsp;</td>
					</tr>
				  </table>
				  <p class="sub-titulo">Datos Actuales ( residencia )</p>
				  <table width="710" border="0">
					<tr> 
					<%if rT("direccion")="S" then %>
					  <td width="328"   class="tex">Direcci&oacute;n <font color="#ff0000" size="4">*</font><br> <input name="txtdireccionact" id="txtdireccionact2" value="<%=diractual%>" size="50" >
					  <%else%>
					  <td width="328"   class="tex">Direcci&oacute;n <font color="#ff0000" size="4">*</font><br> <input name="txtdireccionact" id="txtdireccionact2"  disabled="true" value="<%=diractual%>" size="50" >
					  <%end if%>
					  </td> 
					  <td colspan="4">&nbsp; </td>
					</tr>
					<tr> 
					  <td class="tex"><font id="lblCiudadAct" size="1">Ciudad</font><font color="#ff0000" size="4">*</font><br> 
					  <%if rT("ciudad")="S" then %>
						<input name="txtciudadact" id="txtciudadact3" value="<%=ciudad%>" size="40" >
						<A href="javascript:F8BuscarCiudad();"><IMG height=18 src="ImagenBoton/f8.jpg" width=18></A> 
						<%else%>
						<input name="txtciudadact" id="txtciudadact3"  disabled="true" value="<%=ciudad%>" size="40" >
						<%end if%>
						
					  </td>
					  <td colspan="4"> <span id="lblComunaAct" class="tex">Comuna</span> <font color="#ff0000" size="4">*</font><br>
						<%if rt("comuna")="S" then  %>
						<input name="txtcomunaact" id="txtcomunaact" value="<%=comuna%>"  size="20">
						   <A href="javascript:F8BuscarComuna();"><IMG height=18 src="ImagenBoton/f8.jpg" width=18> 
						</A>
						<%else%>
					   <input name="txtcomunaact" id="txtcomunaact"  disabled="true" value="<%=comuna%>"  size="20">
					   <%end if%>
					  </td>
					</tr>
					<tr> 
					  <td class="tex">Tel&eacute;fono <font color="#ff0000" size="4">*</font><br>
					   <%if rt("fono")="S" then %>
						<input name="txtfono" id="txtfono2" value="<%=fonoact%>" size="15" >
						<%else%>
						<input name="txtfono" id="txtfono2"  disabled="true" value="<%=fonoact%>" size="15" >
						<%end if%>
					  </td>
					  <td width="212">&nbsp;</td>
					  <td width="105">&nbsp;</td>
					  <td width="47" colspan="2">&nbsp;</td>
					</tr>
				  </table>
				  <p class="sub-titulo">Datos de Procedencia ( origen )</p>
				  <table width="710" border="0">
					<tr> 
					  <td width="322" class="tex">Direcci&oacute;n <br> 
					  <%if rt("DIRECCION_PROCEDENCIA")="S" then %>
					  <input name="txtdireccionpro" id="txtdireccionpro2" value="<%=dirproc%>" size="50"> 
					  <%else%>
					  <input name="txtdireccionpro" id="txtdireccionpro2"  disabled="true" value="<%=dirproc%>" size="50"> 
					  <%end if%>
					  </td>
					  <td colspan="4">&nbsp; </td>
					</tr>
					<tr> 
					  <td class="tex"><font id="lblCiudadOri" size="1">Ciudad</font><br> <A href="javascript:F8BuscarCiudadProcedencia();"> 
						<%if rt("CIUDAD_PROCEDENCIA")="S" then%>
						<input name="txtciudadpro" id="txtciudadpro2" value="<%=ciuproc%>" size="40">
						<A href="javascript:F8BuscarCiudadProcedencia();"><IMG height=18 src="ImagenBoton/f8.jpg" width=18></A> 
						<%else%>
						<input name="txtciudadpro" id="txtciudadpro2"  disabled="true" value="<%=ciuproc%>" size="40">
						<%end if%>
					  
					  </td>
					  <td colspan="4"><span id="lblComunaOri" class="tex">Comuna</span> <br> 
						<%if rt("COMUNA_PROCEDENCIA")="S" then%>
						<input name="txtcomunapro" id="txtcomunapro2" value="<%=comunaproc%>" size="20"> 
						<A href="javascript:F8BuscarComunaProcedencia();"><IMG height=18 src="ImagenBoton/f8.jpg" width=18></A> 
						<%else%>
						<input name="txtcomunapro" id="txtcomunapro2"  disabled="true" value="<%=comunaproc%>" size="20"> 
						<%end if%>
						
					  </td>
					</tr>
					<tr> 
					  <td class="tex">Tel&eacute;fono <br>
						<%if rt("FONO_PROCEDENCIA")="S" then%>
						<input name="txtfonoproc" id="txtfonoproc3" value="<%=fonoproc%>" size="15" onBlur="Javascript:ValidaFono();"> 
						<%else%>
						<input name="txtfonoproc" id="txtfonoproc3"  disabled="true" value="<%=fonoproc%>" size="15" onBlur="Javascript:ValidaFono();"> 
						<%end if%>
					  </td>
		
					</tr>
				  </table>
				  <p class="sub-titulo">Tu Mail</p>
				  <table width="710" border="0">
					<tr> 
					  <td width="322" class="tex-normales">E-Mail Actual ( correo personal )<font color="#ff0000" size="4"> 
						* <span class="titulo2"> </span> </font><br> 
					<%if rt("e_mail")="S" then%>
					  <input name="txtmail" id="txtmail2" value="<%=mail%>" size="40" onBlur="javascript:ValidaMail();">
					  <%else%>
					  <input name="txtmail" id="txtmail2"  disabled="true" value="<%=mail%>" size="40" onBlur="javascript:ValidaMail();">
					<%end if%>
					  </td>
                      <td colspan="4" align="left"> 
                      
                      				  <DIV >&nbsp;<span class="resaltar-pregunta"><font size="1" face="Arial, Helvetica, sans-serif"> 
					</font></span><span class="tex-normales"><font size="1" face="Arial, Helvetica, sans-serif">E-Mail 
					Asi</font><font size="1">gnado por la Instituci&oacute;n <br>
					<input name="txtMailInstitucion" type="text" id="txtMailInstitucion"  disabled ="true" value="<%=MailInstitucion%>" size="40" maxlength="40">
					</font></span> </DIV> 
                      </td>
					</tr>
					<tr>
					    <td width="322" style="border:5">
    					<label id="msj_solicita_pass" visible="true" style="font-size:x-small"><%=MSJ_SOLICITA_CLAVE%></label>
					    </td>
					</tr>
				  </table>

				  <table width="710" border="0">
					<tr> 
					  <td width="216" class="tex">
					  </td>
					  <td width="484"></td>
					</tr>
				  </table>
				  <p>&nbsp;</p>
				  <table width="710" border="0">
					<tr> 
					  <td width="446">
                    <!--  <div align="left"><A onMouseOver="MM_swapImage('Image5','','ima/bot/ayuda-on.gif',1)" onmouseout=MM_swapImgRestore() href="#"> 
                      <img src="ima/bot/ayuda-of.gif" name="Image5"  border="0"></A></div>  -->
                      </td>
					  <td width="81"><div align="right"></div></td>
					  <td width="82"><a href="javascript:confirma();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image9','','Imagenes/botones/grabar-on.gif',1)"><img src="Imagenes/botones/grabar-of.gif" name="Image9"  border="0"></a></td>
					  <td width="83"><A onMouseOver="MM_swapImage('Image7','','Imagenes/botones/volver-on.gif',1)" onmouseout=MM_swapImgRestore() href="javascript: window.top.location.href ='menu_perfil.asp'"><IMG src="Imagenes/botones/volver-of.gif" border="0"  name="Image7"></A></td>
					</tr>
				  </table>
			  </form>
				<p>&nbsp;</p>
				<P>&nbsp;</P>
			  </div>
		  </tr>
	  </table>
	</td>
  </tr>
</table>
	</td>
  </tr>
</table>
</body>
<%ObjetosLocalizacion("Actualizador.asp")%>
</html>

<%
Function LimpiarTextoMail(ByVal texto)
    
    Dim objRegExp
    Set objRegExp = New Regexp
    
    objRegExp.IgnoreCase = True
    objRegExp.Global = True
    
    objRegExp.Pattern = "\s+"
    texto = objRegExp.Replace(texto, " ")
    
    objRegExp.Pattern = "[()?*"",\\<>&#~%{}+:\/!;']+"
    texto = objRegExp.Replace(texto, "")
    
    Dim i, s1, s2
    s1 = "¡¿…»Õœ”“⁄‹·‡ËÈÌÔÛÚ˙¸ÒÁ"
    s2 = "AAEEIIOOUUaaeeiioouunc-"
    If Len(texto) <> 0 Then
        For i = 1 To Len(s1)
            texto = Replace(texto, Mid(s1,i,1), Mid(s2,i,1))
        Next
    End If

    LimpiarTextoMail = LCase(texto)
End Function
%>
<!--#INCLUDE file="include/desconexion.inc" -->