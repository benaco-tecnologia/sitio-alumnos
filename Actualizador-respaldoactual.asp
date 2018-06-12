<%
	Server.scripttimeout=900000
	response.buffer=false
	Response.ExpiresAbsolute = Now() - 1
	Response.AddHeader "pragma", "no-cache"
	Response.Expires = -1
    Session.Timeout = 999
%>
<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%
Dim str,Sql,sTRSQL,Sl,mySTR,sT,Strs,MStr
dim rst,Rql,rSTSQL,Rs,myRST,rT,rsts,MRst,rTQ
dim rut,dv,paterno,materno,nombre,estado,fecnac,diractual,comuna,ciudad,fonoact,dirproc,ciuproc,comunaproc,fonoproc
dim mail,Opiniones,Hora,Fecha,Carrera,MailInstitucion
'set rst=Server.CreateObject ("ADODB.Recordset")
dim CiudadActual
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

	rut=request.Form("txtrut")
	paterno=request.Form("txtpaterno")
	materno=request.Form("txtmaterno")
	nombre=request.Form("txtnombres")
	fecnac=request.Form("txtfecha")
	estado=request.Form("cbestado")
	diractual=request.Form("txtdireccionact")
	comuna=request.Form("txtcomunaact")
	ciudad=request.Form("txtciudadact")
	fonoact=request.Form("txtfono")
	dirproc=request.Form("txtdireccionpro")
	ciuproc=request.Form("txtciudadpro")
	comunaproc=request.Form("txtcomunapro")
	fonoproc=request.Form("txtfonoproc")
	mail=request.Form("txtmail")
	MailInstitucion=request.Form("txtmailinstitucion")
	
	mySTR="select * from ra_mantieneactWEB with (nolock) "
	set myRST=conn.execute(mySTR)
	Sql="Update mt_client set "
	if  myRST("paterno")="S" then
		Sql=Sql & " paterno='"& paterno & "',"
	end if
	if  myRST("materno")="S" then
		Sql=Sql & " materno='"& materno & "',"
	end if
	if  myRST("nombres")="S" then
		Sql=Sql & " nombre='"& nombre & "',"
	end if
	if  myRST("fecha_nacimiento")="S" then
		Sql=Sql & " fecnac='"& trim(fecnac) &"',"
	end if
	if  myRST("direccion")="S" then
		Sql=Sql & "diractual='"& trim(diractual) &"',"
	end if
	if  myRST("comuna")="S" then
		Sql=Sql & " comuna='"& trim(comuna) &"',"
	end if
	if  myRST("ciudad")="S" then
		Sql=Sql & "ciudadact='"& trim(ciudad) &"', "
	end if
	if  myRST("estado_civil")="S" then
		Sql=Sql & "estcivil='"& trim(estado) &"', "
	end if
	if  myRST("fono")="S" then
		Sql=Sql & " fonoact ='"& trim(fonoact) &"',"
	end if
	if  myRST("direccion_procedencia")="S" then
		Sql=Sql & " dirproc='"& trim(dirproc) &"',"
	end if
	if  myRST("ciudad_procedencia")="S" then
		Sql=Sql & " ciuproc='"& trim(ciuproc) &"',"
	end if
	if  myRST("comuna_procedencia")="S" then
		Sql=Sql & " comunapro='"& trim(comunaproc) &"', "
	end if
	if  myRST("fono_procedencia")="S" then
		Sql=Sql & " fonoproc='"& trim(fonoproc) &"',"
	end if
	if  myRST("e_mail")="S" then
		Sql=Sql & " mail='"& trim(replace(mail,",",".")) &"'"
	end if
	Sql=Sql &  " where codcli='"& trim(session("rut")) &"'" 
	'response.Write(sql)
	'response.end 
	set Rql=conn.execute(Sql)
	Strs="insert into SIS_REG_ACTUALIZADATOS (RUT) "
	Strs=Strs & " values ('"& trim(session("rut")) &"') "
	set rsts=conn.execute(Strs)
	
	'Se debiò acrualizar el registro.
	'sTRSQL="Select * from RA_OPINIONES WHERE rut='"& rut &"' and fecha = '"& fecha &"'"
	'set rSTSQL=conn.execute(sTRSQL)
	 'if rSTSQL.eof then 
	 '	 sl="Insert into RA_OPINIONES (rut,fecha,hora,codcarr,opinion) "
	 '	 sl=sl & " values ('"& rut &"','"& fecha &"','"& Hora &"','"& Carrera &"','"& Opiniones &"') "
	 '	 set rs=conn.execute(sl)
	 'else
	 ' 	sl="Update RA_OPINIONES set hora='"& Hora &"',opinion='"& Opiniones &"' "
     '	sl=sl & " where rut='"& rut &"' and fecha='"& fecha &"' and codcarr='"& Carrera &"'"
     '		 set rs=conn.execute(sl)
	 ' end if
	 '**************** Se terminò la actualizacion ******************
VALOR=""
end Function
function Llamar ()
	str="select  distinct b.codcli, b.dig, b.paterno,b.materno,b.nombre,b.estcivil,b.fecnac, "
	str=str & " b.diractual,b.comuna,b.ciudadact,b.fonoact,b.dirproc,b.ciuproc,b.comunapro,b.fonoproc, "
	str=str & " replace(b.mail,',','.') as mail "
	str=str & " from mt_alumno a with (nolock), mt_client  b with (nolock)"
	str=str & " where a.rut=b.codcli and "
	str=str & " a.rut='"& Session("RutCli") &"'"
	'response.Write(str)
	'response.end()
	
	set rst=conn.execute(str)
	rut=rst("codcli")
	session("rut")=rut
	dv=rst("dig")
	paterno=rst("paterno")
	materno=rst("materno")
	nombre=rst("nombre")
	estado=rst("estcivil")
	fecnac=rst("fecnac")
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
	MStr=" select e_mail from mt_webmail with (nolock) where rut ='"& Session("RutCli")  &"' "
	set MRst=conn.execute(MStr)
	if not Mrst.eof then
	  	MailInstitucion=MRst("e_mail")
	end if
	VALOR=""
end function
function Correo()
'dim strstr 
'dim rstrst
'strstr ="select habilitacorreo,correotitular,correocopia from mt_parame "
'set rstrst=conn.execute(strstr)
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
sql="select descripcion from ra_estacivil with (nolock)"
set rTQ=Conn.Execute (sql)
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
		cbestado= cbestado & "<option value=""" & rTQ("descripcion") & """ " & sel & ">"
		cbestado=	cbestado & rTQ("descripcion") & "</option>" & vbcrlf
		rTQ.MoveNext 
	wend
end if


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/css/tex-normales.css" rel="stylesheet" type="text/css">
<link href="css/css/parrafo.css" rel="stylesheet" type="text/css">
<link href="css/ed.css" rel="stylesheet" type="text/css">
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<script language="JavaScript" src="include/calendario.js">
</script>
<script language="JavaScript" type="text/JavaScript">
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
function ValidaMail() {
if ((TheForm.txtmail.value.indexOf ('@', 0) == -1)||(TheForm.txtmail.value.length < 5)) { 
    alert("Escriba una dirección de correo válida en el campo \"Dirección de correo\"."); 
    return; 
  }
  return (true); 
}
function ValidaBlanco(){

	if  (document.TheForm.txtfecha.value=="") 	{
			alert("Debe ingresar su Fecha de Nacimiento");
			document.TheForm.elements['txtfecha'].focus();
			return;
												}
	if  (document.TheForm.txtdireccionact.value=="") 	{
			alert("Debe ingresar tu direccion Actual");
			document.TheForm.elements['txtdireccionact'].focus();
			return;
												}
	if  (document.TheForm.txtciudadact.value=="") 	{
			alert("Debe ingresar tu Ciudad Actual");
			document.TheForm.elements['txtciudadact'].focus();
			return;
												}
	if  (document.TheForm.txtcomunaact.value=="") 	{
			alert("Debe ingresar tu Comuna Actual");
			document.TheForm.elements['txtcomunaact'].focus();
			return;
												}
	if  (document.TheForm.txtfono.value=="") 	{
			alert("Debe ingresar tu Fono Actual");
			document.TheForm.elements['txtfono'].focus();
			return;
												}												
}
function Grabar() {

          	document.TheForm.VALOR.value="GRABAR";
   			document.TheForm.submit();
  }

function confirma()
{
var respuesta=confirm("Deseas Grabar tus Datos.. ? ");

if (respuesta==true){
alert("Sus Datos han sido Grabados, Vuelva a Ingresar");
   
   	document.TheForm.VALOR.value="GRABAR";
   	document.TheForm.submit();
	window.top.location="menu_perfil.asp";
	}		

	

else
alert("Ha CANCELADO la Actualización");
}
function Actualizar() {
  document.TheForm.submit();
}
</script>
<body background="ima/iso-uhc.gif" onLoad="MM_preloadImages('ima/bot/ayuda-on.gif','ima/bot/salir-on.gif','ImagenBoton/grabar-on.gif')" bgColor=white>
<table border="0" cellpadding="0" cellspacing="0"  align="center">
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
				<p align="left"><img src="Imagenes/T-actualiza_datos.gif"> </p>
				<p class="sub-titulo">Antecedentes del Alumno </p>
				<%sT="select * from ra_mantieneactWEB with (nolock)"
				set Rt=conn.execute(sT)
				%>
				<form action="" method="post" name="TheForm" id="TheForm" >
				<input type="hidden" name="VALOR" value="<%=VALOR%>">
				  <table width="710" border="0">
					<tr> 
					  <td width="240" class="tex">CI :<br> <input name="txtrut" id="txtrut4" disabled="true" value=" <%=rut%>"  size="12" >
						- 
						<input name="txtdv" id="txtdv2"  disabled="true" value=" <%=dv%>" size="2" maxlength="1"> 
						<br> </td>
					  <td colspan="3">&nbsp; </td>
					</tr>
					<tr> 
					  <%if rt("paterno")="S" then%>
					  <td class="tex">Apellido Paterno<br> <input name="txtpaterno" id="txtpaterno2"  value=" <%=paterno%>" size="30"> 
					  <% else %>
					  <td class="tex">Apellido Paterno<br> <input name="txtpaterno" id="txtpaterno2" disabled="true"  value=" <%=paterno%>" size="30"> 
					  <%end if%>
					  </td>
					  <%if rt("materno")="S" then%>
					  <td width="185" class="tex">Apellido Materno<br> <input name="txtmaterno" id="txtmaterno2"  value=" <%=materno%>" size="30"> 
					  <%else%>
					  <td width="185" class="tex">Apellido Materno<br> <input name="txtmaterno" id="txtmaterno2" disabled="true" value=" <%=materno%>" size="30"> 
					  <%end if%>
					  </td>
					  <td colspan="2" class="tex">&nbsp;</td>
					</tr>
					<tr> 
					  <%if rt("Nombres")="S" then%>
					  <td class="tex">Nombres<br> <input name="txtnombres" id="txtnombres2" value=" <%=nombre%>" size="40"> 
					  <%else%>
					  <td class="tex">Nombres<br> <input name="txtnombres" id="txtnombres2" disabled="true" value=" <%=nombre%>" size="40"> 
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
					  <td><span class="tex">Fecha Nacimiento</span> <A href="javascript:LlamarLeyenda();"> 
						<font color="#ff0000" size="4">*</font><br>
						  <%if rt("fecha_nacimiento")="S" then%>
						<input name="txtfecha" id="txtfecha3"  value=" <%=fecnac%>" size="12" maxlength="12" >
						 <A href="javascript:putcal('TheForm','txtfecha')"> <IMG style="LEFT: 100px; TOP: 25px" height=20 src="ImagenBoton/Calend.gif" width=20 > 
						</A>
						<%else%>
						<input name="txtfecha" id="txtfecha3"  disabled="true" value=" <%=fecnac%>" size="12" maxlength="12" >
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
					  <td width="328"   class="tex">Direcciòn <font color="#ff0000" size="4">*</font><br> <input name="txtdireccionact" id="txtdireccionact2" value=" <%=diractual%>" size="50" onBlur="javascript:ValidaBlanco();">
					  <%else%>
					  <td width="328"   class="tex">Direcciòn <font color="#ff0000" size="4">*</font><br> <input name="txtdireccionact" id="txtdireccionact2"  disabled="true" value=" <%=diractual%>" size="50" onBlur="javascript:ValidaBlanco();">
					  <%end if%>
					  </td> 
					  <td colspan="4">&nbsp; </td>
					</tr>
					<tr> 
					  <td class="tex">Ciudad <font color="#ff0000" size="4">*</font><br> 
					  <%if rT("ciudad")="S" then %>
						<input name="txtciudadact" id="txtciudadact3" value=" <%=ciudad%>" size="40" onBlur="javascript:ValidaBlanco()">
						<A href="javascript:F8BuscarCiudad();"><IMG height=18 src="ImagenBoton/f8.jpg" width=18></A> 
						<%else%>
						<input name="txtciudadact" id="txtciudadact3"  disabled="true" value=" <%=ciudad%>" size="40" onBlur="javascript:ValidaBlanco()">
						<%end if%>
						
					  </td>
					  <td colspan="4"> <span class="tex">Comuna</span> <font color="#ff0000" size="4">*</font><br>
						<%if rt("comuna")="S" then  %>
						<input name="txtcomunaact" id="txtcomunaact" value=" <%=comuna%>" onBlur="javascript:ValidaBlanco();" size="20">
						   <A href="javascript:F8BuscarComuna();"><IMG height=18 src="ImagenBoton/f8.jpg" width=18> 
						</A>
						<%else%>
					   <input name="txtcomunaact" id="txtcomunaact"  disabled="true" value=" <%=comuna%>" onBlur="javascript:ValidaBlanco();" size="20">
					   <%end if%>
					  </td>
					</tr>
					<tr> 
					  <td class="tex">Telèfono <font color="#ff0000" size="4">*</font><br>
					   <%if rt("fono")="S" then %>
						<input name="txtfono" id="txtfono2" value=" <%=fonoact%>" size="15" onBlur="javascript:ValidaBlanco();">
						<%else%>
						<input name="txtfono" id="txtfono2"  disabled="true" value=" <%=fonoact%>" size="15" onBlur="javascript:ValidaBlanco();">
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
					  <td width="322" class="tex">Direccion <br> 
					  <%if rt("DIRECCION_PROCEDENCIA")="S" then %>
					  <input name="txtdireccionpro" id="txtdireccionpro2" value=" <%=dirproc%>" size="50"> 
					  <%else%>
					  <input name="txtdireccionpro" id="txtdireccionpro2"  disabled="true" value=" <%=dirproc%>" size="50"> 
					  <%end if%>
					  </td>
					  <td colspan="4">&nbsp; </td>
					</tr>
					<tr> 
					  <td class="tex">Ciudad<br> <A href="javascript:F8BuscarCiudadProcedencia();"> 
						<%if rt("CIUDAD_PROCEDENCIA")="S" then%>
						<input name="txtciudadpro" id="txtciudadpro2" value=" <%=ciuproc%>" size="40">
						<A href="javascript:F8BuscarCiudadProcedencia();"><IMG height=18 src="ImagenBoton/f8.jpg" width=18></A> 
						<%else%>
						<input name="txtciudadpro" id="txtciudadpro2"  disabled="true" value=" <%=ciuproc%>" size="40">
						<%end if%>
					  
					  </td>
					  <td colspan="4"><span class="tex">Comuna</span> <br> 
						<%if rt("COMUNA_PROCEDENCIA")="S" then%>
						<input name="txtcomunapro" id="txtcomunapro2" value=" <%=comunaproc%>" size="20"> 
						<A href="javascript:F8BuscarComunaProcedencia();"><IMG height=18 src="ImagenBoton/f8.jpg" width=18></A> 
						<%else%>
						<input name="txtcomunapro" id="txtcomunapro2"  disabled="true" value=" <%=comunaproc%>" size="20"> 
						<%end if%>
						
					  </td>
					</tr>
					<tr> 
					  <td class="tex">Telèfono <br>
						<%if rt("FONO_PROCEDENCIA")="S" then%>
						<input name="txtfonoproc" id="txtfonoproc3" value=" <%=fonoproc%>" size="15" onBlur="Javascript:ValidaFono();"> 
						<%else%>
						<input name="txtfonoproc" id="txtfonoproc3"  disabled="true" value=" <%=fonoproc%>" size="15" onBlur="Javascript:ValidaFono();"> 
						<%end if%>
					  </td>
					  <td width="271">&nbsp;</td>
					  <td width="98">&nbsp;</td>
					  <td width="1" colspan="2">&nbsp;</td>
					</tr>
				  </table>
				  <p class="sub-titulo">Tu Mail</p>
				  <table width="710" border="0">
					<tr> 
					  <td class="tex-normales">E-Mail Actual ( correo personal )<font color="#ff0000" size="4"> 
						* <span class="titulo2"> </span> </font><br> 
					<%if rt("e_mail")="S" then%>
					  <input name="txtmail" id="txtmail2" value=" <%=mail%>" size="40" onBlur="javascript:ValidaMail();">
					  <%else%>
					  <input name="txtmail" id="txtmail2"  disabled="true" value=" <%=mail%>" size="40" onBlur="javascript:ValidaMail();">
					<%end if%>
					  </td>
					</tr>
				  </table>
				  <DIV align=center>&nbsp;<span class="resaltar-pregunta"><font size="1" face="Arial, Helvetica, sans-serif"> 
					</font></span><span class="tex-normales"><font size="1" face="Arial, Helvetica, sans-serif">E-Mail 
					Asis</font><font size="1">gnado por la Instituci&oacute;n <br>
					<input name="txtMailInstitucion" type="text" id="txtMailInstitucion"  disabled ="true" value="<%=MailInstitucion%>" size="40" maxlength="40">
					</font></span> </DIV>
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
					  <td width="446"><div align="left"><A onMouseOver="MM_swapImage('Image5','','ima/bot/ayuda-on.gif',1)" onmouseout=MM_swapImgRestore() href="#"> 
						  </A>&nbsp;</div></td>
					  <td width="81"><div align="right"></div></td>
					  <td width="82"><a href="javascript:confirma();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image9','','Imagenes/grabar-on.gif',1)"><img src="Imagenes/grabar-of.gif" name="Image9" width="83" height="21" border="0"></a></td>
					  <td width="83"><A onMouseOver="MM_swapImage('Image7','','Imagenes/salir-on.gif',1)" onmouseout=MM_swapImgRestore() href="javascript:parent.parent.mainFrame.history.go(-1)"><IMG src="Imagenes/salir-of.gif" width="83" border="0" height="21" name="Image7"></A></td>
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
</html>
<!--#INCLUDE file="include/desconexion.inc" -->