<!--
if(document.getElementById&&!document.all){ns6=1;}else{ns6=0;}
var agtbrw=navigator.userAgent.toLowerCase();
var operaaa=(agtbrw.indexOf('opera')!=-1);
var head="display:''";
var folder='';
var OldStyle='';
var SwStyle='';

function Mid(s, n, c){
	var numargs=Mid.arguments.length;

	if(numargs<3)
	c=s.length-n+1;

	if(c<1)
	c=s.length-n+1;

	if(n+c >s.length)
	c=s.length-n+1;

	if(n>s.length)
	return "";
return s.substring(n-1,n+c-1);
}

function IsDate(s){
	var Tentativa = new Date(s);
	return !isNaN(Tentativa)
}

function ValidaDatos(){
	xLoginRut  = document.Login.LoginRut.value;
	xLoginPass = document.Login.LoginPassWord.value;
	if (xLoginRut.length == 0 || xLoginPass.length==0) {
		alert ("Debe Ingresar Login y Password!");
		return false;
	}
	return true;
}

function expandit(curobj){ 
	if(document.getElementById(curobj)){
  		folder=document.getElementById(curobj).style;
  	} else {
		if(ns6==1||operaaa==true){
				folder=curobj.nextSibling.nextSibling.style;
		} else {
			folder=document.all[curobj.sourceIndex+1].style;
		  }
   	  }
	if (folder.display=="none"){folder.display="";}else{folder.display="none";}
}

function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

function jsGrabaAdmin(){
	if (jsValidaInputDataAdmin()){
		document.frm.action="admin_graba.asp";
		document.frm.submit();
	}
}

function jsValidaInputDataAdmin(){
	if (document.frm.SelectClasificacion.value == "0"){
		alert ("Debe Seleccionar una Clasificación del Usuario.!");
		document.frm.SelectClasificacion.focus();
		return false;
	}

	if (document.frm.NombreAdmin.value == ""){
		alert ("Ingrese la Identificación del Administrador!");
		document.frm.NombreAdmin.focus();
		return false;
	}

	if (document.frm.User.value == ""){
		alert ("Ingrese Nombre de Usuario!");
		document.frm.User.focus();
		return false;
	}

	if (document.frm.Pass.value == "" ||  document.frm.rePass.value == ""){
		alert ("Contraseñas no pueden estar en blanco!");
		document.frm.Pass.focus();
		return false;
	}

	if (document.frm.Pass.value != document.frm.rePass.value){
		alert ("Contraseñas no coinciden.!");
		document.frm.Pass.focus();
		return false;
	}
	return true
}

function jsGrabaScroll(IDmsg){
	if (jsValidaInputDataScroll()){
		document.frm.action="scroll_graba.asp?IDmsg"+IDmsg;
		document.frm.submit();
	}
}

function jsGrabaEvento(IDmsg){
	if (jsValidaInputDataEvento()){
		document.frm.action="eventos_graba.asp?IDmsg"+IDmsg;
		document.frm.submit();
	}
}

function jsValidaInputDataScroll(){
	if (document.frm.FechaInicio.value.length<10 || document.frm.FechaInicio.value=="" || IsDate(Mid(document.frm.FechaInicio.value,4,2) + Mid(document.frm.FechaInicio.value,1,2) + Mid(document.frm.FechaInicio.value,7,4))){
		alert("Fecha de Inicio Erronea!");
		document.frm.FechaInicio.focus();
		return false;
	}
	if (document.frm.FechaTermino.value.length<10 || document.frm.FechaTermino.value=="" || IsDate(Mid(document.frm.FechaTermino.value,4,2) + Mid(document.frm.FechaTermino.value,1,2) + Mid(document.frm.FechaTermino.value,7,4))){
		alert("Fecha de Termino Erronea!");
		document.frm.FechaTermino.focus();
		return false;
	}
	if (document.frm.Mensaje.value == ""){
		alert ("Debe Agregar un mensaje!");
		document.frm.Mensaje.focus();
		return false;
	}	
	return true
}
function jsGrabaNoticia(IDmsg){
	if (jsValidaInputDataNoticia()){
		document.frm.action="noticias_graba.asp?IDmsg"+IDmsg;
		document.frm.submit();
	}
}
function jsValidaInputDataNoticia(){
	if (document.frm.FechaInicio.value.length<10 || document.frm.FechaInicio.value=="" || IsDate(Mid(document.frm.FechaInicio.value,4,2) + Mid(document.frm.FechaInicio.value,1,2) + Mid(document.frm.FechaInicio.value,7,4))){
		alert("Fecha de Inicio Erronea!");
		document.frm.FechaInicio.focus();
		return false;
	}
	if (document.frm.FechaTermino.value.length<10 || document.frm.FechaTermino.value=="" || IsDate(Mid(document.frm.FechaTermino.value,4,2) + Mid(document.frm.FechaTermino.value,1,2) + Mid(document.frm.FechaTermino.value,7,4))){
		alert("Fecha de Termino Erronea!");
		document.frm.FechaTermino.focus();
		return false;
	}
	if (document.frm.Encabezado.value == ""){
		alert ("Debe Agregar un Encabezado!");
		document.frm.Encabezado.focus();
		return false;
	}	
	return true
}

function jsValidaInputDataEvento(){
	if (document.frm.FechaInicio.value.length<10 || document.frm.FechaInicio.value=="" || IsDate(Mid(document.frm.FechaInicio.value,4,2) + Mid(document.frm.FechaInicio.value,1,2) + Mid(document.frm.FechaInicio.value,7,4))){
		alert("Fecha de Inicio Erronea!");
		document.frm.FechaInicio.focus();
		return false;
	}
	if (document.frm.FechaTermino.value.length<10 || document.frm.FechaTermino.value=="" || IsDate(Mid(document.frm.FechaTermino.value,4,2) + Mid(document.frm.FechaTermino.value,1,2) + Mid(document.frm.FechaTermino.value,7,4))){
		alert("Fecha de Termino Erronea!");
		document.frm.FechaTermino.focus();
		return false;
	}
	if (document.frm.Mensaje.value == ""){
		alert ("Debe Agregar un mensaje!");
		document.frm.Mensaje.focus();
		return false;
	}	
	return true
}

function jsReLoadProfe(){
	document.frm.action="apuntes.asp?s=1";
	document.frm.submit();
}

function jsGrabaDocumento(IDrut){
	if (jsValidaInputDocumento()){
		return true;
//		document.frm.action="apuntes_graba.asp?ID="+IDrut;
//		document.frm.submit();
	} else {
		return false;
	}
}

function jsValidaInputDocumento(){
	if (document.frm.Descripcion.value==""){
		alert ("Escriba una descripción breve reconocimiento posterior");
		document.frm.Descripcion.focus();
		return false;
	}
	if (document.frm.txtSelectCarrera.value=="0"){
		alert ("Debe Seleccionar Una Carrera");
		document.frm.txtSelectCarrera.focus();
		return false;
	}
	if (document.frm.SelectAsignatura.value=="0"){
		alert ("Debe Seleccionar Una Asignatura");
		document.frm.SelectAsignatura.focus();
		return false;
	}	
	if (document.frm.Docto1.value=="" && document.frm.Docto2.value=="" && document.frm.Docto3.value==""){
		alert ("Debe Seleccionar Al Menos un Documento");
		document.frm.SelectAsignatura.focus();
		return false;
	}
	return true;
}
//-->