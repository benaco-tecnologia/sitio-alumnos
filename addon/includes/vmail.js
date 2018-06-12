<!--
function validarEmailDatos(inForm){
	if (inForm.NombreRemitente.value == ""){
		alert ("DEBE INGRESAR REMITENTE...!!!");
		return false;
	}
	if (!validarEmail(inForm)){
		return false;
	}
}

 function validarEmail(inForm) {
  if (/^\w+([\.-]?\w+)*@\w+([\.-]?\w+)*(\.\w{2,3})+$/.test(inForm.MailRemitente.value)){
   // alert("La dirección de email " + inForm.MailRemitente.value + " es correcta.") 
   return (true)
  } else {
   alert("La dirección de email es incorrecta.");
   return (false);
  }
 }
 
 function validarEnvioMailRegistro(inForm){
	if (inForm.SelectRemitente.value == "0"){
		alert ("DEBE SELECCIONAR UN REMITENTE");
		inForm.SelectRemitente.focus();
		return false;
	}
	if (inForm.Asunto.value == ""){
		alert ("DEBE INGRESAR ASUNTO DEL MAIL");
		inForm.Asunto.focus();
		return false;
	}
	if (inForm.SelectCarrera.value == "0" && inForm.SelectProfesor.value == "0" && inForm.SelectCoordinador.value == "0"){
		alert ("DEBE SELECCIONAR AL MENOS UN DESTINATARIO");
		inForm.SelectCarrera.focus();
		return false;
	}
	if (inForm.Mensaje.value == ""){
		alert ("DEBE ESCRIBIR MENSAJE");
		inForm.Mensaje.focus();
		return false;
	}
}
 
 
 //-->