var cont_click = 0;
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
function ValidaDatos() {
cont_click = cont_click + 1;
	rut = document.Login.LoginRut.value;
	clave = document.Login.LoginPassWord.value;

  var tmpstr = "";
  for ( i=0; rut.charAt(i) == '0' ; i++ );
  for (; i < rut.length ; i++ )
     tmpstr = tmpstr + rut.charAt(i);
  rut = tmpstr;
  document.Login.LoginRut.value=rut;

if (cont_click==1)
{
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
			return true;
		
	}
}
else
{
	return false;
}
}


function checkFields( rut, clave )
{

  var tmpstr = "";
  var SoloRut="";
  var RutNum;
  window.document.Login.pin.value =   window.document.Login.LoginPassWord.value;
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

  window.document.Login.rut.value = rut;

  //window.document.Login.LoginPassWord.value = "";
  //window.document.Login.LoginRut.value="";
  SoloRut=rut.substring(0,rut.length-1);

  RutNum=SoloRut
 document.Login.rut.value = document.Login.rut.value.toUpperCase();
 //document.passemp.rut.value = document.passemp.rut.value.toUpperCase();

  return true;
}

function checkCDV( dvr )

{
  dv = dvr + "";
  if ( dv != '0' && dv != '1' && dv != '2' && dv != '3' && dv != '4' && dv != '5' && dv != '6' && dv != '7' && dv != '8' && dv != '9' && dv != 'k'  && dv != 'K')
  {
    alert("El dígito verificador ingresado no es válido.");
    window.document.Login.LoginRut.focus();
    window.document.Login.LoginRut.select();
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
    window.document.Login.LoginRut.focus();
    window.document.Login.LoginRut.select();
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
    window.document.Login.LoginRut.focus();
    window.document.Login.LoginRut.value = "";
    return false;
  }

      return true;
}

function checkPinField()
{

 if ( window.document.Login.LoginPassWord.value.length < 3 )
  {
    alert("La clave debe poseer un largo mínimo de 4 digitos.");
    window.document.Login.LoginPassWord.focus();
    window.document.Login.LoginPassWord.select();
    return false;
  }
 if (ValidaCampo(window.document.Login.LoginPassWord.value))
  return true;
 else
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


function ValidaKey(){

if (window.document.Login.LoginPassWord.value.length>=3){
   window.document.Login.LoginRut.focus();
   }
}

function checkRutField(texto) {
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
  
  if ( largo < 2 ) {
    alert("Debe ingresar el RUT completo.");
    window.document.Login.LoginRut.focus();
    window.document.Login.LoginRut.select();
    return false;
  }

  for (i=0; i < largo ; i++ ) {
    if ( texto.charAt(i) !="0" && texto.charAt(i) != "1" && texto.charAt(i) !="2" && texto.charAt(i) != "3" && texto.charAt(i) != "4" && texto.charAt(i) !="5" && texto.charAt(i) != "6" && texto.charAt(i) != "7" && texto.charAt(i) !="8" && texto.charAt(i) != "9" && texto.charAt(i) !="k" && texto.charAt(i) != "K" ) {
      alert("El RUT ingresado no es válido.");
      window.document.Login.LoginRut.focus();
      window.document.Login.LoginRut.select();
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

  window.document.Login.LoginRut.value = invertido;

  if ( checkDV(texto) )
  {
    return true;
   }	
	
  return false;
}

function ComparaClaves(clave1,clave2){
  if (clave1 != clave2){
     alert("El nuevo password y su repetición No coinciden.");
     return false;
  }  
  return true;
}
