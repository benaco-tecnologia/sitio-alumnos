<!--#INCLUDE FILE="include/conexion.inc" -->

<html>
<head>
<meta charset="utf-8">
<title>Formulario de Inscripción</title> 

<script language="JavaScript">

function validaDatos(f)
{	
	if ( f.rut.value == null || f.rut.value  == '' ) 
	{ 
		alert ('El campo Rut es obligatorio');  
		f.rut.focus(); 
		return false; 
	} 
	
	if ( f.nombre.value  == '' ) 
	{ 
		alert ('El campo Nombre es obligatorio'); 
		f.nombre.focus(); 
		return false; 
	} 
	
	if ( f.apellido.value   == '' ) 
	{ 
		alert ('El campo Apellido es obligatorio');  
		f.apellido.focus(); 
		return false; 
	} 
	
	if ( f.direccion.value  == '' ) 
	{ 
		alert ('El Direcci\u00f3n es obligatorio'); 
		f.direccion.focus(); 
		return false; 
	} 
	
	if ( f.telMovil.value   == '' ) 
	{ 
		alert ('El campo Tel\u00e9fono M\u00f3vil es obligatorio');  
		f.telMovil.focus(); 
		return false; 
	} 
	
	if ( f.nombreTutor.value  == '' ) 
	{ 
		alert ('El campo Nombre del Tutor es obligatorio'); 
		f.nombreTutor.focus(); 
		return false; 
	} 
	
	if ( f.telTutor.value   == '' ) 
	{ 
		alert ('El campo Tel\u00e9fono del Tutor es obligatorio');  
		f.telTutor.focus(); 
		return false; 
	} 
	
	var x = document.forms["form"]["confirmaEmail"].value;
	var y = document.forms["form"]["email"].value;
	if ( x != y )
	{
		alert("confirmación de e-mail incorrecta");
		document.forms["form"]["confirmaEmail"].value = "";
		f.confirmaEmail.focus();
		return false;
	}
	
	return true;	
}

function checkRutField(texto)
{
  if(texto != "")
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
	  var rut = document.getElementById("rut");
	  rut.value = "";
	  rut.focus();
	  rut.select();
	  //document.getElementById('rut').focus();
	  //document.getElementById("rut").select();
	  return false;
	}
  
	for (i=0; i < largo ; i++ )
	{
	  if ( texto.charAt(i) !="0" && texto.charAt(i) != "1" && texto.charAt(i) !="2" && texto.charAt(i) != "3" && texto.charAt(i) != "4" && texto.charAt(i) !="5" && texto.charAt(i) != "6" && texto.charAt(i) != "7" && texto.charAt(i) !="8" && texto.charAt(i) != "9" && texto.charAt(i) !="k" && texto.charAt(i) != "K" )
	  {
		alert("El RUT ingresado no es válido.");
		var rut = document.getElementById("rut");
		rut.value = "";
		rut.focus();
		rut.select();
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
  
	window.document.form.rut.value = invertido;
  
	if ( checkDV(texto) )
	{
	  return true;
	}	
	  
	return false;
  }
}

function checkDV( crut )
{
  largo = crut.length;
  if ( largo < 2 )
  {
    alert("Por favor ingrese un RUT válido.");
	  var rut = document.getElementById("rut");
	  rut.value = "";
	  rut.focus();
	  rut.select();
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

  if ( dvr != dv.toLowerCase() )
  {
    alert("El RUT ingresado es incorrecto.");
	  var rut = document.getElementById("rut");
	  rut.value = "";
	  rut.focus();
	  rut.select();
    return false;
  }
      return true;
}

function checkCDV( dvr )
{
  dv = dvr + "";
  if ( dv != '0' && dv != '1' && dv != '2' && dv != '3' && dv != '4' && dv != '5' && dv != '6' && dv != '7' && dv != '8' && dv != '9' && dv != 'k'  && dv != 'K')
  {
    alert("El dígito verificador ingresado no es válido.");
	  var rut = document.getElementById("rut");
	  rut.value = "";
	  rut.focus();
	  rut.select();
    return false;
  }
  return true;
}

  function soloLetras(e) 
  {
	  key = e.keyCode || e.which;
	  tecla = String.fromCharCode(key).toLowerCase();
	  letras = " áéíóúabcdefghijklmnñopqrstuvwxyz";
	  especiales = [8, 9, 37, 39, 46];
  
  	var keyCode = e.keyCode || e.which; 


		
	  tecla_especial = false
	  for(var i in especiales) 
	  {
		  if(key == especiales[i]) 
		  {
			  tecla_especial = true;
			  break;
		  }
	  }  
	  if(letras.indexOf(tecla) == -1 && !tecla_especial)
		  return false;
  }
  
  function soloNumeros(e) 
  {
	key = e.keyCode || e.which;
	tecla = String.fromCharCode(key).toLowerCase();
	numeros = " -0123456789";
	especiales = [8, 9, 37, 39, 46];

	tecla_especial = false
	for(var i in especiales) 
	{
		if(key == especiales[i]) 
		{
			tecla_especial = true;
			break;
		}
	}
	if(numeros.indexOf(tecla) == -1 && !tecla_especial)
		return false;
  }
  
  function soloRut(e) 
  {
	key = e.keyCode || e.which;
	tecla = String.fromCharCode(key).toLowerCase();
	numeros = ".-0123456789k";
	especiales = [8, 9, 37, 39, 46];

	tecla_especial = false
	for(var i in especiales) 
	{
		if(key == especiales[i]) 
		{
			tecla_especial = true;
			break;
		}
	}
	if(numeros.indexOf(tecla) == -1 && !tecla_especial)
		return false;
  }
  
  function validaEmail() 
  {
	var x = document.forms["form"]["email"].value;
	if( x != "")
	{	  
	  var atpos = x.indexOf("@");
	  var dotpos = x.lastIndexOf(".");
	  if (atpos< 1 || dotpos<atpos+2 || dotpos+2>=x.length) 
	  {
		  alert("e-mail no válido");
		  document.forms["form"]["email"].value = "";
		  return false;
	  }
	}
  }
  
  function validaEmail2() 
  {
	var x = document.forms["form"]["confirmaEmail"].value;
	var y = document.forms["form"]["email"].value;
	if ( x == y )
	{
	  if( x != "")
	  {
		
		var atpos = x.indexOf("@");
		var dotpos = x.lastIndexOf(".");
		if ( atpos < 1 || dotpos < atpos + 2 || dotpos + 2 >= x.length ) 
		{
			alert("e-mail de confirmación no válido");
			document.forms["form"]["confirmaEmail"].value = "";
			return false;
		}
	  }
	}
	else
	{
		alert("confirmación de e-mail incorrecta");
		document.forms["form"]["confirmaEmail"].value = "";
		return false;
	}
  }
  
</script>  
   
<style>
	body {font-family: Impact, Haettenschweiler, "Franklin Gothic Bold", "Arial Black", sans-serif;} 
</style>

</head>

<body style="font-family:arial; font:bold">
<table align="center" >
<tr>
<td>
<br>
<br>
<fieldset style="border-radius:10px;border:3px solid #C30">
<h3 align="center" >Formulario de Inscripción</h3>
<h5 align="center">Completa tus datos y participa en el sorteo de una Tablet</h5>
<table align="center" width="600px">
<hr width=98% align="center" style="border:1px solid #C30">
<br>
	<form name="form" method="post" action="grabainscripcion.asp" style="font-family:'Lucida Grande', 'Lucida Sans Unicode', 'Lucida Sans', 'DejaVu Sans', Verdana, sans-serif; font:bold" onsubmit="return validaDatos(this);" >

<tr>
  <td width="210px">  
  Rut *
  </td>
  <td > 
	  :   	
	</td>
  <td>
  	<input type="text" id="rut" name="rut" onBlur="return checkRutField(this.value);" maxlength="30" size="45" style="border-radius:5px" onKeyPress="return soloRut(event)"><br>
  </td>
    <td align="right" width="55px">
	</td>
</tr>
<tr>
  <td>
  Nombre *
  </td>
  <td >
  	:    	
	</td>
  <td>
  	<input id="nombre" type="text" name="nombre" size="45" onkeypress="return soloLetras(event)" maxlength="100" style="border-radius:5px"><br>
  </td>
</tr>
<tr>
  <td>
  Apellido *
  </td>
  <td >
  	:    	
	</td>
  <td>
  	<input type="text" name="apellido" size="45" onkeypress="return soloLetras(event)" maxlength="50" style="border-radius:5px"><br>
  </td>
</tr>
<tr>
  <td>
  Dirección *
  </td>
  <td > 
  	:   	
	</td>
  <td>
  	<input type="text" name="direccion" size="45" maxlength="200" style="border-radius:5px"><br>
  </td>
</tr>
<tr>
  <td>
  E-Mail
  </td>
  <td > 
  	:   	
	</td>
  <td>
    <input type="text" name="email" size="45" maxlength="100" style="border-radius:5px" onBlur="return validaEmail()"><br>
  </td>
</tr>
<tr>
  <td>
  Confirmación de E-Mail
  </td>
  <td > 
  	:   	
	</td>
  <td>
    <input type="text" name="confirmaEmail" size="45" maxlength="100" style="border-radius:5px" onBlur="validaEmail2()"><br>
  </td>
</tr>
<tr>
  <td>
  Teléfono Fijo
  </td>
  <td >
  	:    	
	</td>
  <td>
    <input type="text" name="telFijo" size="45" onKeyPress="return soloNumeros(event)" maxlength="11" style="border-radius:5px"><br>
  </td>
</tr>
<tr>
  <td>
  Teléfono Móvil *
  </td>
  <td > 
  	:   	
	</td>
  <td>
    <input type="text" name="telMovil" size="45" onKeyPress="return soloNumeros(event)" maxlength="11" style="border-radius:5px"><br>
  </td>
</tr>
<tr>
  <td>
  Nombre del Tutor *
  </td>
    <td > 
    	:   	
	</td>
  <td>
    <input type="text" name="nombreTutor" size="45" onkeypress="return soloLetras(event)" maxlength="100" style="border-radius:5px"><br>
  </td>
</tr>
<tr>
  <td>
  	Teléfono del Tutor *
  </td>
    <td >    	
    :
	</td>
  <td>
    <input type="text" name="telTutor" size="45" onKeyPress="return soloNumeros(event)" maxlength="11" style="border-radius:5px"><br>
  </td>
</tr>
<tr>
  <td style="vertical-align:text-top;">
  	Comentarios    
  </td>
    <td style="vertical-align:text-top;">
    :
	</td>
  <td colspan="2" width="400px">
  	<textarea name="comentarios" maxlength="5000" style="border-radius:5px; width:293" >
    </textarea>
  </td>
  <td >  	
	</td>
</tr>
<tr>
<td >   	
	</td>
	<td align="right" colspan="2">
		<input type="submit" value="ENVIAR" style="visibility:hidden;" align="right"><input type="image" src="boton_enviar.gif"
         style="width:70px"/>                
	</td>
  	<td>   	
	</td>
</tr>
</form>
</table>
</fieldset>
</td>
</tr>
</table>

<!--#INCLUDE file="include/desconexion.inc" -->