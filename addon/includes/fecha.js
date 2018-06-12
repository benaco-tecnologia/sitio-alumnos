<!-- This script and many more are available free online at -->
<!-- The JavaScript Source!! http://javascript.internet.com -->

<!-- Original:  Rob Patrick (rpatrick@mit.edu) -->


	var month_names = new Array("Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre");

function ocultar(){
	document.all.contenedor.style.visibility='hidden'
	document.all.contenedor1.style.visibility='hidden'
	document.all.contenedor2.style.visibility='hidden'
}
function setToday(cal) {
	var now   = new Date();
	var day   = now.getDate();
	var month = now.getMonth();
	var year  = now.getYear();
	if (year < 2000)    // Y2K Fix, Isaac Powell
		year = year + 1900; // http://onyx.idbsu.edu/~ipowell
	this.focusDay = day;
	document.all.month.value = month;
	document.all.year.value = year;
	
	if ((document.all.opcB.value==1)&&(document.all.opc.value!=1)){
		document.all.contenedor.style.visibility='hidden'
	} 
	if ((document.all.opcB.value==2)&&(document.all.opc.value!=2)){
		document.all.contenedor2.style.visibility='hidden'
	}
	if ((document.all.opcB.value==3)&&(document.all.opc.value!=3)){
		document.all.contenedor1.style.visibility='hidden'
	}
	displayCalendar(month, year, cal);
}

function isFourDigitYear(year) {
	if (year.length != 4) {
		alert ("Sorry, the year must be four-digits in length.");
		document.all.year.select();
		document.all.year.focus();
	} else { return true; }
}

function selectDate(cal) {
	var year  = document.all.year.value;
	if (isFourDigitYear(year)) {
		var day   = 0;
		var month = document.all.month.value;
		displayCalendar(month, year, cal);
    }
}

function setPreviousYear(cal) {
	var year  = document.all.year.value;
	if (isFourDigitYear(year)) {
		var day   = 0;
		var month = document.all.month.value;
		year--;
		document.all.year.value = year;
		displayCalendar(month, year, cal);
	   }
}

function setPreviousMonth(cal) {
	var year  = document.all.year.value;
	if (isFourDigitYear(year)) {
		var day   = 0;
		var month = document.all.month.value;
		if (month == 0) {
			month = 11;
			if (year > 1000) {
				year--;
				document.all.year.value = year;
			}
		} else { month--; }
		document.all.month.value = month;
		displayCalendar(month, year, cal);
	}
}

function setNextMonth(cal) {
	var year  = document.all.year.value;
	if (isFourDigitYear(year)) {
		var day   = 0;
		var month = document.all.month.value;
		if (month == 11) {
			month = 0;
			year++;
			document.all.year.value = year;
		} else { month++; }
		document.all.month.value = month;
		displayCalendar(month, year, cal);
   }
}

function setNextYear(cal) {
	var year = document.all.year.value;
	if (isFourDigitYear(year)) {
		var day = 0;
		var month = document.all.month.value;
		year++;
		document.all.year.value = year;
		displayCalendar(month, year, cal);
   }
}

function llenar(dia, mes, ano, cal){
	if (dia < 10)
		dia = "0"+dia;
	mes++;
	if (mes < 10)
		mes = "0"+mes;
	componente = eval(cal);
	componente.value = dia+"/"+mes+"/" +ano;
	document.all.contenedor.style.visibility='hidden'
	document.all.contenedor2.style.visibility='hidden'
	document.all.contenedor1.style.visibility='hidden'
}

function displayCalendar(month, year, cal) {
	cal = "" + cal + "";
	var sCalendario;

	var nuevo   = new Date();
	var dia   = nuevo.getDate();
	var mes = nuevo.getMonth();
	var año  = nuevo.getYear();

	month = parseInt(month);
	year = parseInt(year);
	var i = 0;
	var desplegar = 0;
	var days = getDaysInMonth(month+1,year);
	var firstOfMonth = new Date (year, month, 1);
	var startingPos = firstOfMonth.getDay();
	days += startingPos;
	sCalendario  = "<TABLE CALLSPACING=0 CELLPADDING=0 BGCOLOR='#e4e5ea'  style='BORDER-RIGHT: #df6c01 1px solid;BORDER-TOP: #df6c01 1px solid;BORDER-LEFT: #df6c01 1px solid;BORDER-BOTTOM: #df6c01 1px solid;'>";
	sCalendario += "<TR BGCOLOR='#c0c0c0'  style='FONT-WEIGHT: normal;FONT-SIZE: 10px;FONT-STYLE: normal;FONT-FAMILY: Arial;'><TH><input class='INPUTB' type='button' value='«' onclick='setPreviousMonth(\"" + cal + "\")'  style='CURSOR: hand'></TH><TH COLSPAN='7' style='BORDER-RIGHT: #000000 1px solid;BORDER-TOP: #000000 1px solid;BORDER-LEFT: #000000 1px solid;BORDER-BOTTOM: #000000 1px solid;' onclick='ocultar();'  style='CURSOR: hand'>" + month_names[month] + " " + year + "</TH><TH><input type='button' value='»' onclick='setNextMonth(\"" + cal +"\")'  style='CURSOR: hand'></TH></TR>";
	sCalendario += "<TR style='FONT-WEIGHT: normal;FONT-SIZE: 10px;FONT-STYLE: normal;FONT-FAMILY: Arial'><TH></TH><TH BGCOLOR='#df6c01'>Dom</TH><TH BGCOLOR='#df6c01'>Lun</TH><TH BGCOLOR='#df6c01'>Mar</TH><TH BGCOLOR='#df6c01'>Mie</TH><TH BGCOLOR='#df6c01'>Jue</TH><TH BGCOLOR='#df6c01'>Vie</TH><TH BGCOLOR='#df6c01'>Sab</TH></TR>";
	sCalendario += "<TR ALIGN=RIGHT  style='FONT-WEIGHT: normal;FONT-SIZE: 10px;FONT-STYLE: normal;FONT-FAMILY: Arial'>";
	for (i = 0; i < startingPos; i++) {
		if ( i%7 == 0 )
			sCalendario += "</TR><TR ALIGN='RIGHT'  style='FONT-WEIGHT: normal;FONT-SIZE: 10px;FONT-STYLE: normal;FONT-FAMILY: Arial'><TD></TD>";
			sCalendario += "<TD> </TD>";
	}

	for (i = startingPos; i < days; i++) {
		desplegar	= i-startingPos+1
		if ((desplegar == dia) && (month == mes) && (year == año)){
			if (i%7 == 0)
				sCalendario += "<TD><B>";
			else 
				sCalendario += "<TD BGCOLOR='#df6c01' onClick=llenar("+desplegar+","+month+","+year+",'"+cal+"')  style='CURSOR: hand'><B>";
		}
		else
		{
			if (i%7 == 0)
				sCalendario += "<TD><B>";
			else
				sCalendario += "<TD BGCOLOR='#C0C0C0' onClick=llenar("+desplegar+","+month+","+year+",'"+cal+"')  style='CURSOR: hand'><B>"; //BGCOLOR celdas
		}

		if ( i%7 == 0 )
			if ((desplegar == dia) && (month == mes) && (year == año))
				sCalendario += "</B></TD></TR><TR ALIGN='RIGHT'  style='FONT-WEIGHT: normal;FONT-SIZE: 10px;FONT-STYLE: normal;FONT-FAMILY: Arial'><TD></TD><TD BGCOLOR='#df6c01' onClick=llenar("+desplegar+","+month+","+year+",'"+cal+"') style='CURSOR: hand'><B>";
		 	else 
				sCalendario += "</B></TD></TR><TR ALIGN='RIGHT'  style='FONT-WEIGHT: normal;FONT-SIZE: 10px;FONT-STYLE: normal;FONT-FAMILY: Arial'><TD></TD><TD BGCOLOR='#C0C0C0' onClick=llenar("+desplegar+","+month+","+year+",'"+cal+"') style='CURSOR: hand'><B>";

			if (i-startingPos+1 < 10)
				sCalendario += "0";
		sCalendario += i-startingPos+1;
		sCalendario += "</B></TD>";
	}
	for (i=days; i<42; i++)  {
		if ( i%7 == 0 ) sCalendario += "</TR><TR ALIGN=RIGHT>";
			sCalendario += "<TD> </TD>";
	}

	if (document.all.opc.value == 1){
		document.all.opcB.value=1;
		document.getElementById("contenedor").innerHTML = sCalendario;
	}

	if (document.all.opc.value == 2){
		document.all.opcB.value=2;
		document.getElementById("contenedor2").innerHTML = sCalendario;
	}
	if (document.all.opc.value == 3){
		document.all.opcB.value=3;
		document.getElementById("contenedor1").innerHTML = sCalendario;
	}
}

function getDaysInMonth(month,year)  {
	var days;
	if (month==1 || month==3 || month==5 || month==7 || month==8 || month==10 || month==12)  days=31;
	else if (month==4 || month==6 || month==9 || month==11) days=30;
	else if (month==2)  {
		if (isLeapYear(year)) { days=29; }
	else { days=28; }
	}
	return (days);
}

function isLeapYear (Year) {
	if (((Year % 4)==0) && ((Year % 100)!=0) || ((Year % 400)==0)) {
		return (true);
	} else { return (false); }
}

function fechaactual(fecha_actual){
	var fec   = new Date();
	var d = fec.getDate();
	if (d < 10)
		d= "0" + d;
	var m = fec.getMonth();
	m++;
	if (m < 10)
		m= "0" + m;
	fecha_actual.value = d+"/"+m+"/" +fec.getYear();
}

<!-- Validar Fecha MM-AAAA -->

   function esDigito(sChr){ 
    var sCod = sChr.charCodeAt(0); 
    return ((sCod > 47) && (sCod < 58)); 
   } 

   function valMes(oTxt){ 
    var bOk = false; 
    var nMes = parseInt(oTxt.value.substr(4, 2), 10); 
    bOk = bOk || ((nMes >= 1) && (nMes <= 12)); 
    return bOk; 
   } 

   function valAno(oTxt){ 
    var bOk = true; 
    var nAno = oTxt.value.substr(0, 4); 
    bOk = bOk && (nAno.length == 4); 
    if (bOk){ 
     for (var i = 0; i < nAno.length; i++){ 
      bOk = bOk && esDigito(nAno.charAt(i)); 
     } 
    } 
    return bOk; 
   } 

   function valFecha(oTxt,i){ 
    var bOk = true; 
    if (oTxt.value != ""){ 
     bOk = bOk && (valAno(oTxt)); 
     bOk = bOk && (valMes(oTxt)); 
     if (!bOk){ 
      alert("Fecha inválida"); 
	  if (i == 0)
	  {
		  fechames(oTxt); 
   		  oTxt.select();
	  } else { 
		  oTxt.value = i;
  		  oTxt.select();
	  }
     } 
    } 
   } 

	function fechames(oTxt){
		var fec   = new Date();
		var m = fec.getMonth();
		m++;
		if (m < 10)
			m= "0" + m;
		oTxt.value = fec.getYear()+"" +m;
	}

<!-- Validar Fecha DD-MM-AAAA -->
   function esSeparador(oTxt){ 
    var bOk = false; 
    bOk = bOk || ((oTxt.value.substr(2, 1) == "/") && (oTxt.value.substr(5, 1) == "/")); 
    return bOk; 
   } 

   function valDiaB(oTxt){ 
    var bOk = false; 
    var nDia = parseInt(oTxt.value.substr(0, 2), 10);
    bOk = bOk || ((nDia >= 1) && (nDia <= getDaysInMonth(parseInt(oTxt.value.substr(3, 2), 10),parseInt(oTxt.value.substr(6, 4))))); 
    return bOk; 
   } 

   function valMesB(oTxt){ 
    var bOk = false; 
    var nMes = parseInt(oTxt.value.substr(3, 2), 10); 
    bOk = bOk || ((nMes >= 1) && (nMes <= 12)); 
    return bOk; 
   } 

   function valAnoB(oTxt,i){ 
    var bOk = true; 
    var nAno = oTxt.value.substr(6, 4); 
    bOk = bOk && (nAno.length == 4); 
    if (bOk){ 
     for (var i = 0; i < nAno.length; i++){ 
      bOk = bOk && esDigito(nAno.charAt(i)); 
     } 
    } 
    return bOk; 
   } 

   function valFechaB(oTxt,i){ 
    var bOk = true; 
    if (oTxt.value != ""){ 
     bOk = bOk && (esSeparador(oTxt)); 
     bOk = bOk && (valDiaB(oTxt)); 
	 bOk = bOk && (valMesB(oTxt)); 
	 bOk = bOk && (valAnoB(oTxt)); 
     if (!bOk){ 
      alert("Fecha inválida"); 
	  if (i == 0)
	  {
		  fechaactual(oTxt);
   		  oTxt.select();
	  } else { 
		  oTxt.value = i;
  		  oTxt.select();
	  }
     } 
    } 
   } 

<!-- Validar HORA -->

	function CheckTime(str){
		hora=str.value
		if (hora=='') {return}
		if (hora.length!=8) {alert("Introducir HH:MM:SS");	str.value = actual();return}
		a=hora.charAt(0) //<=2
		b=hora.charAt(1) //<4
		c=hora.charAt(2) //:
		d=hora.charAt(3) //<=5
		e=hora.charAt(5) //:
		f=hora.charAt(6) //<=5
		if ((a==2 && b>3) || (a>2)) {
			alert("La Hora es Incorrecta");
			str.value = actual()
			str.select();
			return
			}
		if (d>5) {
			alert("La Hora es Incorrecta");
			str.value = actual()
			str.select();
			return
			}
		if (f>5) {
			alert("La Hora es Incorrecta");
			str.value = actual()
			str.select();
			return
			}
		if (c!=':' || e!=':') {
			alert("Introducir la Hora en el siguinets formato HH:MM:SS");
			str.value = actual()
			str.select();
			return
			}
	}

	function actual(){

		marcacion = new Date() 
		Hora = marcacion.getHours() 
		Minutos = marcacion.getMinutes() 
		Segundos = marcacion.getSeconds() 

		if (Hora<=9) 
		Hora = "0" + Hora 

		if (Minutos<=9) 
		Minutos = "0" + Minutos 

		if (Segundos<=9) 
		Segundos = "0" + Segundos 

		return Hora +":" + Minutos + ":" + Segundos
	}
