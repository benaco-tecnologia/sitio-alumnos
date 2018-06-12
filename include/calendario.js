//globals

var month;      var day;        var year;   var firsttime;   
var delim = new Array(":","/","\\","-"," ",".");
var monthArray = new Array(0,31,29,31,30,31,30,31,31,30,31,30,31);

function vd(form, fieldName) 
{
   //init  
        dtString = eval("document." + form + "." + fieldName + ".value");
        // trim date string
        while ((dtString.charAt(0) == " ") && (dtString.length != 0))
                dtString = dtString.substring(1,dtString.length - 1)
        while ((dtString.charAt(dtString.length - 1) == " ") && (dtString.length != 0))
                dtString = dtString.substring(0,dtString.length - 1)
        //get date components
        i = 0; startPos = 0;    pos = 0;
        //get day
        do {
             pos = dtString.indexOf(delim[i], startPos);
             i++
        } while ((pos == -1) && (i < delim.length));
        if (pos == -1)return false;
        day  = parseInt(dtString.substring(startPos,pos),10);
        startPos = pos + 1;
        if ((day < 1) || (day > monthArray[month])) return false;
        //get month
        i = 0;
        do {
                pos = dtString.indexOf(delim[i], startPos);
                i++
        } while ((pos == -1) && (i < delim.length));

        if (pos == -1) return false;
        month  = parseInt(dtString.substring(startPos,pos),10);
        startPos = pos + 1;
        if ((month < 1) || (month > 12)) return false;

        //get year
        year = parseInt(dtString.substring(startPos,dtString.length),10);
        //check for leap year
        if ((month == 2) && (day == 29))
                if ((((year % 4) == 0) && ((year % 100) != 0)) == false)
                {       
                    return false;   
                }
        //if we've gotten this far, return true
	//RTF*
	if (year<1000 )
	{
		return false;
	}
        return true;
} // end function vd


function validateDate(form, dateFieldName,fieldLabel)
{

  if (eval("document." + form + "." + dateFieldName + ".value") =="")
  {
  } else
  {
    if (!vd(form, dateFieldName))
    {
      alert(fieldLabel + " formato de Fecha incorrecto");
      eval("document." + form + "." + dateFieldName + ".select();");
      eval("document." + form + "." + dateFieldName + ".focus();");   
    }
  }

}//end function validateDate


function getToday(){
        today = new Date();
        day = today.getDate();
        month = today.getMonth();
        month++;
        year = today.getYear();
        year = (year <= 100) ? 1900 + year : year;

}

function getMonth_and_Date(form,fieldName)
{
        dtString = eval("document." + form + "." + fieldName + ".value");
        // trim date string
        while ((dtString.charAt(0) == " ") && (dtString.length != 0))
                dtString = dtString.substring(1,dtString.length - 1)
        while ((dtString.charAt(dtString.length - 1) == " ") && (dtString.length != 0))
                dtString = dtString.substring(0,dtString.length - 1)
        //get date components
        i = 0; startPos = 0;    pos = 0;
        //get day
        do {
            pos = dtString.indexOf(delim[i], startPos);
            i++
        }while ((pos == -1) && (i < delim.length));
        if (pos == -1)
        {
                getToday();
                return;
        }
        day  = parseInt(dtString.substring(startPos,pos),10);
        startPos = pos + 1;
        if ((day < 1) || (day > monthArray[month])){
                getToday();
                return;
        }
        //get month
        i = 0;
        do {
            pos = dtString.indexOf(delim[i], startPos);
            i++
        }while ((pos == -1) && (i < delim.length));
        if (pos == -1){//there's no month
                getToday();     
                return;
        }
        month  = parseInt(dtString.substring(startPos,pos),10) - 1;
        startPos = pos + 1;
        if ((month < 0) || (month > 12)){ //no valid month
                getToday();
                return;
        }
        else
            month++;                        
        //get year
        year = parseInt(dtString.substring(startPos,dtString.length),11)
        year = (year <= 1000) ? 1900 + year : year;
		//alert(year);

}//getMonth_and_Date


function putDate(form,fieldName,value)
{
    eval("document." + form + "." + fieldName + ".value=" + value);
    eval("document." + form + "." + fieldName + ".select();");
    eval("document." + form + "." + fieldName + ".focus();");   

}

function gm(num) {
 var mydate = new Date();
 mydate.setDate(1);
 mydate.setMonth(num-1);
 var datestr = "" + mydate;
 return datestr.substring(4,7);
}


function gy(num) {
  var mydate = new Date();
  //alert(mydate);
  return (1900 + eval(mydate.getYear()) - 4 + num);
  
}


function ud(mon) {
  var i = mon.selectedIndex;
  if(mon.options[i].value == "2") {
    document.myform.day.options[30] = null;
    document.myform.day.options[29] = null;
    var j = document.myform.year.selectedIndex;
    var year = eval(document.myform.year.options[j].value);
    if ( ((year%400)==0) || (((year%100)!=0) && ((year%4)==0)) ) {
      if (document.myform.day.options[28] == null) {
        document.myform.day.options[28] = new Option("29");
        document.myform.day.options[28].value = "29";
      }
    } else {
      document.myform.day.options[28] = null;
    }
  }
  if(mon.options[i].value == "1" ||
     mon.options[i].value == "3" ||
     mon.options[i].value == "5" ||
     mon.options[i].value == "7" ||
     mon.options[i].value == "8" ||
     mon.options[i].value == "10" ||
     mon.options[i].value == "12")
  {
    if (document.myform.day.options[28] == null) {
      document.myform.day.options[28] = new Option("29");
      document.myform.day.options[28].value = "29";
    }
    if (document.myform.day.options[29] == null) {
      document.myform.day.options[29] = new Option("30");
      document.myform.day.options[29].value = "30";
    }
    if (document.myform.day.options[30] == null) {
      document.myform.day.options[30] = new Option("31");
      document.myform.day.options[30].value = "31";
    }
  }

  if(mon.options[i].value == "4" ||
     mon.options[i].value == "6" ||
     mon.options[i].value == "9" ||
     mon.options[i].value == "11")
  {
    if (document.myform.day.options[28] == null) {
      document.myform.day.options[28] = new Option("29");
      document.myform.day.options[28].value = "29";
    }
    if (document.myform.day.options[29] == null) {
      document.myform.day.options[29] = new Option("30");
      document.myform.day.options[29].value = "30";
    }
    document.myform.day.options[30] = null;
  }

  if (document.myform.day.selectedIndex == -1)
    document.myform.day.selectedIndex = 0;
}

function showdate() {
  var i = document.myform.month.selectedIndex;
  var j = document.myform.day.selectedIndex;
  var k = document.myform.year.selectedIndex;

  alert(document.myform.month.options[i].value + "/" +
        document.myform.day.options[j].value + "/" +
        document.myform.year.options[k].value)

}


function putcal(form, dateFieldName) {
  var version = navigator.appVersion;
  if (navigator.appVersion.indexOf("Mac") != -1) {
	 
    calwin = open("","calwin","width=280,height=230,resizable=no");
  } else {
	 
    calwin = open("","calwin","width=250,height=195,resizable=no");
  }
  firsttime = 1;
  
  calccal(calwin,form,dateFieldName);
}

function calccal(targetwin,form,dateFieldName) { 

  var monthname = new Array(12);
  
  monthname[0] = "Enero";
  monthname[1] = "Febrero";
  monthname[2] = "Marzo";
  monthname[3] = "Abril";
  monthname[4] = "Mayo";
  monthname[5] = "Junio";
  monthname[6] = "Julio";
  monthname[7] = "Agosto";
  monthname[8] = "Septiembre";
  monthname[9] = "Octubre";
  monthname[10] = "Noviembre";
  monthname[11] = "Diciembre";

  // Cambios R.C.
  if (firsttime == 1)
  {
    getMonth_and_Date(form, dateFieldName);
 //   alert(eval(form + "." + dateFieldName + ".value"));
 //  if (eval(form + "." + dateFieldName + ".value") == "") 
 //   { getToday(); }
 //   else { getMonth_and_Date(form, dateFieldName); }
    firsttime = 0;
  }
  var endday = calclastday(eval(month),eval(year));
  mystr = month + "/01/" + year;
  mydate = new Date(mystr);

 alert(mystr);

  firstday = mydate.getDay();
  var cnt = 0;
  var day = new Array(6);
  for (var i=0; i<6; i++)
    day[i] = new Array(7);
  for (var r=0; r<6; r++)
  {
    for (var c=0; c<7; c++)
    {
      if ((cnt==0) && (c!=firstday))
        continue;
      cnt++;
      day[r][c] = cnt;
      if (cnt==endday)
        break;
    }
    if (cnt==endday)
      break;
  }

  targetwin.document.open()
  targetwin.document.writeln("<HEAD><STYLE>A:link{text-decoration:none};A:hover{color:red}</STYLE></HEAD>")
  targetwin.document.writeln("<FORM><CENTER><TABLE><TR VALIGN=TOP>");

//FONDO DE VENTANA***************************************************
  targetwin.document.writeln("<BODY BACKGROUND = 'colora.jpg'>");
  //var prevyear = eval(year) - 1;
  var prevyear = eval(year);
//alert("prevyear")
//  targetwin.document.writeln("<TD><INPUT TYPE=BUTTON NAME=prevyearbutton VALUE='<<'"+
//   " onclick='opener.month = " + month + "; opener.year = " + prevyear + ";document.clear();opener.calccal(opener.calwin,opener.document." + form + ",\"" + dateFieldName + "\")'></TD>");
  targetwin.document.writeln("<TD><INPUT TYPE=BUTTON NAME=prevyearbutton VALUE='<<'"+
   " onclick='opener.month = " + month + ""+ "; opener.year = " + prevyear + ";document.clear();opener.calccal(opener.calwin, \"" + form + "\",\"" + dateFieldName + "\")'></TD>");
  var prevmonth = (month == 1) ? 12 : month - 1;
  var prevmonthyear = (month == 1) ? year - 1 : year;
 // alert (prevmonthyear);
//MESNUEVO
//  targetwin.document.writeln("<TD><INPUT TYPE=BUTTON NAME=prevmonthbutton VALUE='&nbsp;<&nbsp;'"+
//   " onclick='opener.month = " + prevmonth + "; opener.year = " + prevmonthyear + ";document.clear();opener.calccal(opener.calwin,opener.document." + form + ",\"" + dateFieldName + "\")'></TD>");
  targetwin.document.writeln("<TD><INPUT TYPE=BUTTON NAME=prevmonthbutton VALUE='&nbsp;<&nbsp;'"+
   " onclick='opener.month = " + prevmonth + ""+ "; opener.year = " + prevmonthyear + ";document.clear();opener.calccal(opener.calwin, \"" + form + "\",\"" + dateFieldName + "\")'></TD>");
  targetwin.document.writeln("<TD COLSPAN=3 ALIGN=CENTER>");
//MESANTIGUO
  var index = eval(month) - 1;
  targetwin.document.writeln("<B><FONT FACE = 'MS Sans Serif' SIZE = 1>" + monthname[index] + " " + year + "</FONT></B></TD>");
  var nextyear = eval(year) + 1;        
  var nextmonth = (month == 12) ? 1 : month + 1;
  var nextmonthyear = (month == 12) ? year + 1 : year;
//  targetwin.document.writeln("<TD><INPUT TYPE=BUTTON NAME=nextmonthbutton VALUE='&nbsp;>&nbsp;'"+
//   " onclick='opener.month = " + nextmonth + "; opener.year = " + nextmonthyear + ";document.clear();opener.calccal(opener.calwin,opener.document." + form + ",\"" + dateFieldName + "\")'></TD>");
//  targetwin.document.writeln("<TD><INPUT TYPE=BUTTON NAME=nextyearbutton VALUE='>>'"+
//  " onclick='opener.month = " + month + "; opener.year = " + nextyear + ";document.clear();opener.calccal(opener.calwin,opener.document." + form + ",\"" + dateFieldName + "\")'></TD>");
  targetwin.document.writeln("<TD><INPUT TYPE=BUTTON NAME=nextmonthbutton VALUE='&nbsp;>&nbsp;'"+
   " onclick='opener.month = " + nextmonth + "; opener.year = " + nextmonthyear + ";document.clear();opener.calccal(opener.calwin, \"" + form + "\",\"" + dateFieldName + "\")'></TD>");
  targetwin.document.writeln("<TD><INPUT TYPE=BUTTON NAME=nextyearbutton VALUE='>>'"+
   " onclick='opener.month = " + month + "; opener.year = " + nextyear + ";document.clear();opener.calccal(opener.calwin, \"" + form + "\",\"" + dateFieldName + "\")'></TD>");
  targetwin.document.writeln("</TR><TR>");
  targetwin.document.writeln("<TD><FONT FACE='MS Sans Serif' SIZE = 1>Dom</FONT></TD>");
  targetwin.document.writeln("<TD><FONT FACE='MS Sans Serif' SIZE = 1>Lun</FONT></TD>");
  targetwin.document.writeln("<TD><FONT FACE='MS Sans Serif' SIZE = 1>Mar</FONT></TD>");
  targetwin.document.writeln("<TD><FONT FACE='MS Sans Serif' SIZE = 1>Mie</FONT></TD>");
  targetwin.document.writeln("<TD><FONT FACE='MS Sans Serif' SIZE = 1>Jue</FONT></TD>");
  targetwin.document.writeln("<TD><FONT FACE='MS Sans Serif' SIZE = 1>Vie</FONT></TD>");
  targetwin.document.writeln("<TD><FONT FACE='MS Sans Serif' SIZE = 1>Sab</FONT></TD>");
  targetwin.document.writeln("</TR>");
  targetwin.document.writeln("<TR><TD COLSPAN=7><HR NOSHADE></TD></TR>");
  targetwin.document.writeln("<TR><TD COLSPAN=7><TABLE  WIDTH=\"100%\" BORDER=\"0\">");

  var selectedmonth = eval(month) - 1;

  var today = new Date();

  var thisyear = today.getYear() + 1900;

  var selectedyear = eval(year) - thisyear + 4;

  var conditionalpadder = "";

  for(r=0; r<6; r++)
  {
   targetwin.document.writeln("<TR>");

   for(c=0; c<7; c++)
   {
    targetwin.document.writeln("<TD><FONT FACE = 'MS Sans Serif' SIZE = 1>");

    if(day[r][c] != null) {

      if (day[r][c] < 10)

        conditionalpadder = "&nbsp;"

      else
        conditionalpadder = "";
//ESCRIBE LA FECHA EN EL CAMPO
      var dia = day[r][c]; 
      var mes = month;
      if(dia < 10){
        dia = '0'+ dia;
      }
      if(mes < 10){
        mes = '0'+mes;
      }
      targetwin.document.writeln("<a href=\"javascript:" + 
     "window.opener.document." + form + "." + dateFieldName + ".value = '" + dia + '/' + mes + '/' + year + "'" + 
     ";window.close();\">" + conditionalpadder + day[r][c] + conditionalpadder +  "</a>")
   }
   targetwin.document.writeln("</FONT></TD>");
  }
  targetwin.document.writeln("</TR>");
 }
//FIN de FORM***************************************************

  targetwin.document.writeln("</TABLE></TABLE></CENTER></FORM></BODY>");

  targetwin.document.close()
}

function calclastday(month,year) 
{
  if ((month==2) && ((year%4)==0))
    return 29;
  if ((month==2) && ((year%4)!=0))
    return 28;
  if ((month==1) || (month == 3) || (month == 5) || (month == 7) ||
      (month==8) || (month == 10) || (month ==12))
    return 31;
  return 30;
}

function calcnextmonth(month) 
{
  if (month=="12")
    return "1";
  else
    return (eval(month)+1);
}

function calcnextyear(month,year) 
{
  if (month=="12")
    return (eval(year)+1);
  else
    return (year);
}

function calcprevmonth(month) 
{
  if (month=="1")
    return "12";
  else
    return (eval(month)-1);
}

function calcprevyear(month,year) 
{
  if (month=="1")
    return (eval(year)-1);
  else
    return (year);
}
