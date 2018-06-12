<%  response.buffer = false 
    Response.Expires = -1
%>
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--<OBJECT RUNAT=server PROGID=ADODB.Recordset id=Inscribe name=Inscribe > </OBJECT>-->
<%
dim CodCarrera
CodCarrera = request("C")
id=request.Form("id")
AccionVar = request.form("Accion")
'CodCarrera ="1302C"
etapa=request.form("txtetapa")
fechai=request.Form("txtfechai")
fechaf=request.Form("txtfechaf")
Dim rs,conex,sql
set rs=Server.CreateObject ("ADODB.Recordset")
if AccionVar ="Grabar" then
			if id="" then id=0
				if CodCarrera="TODAS" then 
					call GrabarTodas				 
				else	
					sql="select * from sis_inscripcion"
					sql=sql & " where id="& id & ""
					set rs=Session("Conn").execute(sql)
					
						if not rs.eof then 
							sql="update sis_inscripcion set"
							sql=sql & " codcarr='" & codcarrera & "',"
							sql=sql & "fecha_inicio='"& fechai & "',"
							sql=sql & "fecha_final='"& fechaf & "',"
							sql=sql & "nombre='"& Etapa & "'"
							sql=sql & " where id="& id & ""
						else
							sql="insert into sis_inscripcion (codcarr,fecha_inicio,fecha_final,nombre) Values ("
							sql=sql & "'" & trim(codcarrera)& "',"
							sql=sql & "'" & trim(fechai)& "',"
							sql=sql & "'" & trim(fechaf)& "',"
							sql=sql & "'" & trim(etapa)& "')"
			 
						end if
						'response.write(sql)
						set rs=Session("Conn").execute(sql)
						call limpiar
				end if
			
end if
if CodCarrera="TODAS" then
	sql = "SELECT top 1 id, codcarr,fecha_inicio,fecha_final,nombre from sis_inscripcion "
	else
	sql = "SELECT * from sis_inscripcion where codcarr='" & Codcarrera & "'"
end if
		
set rs =Session("Conn").Execute(sql)

sub limpiar
 etapa=""
 fechai=""
 fechaf=""
 accionvar=""
end sub
sub GrabarTodas
dim StrSql,St
	dim RstSql,Rt
		dim Carrera, tipo 
set RstSql=Server.CreateObject ("ADODB.Recordset")
	StrSql="select * from sis_inscripcion "
		set RstSql=Session("Conn").Execute(StrSql)
			set Rt =Server.CreateObject ("ADODB.Recordset")

if not RstSql.eof then
	St="Select codcarr from mt_carrer order by codcarr"
		set Rt=Session("Conn").Execute(St) 	 	
			rt.movefirst
			while not rt.eof
				Carrera=rt(0)
				Tipo="UPDATE"
				call grabatabla(Carrera,fechai,fechaf,Etapa,Tipo)
				rt.movenext
			wend
	'Actualizar tabla
	else
	St="Select codcarr from mt_carrer order by codcarr"
		set Rt=Session("Conn").Execute(St) 	 	
			rt.movefirst
			while not rt.eof
				Carrera=rt(0)
				Tipo="INSERT"
				call grabatabla(Carrera,fechai,fechaf,Etapa,Tipo)
				rt.movenext
			wend
	'Insertar
	
end if
end sub
sub grabatabla (Carrera,fechai,fechaf,Etapa,Tipo)
dim Sql
dim MyRst
set MyRst=Server.CreateObject ("ADODB.Recordset")
	if tipo="INSERT" then
			sql="insert into sis_inscripcion (codcarr,fecha_inicio,fecha_final,nombre) Values ("
			sql=sql & "'" & trim(Carrera)& "',"
			sql=sql & "'" & trim(fechai)& "',"
			sql=sql & "'" & trim(fechaf)& "',"
			sql=sql & "'" & trim(etapa)& "')"
			set MyRst=Session("Conn").Execute(Sql)
	else
			sql="update sis_inscripcion set "
			sql=sql & "fecha_inicio='"& fechai & "',"
			sql=sql & "fecha_final='"& fechaf & "',"
			sql=sql & "nombre='"& Etapa & "'"
			sql=sql & "where Codcarr='"& Carrera & "'"
			'RESPONSE.WRITE(sql)
			'RESPONSE.End()
			set MyRst=Session("Conn").Execute(Sql)
	end if
end sub
%>
	
<html>
<head>
<title>Definición de Etapas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="TheForm" method="post" action="sis-inscripcion.asp">
<input type="hidden" name="Accion" value="<%=AccionVar%>">
<input type="hidden" name="C" value="<%=CodCarrera%>">
<input type="hidden" name="ID" value="<%=ID%>">
  <table width="70%" border="0" cellspacing="1" cellpadding="1">
    <tr bgcolor="4a5da1"> 
      <td width="37%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Nombre 
        : </font></b></td>
      <td width="17%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Fecha 
        Inicio :</font></b></td>
      <td width="19%"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Fecha 
        Final :</font></b></td>
      <td width="27%">&nbsp;</td>
    </tr>
    <tr  bgcolor="ffe5c3"> 
      <td width="37%"> 
        <input type="text" name="txtEtapa" value="<%=Etapa%>" size="30" maxlength="50">
      </td>
      <td width="17%"> 
        <input type="text" name="txtFechaI" value="<%=FechaI%>" size="12" maxlength="10">
      </td>
      <td width="19%"> 
        <input type="text" name="txtFechaF" value="<%=FechaF%>" size="12" maxlength="10">
      </td>
      <td width="27%"> 
        <div align="center">
          <input name="Submit" type="button" style="color: #FFFFFF;FONT-SIZE: 10px; height:20; font-weight: bold;width:60" onClick="javascript:grabar();" value="Guardar">
          <input name="Submit2" type="button" style="color: #FFFFFF;FONT-SIZE: 10px; height:20; font-weight: bold;width:60" onClick="javascript:Limpiar();" value="Limpiar">
        </div>
      </td>
    </tr>
  </table>
  <br>
  <table width="70%" border="0" cellspacing="1" cellpadding="1">
    <tr bgcolor="4a5da1"> 
      <td colspan="3">
        <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="2">Etapas</font></b></font></div>
      </td>
    </tr>
    <tr bgcolor="4a5da1"> 
      <td width="48%"> 
        <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>Nombre 
          : </b></font></div>
      </td>
      <td width="25%"> 
        <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF"><b>Fecha 
          Inicio :</b></font></div>
      </td>
      <td width="27%"> 
        <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Fecha 
          Final :</font></b></font></div>
      </td>
    </tr>
 <%if not rs.eof then 
	    do while not rs.eof%>
    <tr bgcolor="ffe5c3"> 
	
      <td width="48%"><a href="javascript:pasarvalores('<%=rs("Id")%>','<%=rs("Nombre")%>','<%=rs("fecha_inicio")%>','<%=rs("fecha_final")%>');" </a><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b><%=rs("nombre")%></b></font></td>
      <td width="25%"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b><%=rs("fecha_inicio")%></b></font></td>
      <td width="27%"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="1"><%=rs("fecha_final")%></font></b></font></td>
	</tr>
  	  <%rs.movenext
      loop
   else	  %>
      <tr bgcolor="ffe5c3"> 	
         
      <td width="48%" bordercolor="#CCCCCC">&nbsp;</td>
         <td width="25%">&nbsp;</td>
         <td width="27%">&nbsp;</td>
 	  </tr>
  <%end if%>
  </table>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
</form>
</body>
<script>
    function grabar() {
	       if (document.TheForm.C.value==""){
			   alert("Debe Seleccionar Carrera");
			   return;
			 }		
		   if (document.TheForm.txtEtapa.value==""){
			   alert("Debe Ingresar Nombre Etapa");
			   document.TheForm.elements['txtEtapa'].focus();
			   return;
			 }		  	 
			if (document.TheForm.txtFechaI.value==""){
			   alert("Debe Ingresar Fecha de Inicio");
			   document.TheForm.elements['txtFechaI'].focus();
			   return;			   
			  } 
			if (!ValidaFecha(document.TheForm.txtFechaI.value)==true) {
			    document.TheForm.elements['txtFechaI'].focus();
		        return;
	        } 
			   if (document.TheForm.txtFechaF.value==""){
			   alert("Debe Ingresar Fecha de Final");
			   document.TheForm.elements['txtFechaF'].focus();
			   return;
			 }   
			 if (!ValidaFecha(document.TheForm.txtFechaF.value)==true) {
			    document.TheForm.elements['txtFechaF'].focus();
		        return;
	         }   			 
	         document.TheForm.Accion.value ="Grabar"		
		  	 document.TheForm.submit();							 
			
	}
 
 function pasarvalores (ID,nombre,fecha_inicio,fecha_fin)
 {
 document.TheForm.ID.value=ID
 document.TheForm.txtEtapa.value=nombre
 document.TheForm.txtFechaI.value=fecha_inicio
 document.TheForm.txtFechaF.value=fecha_fin
 }
 
  function Limpiar()
 {
 document.TheForm.ID.value=""
 document.TheForm.txtEtapa.value=""
 document.TheForm.txtFechaI.value=""
 document.TheForm.txtFechaF.value=""
 }
 
 function ValidaFecha(obj){
	// función que valida la fecha 
	if (obj.length == 9){
  		alert("El formato de la fecha debe ser dd/mm/aaaa");
		//obj.focus();
		return;
	}
	Fecha = obj.replace("/","");
	Fecha = Fecha.replace("/","");
	if (noNumeric(Fecha)){
		alert("Debe ingresar una fecha válida");
		//obj.focus();
		return;
	}
	else{
		correcto=true;
		if(Fecha.length!=8) {
			alert("El formato de la fecha debe ser dd/mm/aaaa");
			//obj.focus();
			return;
		}
		var dia = Fecha.substr(0,2);
		var mes = Fecha.substr(2,2);
		var anyo =Fecha.substr(4,7);
		/*
		var dia = parseInt(dia);
		var mes = parseInt(mes);
		var anyo = parseInt(anyo)	   	
		*/
				
		if((anyo<1000) || (anyo>2100) ) {
			alert("La fecha especificada es incorrecta");
			//obj.focus();
			return;
		}
		if((mes>12) || (mes<1)) {
			alert("La fecha especificada es incorrecta");
			//obj.focus();
			return;
		}
		if((mes==4)||(mes==6)||(mes==9)||(mes==11)) {
			if((dia>30) || (dia<1)) {
				alert("La fecha especificada es incorrecta");
	//	obj.focus();
				return;
			}
		}
		if((mes==1)||(mes==3)||(mes==5)||(mes==7)||(mes==8)||(mes==10)||(mes==12)) {
			if((dia>31) || (dia<1)) {
				alert("La fecha especificada es incorrecta");
					//obj.focus();
					return;
			}
		}
		if((mes==2) && (!bisiesto(anyo))) {
			if((dia>28) || (dia<1)) {
				alert("El valor del dia de la fecha es incorrecto");
				//obj.focus();
				return;
			}
		} 
		else if((mes==2) && (bisiesto(anyo))) {
				if((dia>29) || (dia<1)) {
					alert("El valor del dia de la fecha es incorrecto");
//			obj.focus();
					return;
				}
		}		
		
	return true;
	
	}
}
function noNumeric(s)
{
	if (s == "") {
		return true
	} else {
		return isNaN(maskToNum(s));
	}
}
function maskToNum(s)
{
	s = cc(s, ".", "")
	s = cc(s, ",", ".")
	return s
}

function cc(s, c1, c2)
{
	snew = ""
	for (i=0; i < s.length; i++) {
		if (s.charAt(i) == c1) {
			snew = snew + c2;
		} else {
			snew = snew + s.charAt(i);
		}
	}
	return snew
}
function bisiesto(anyo) {  // comprueba si el año es bisiesto
	if ((anyo%4)==0) return true; 
	else return false;
}

</script>

</html>
<!--#INCLUDE file="include/desconexion.inc" -->
