<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<script language="JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<script type="text/javascript">
<!--


function Traspasa(Obj)
{
   i = Obj.name.substring(5,7);
   //Obj.name;
   //alert(i);
   for (j=0; j < document.form1.length; j++)
   { 
      //alert (document.TheForm.elements[j].name );
      if (document.form1.elements[j].name == "Texto" + i)
	  {
	     if (Obj.value == "OTRO")
		 {
	       document.form1.elements[j].value = "";
		   //document.getElementsByName(permuta).value = "";
		   document.form1.elements[j].disabled= false;
		 }else
		 {
	      document.form1.elements[j].value = Obj.value;
		  //document.getElementsByName(permuta).value = Obj.value;
		  document.form1.elements[j].disabled = true;
		 }
	  }
   }   
}


//function TraspasaCheckbox(Obj)
//{
  // i = Obj.name.substring(5,7);
   ////Obj.name;
   ////alert(i);
   //for (j=0; j < document.form1.length; j++)
   //{ 
     // //alert (document.TheForm.elements[j].name );
      //if (document.form1.elements[j].name == "Elimina" + i)
	  //{
	    // if (Obj.checked == "OTRO")
		 //{
	      // document.form1.elements[j].checked = "";
		   ////document.getElementsByName(permuta).value = "";
		   //document.form1.elements[j].disabled= false;
		 //}else
		 //{
	      //document.form1.elements[j].checked = Obj.checked;
		  ////document.getElementsByName(permuta).value = Obj.value;
		  //document.form1.elements[j].disabled = true;
		 //}
	//  }
   //}   
//}


function Rango(Obj)
 {
   i = Obj.name.substring(5,7);
  //alert(Obj.name);
   //alert(i);
   for (j=0; j < document.form1.length; j++)
   { 
      //alert (document.TheForm.elements[j].name );
      if (document.form1.elements[j].name == "Texto" + i){
 //rutina
 if (document.form1.elements[j].value!=""){ 
			    if ((isNaN(document.form1.elements[j].value))==true)
		           {	alert("Valor debe ser Númerico(Entero)");
			            document.form1.elements[j].select();
			            document.form1.elements[j].focus();
			            return;
		        }else {
				if (document.form1.elements[j].value < 1 || document.form.elements[j].value > 100) {
				    alert ("Valor fuera de rango");
					return;
				   }
				} 
	    }	
 //
	  }
   }   
}


function Mienvio(){
   document.form1.action="per-actualizar.asp";
 for (i=0;i<document.form1.length; i++)
 {
    document.form1.elements[i].disabled=false;
 }
//Alert ("test");
// document.TheForm.Action="Graba";
 document.form1.submit();
}

function alertaChecked(){ 
	var check = document.getElementById("elimina");
    if(check == null)
		alert('es nulo');
	else alert('No Es nulo');	
	
} 

function marca(Obj, valor)
			{
			   if (Obj.checked)
			    {Obj.value = "false"}
				 
			   else {Obj.value = valor}
			  //Obj.checked = !Obj.checked;
			  
			}
			
function ActualizaPermuta(obj,codram,codsec,permuta,tipo)
{	
var obj;
    
	 //alert(obj.checked);
	
	 //if (obj.checked){
	 //	alert("es true")
	  //       }else{
	 // 	alert("es false")
	 //}
		  
	if (obj.checked )  
		{ 
		//alert("1");
		window.location = "per-eliminar.asp?C="+codram+"&S="+codsec+"&P="+permuta+"&T="+1;
		}
	else
		{
		//alert("2");
		window.location = "per-eliminar.asp?C="+codram+"&S="+codsec+"&P="+permuta+"&T="+2;
		}
}



function jsValAsist(ID, valor){
	if (valor==''){
		valor='0';
	}
	
	if (isNaN(valor)){
		alert("Error\nDEBE INGRESAR SOLO NUMEROS");
		alert(ID, valor)
		document.getElementById(ID).value='0';
		document.getElementById(ID).focus();
		return false;
	}
	
	if (valor>100) {
		alert("Error\nNO PUEDE SER MAYOR QUE 100");
		document.getElementById(ID).value='0';
		document.getElementById(ID).focus();
		return false;
	}
			
	if (valor<0) {
		alert("Error\nNO PUEDE SER MENOR 1");
		document.getElementById(ID).value='0';
		document.getElementById(ID).focus();
		return false;
	}		
}


//-->
</script>

<link rel="stylesheet" href="css/tex-normales.css" type="text/css">

<OBJECT RUNAT=server PROGID=ADODB.Recordset id=RamosDebe name=RamosDebe></OBJECT> 
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=RamosSolici name=RamosSolici></OBJECT> 
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=RamosPuede name=RamosPuede></OBJECT> 
<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if
%>

<%
Dim strRamosDebe, strRamosPuede, Rut, CodCli
'Dim Codsecc
CodCli = Session("CodCli")
CodSede = Session("CodSede")
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if
id_proceso = Session("Proceso")
Ano = Session("Ano")
Periodo = Session("Periodo")

  strsql = " select b.id_proceso, b.tipo "
  strsql = strsql & " from ra_calendario_tr a, ra_proceso_ins b "
  strsql = strsql & " where a.id_proceso = b.id_proceso "
  strsql = strsql & " and getdate() between a.fecha_ini and a.fecha_fin "
  strsql = strsql & " and ano = '" & Ano & "' and periodo = '" & Periodo &"' "
  
 ' response.Write(strsql)
  'response.End()
  if bcl_ado(strsql, rstsql) then
        proceso = rstsql("id_proceso")
		tipo = rstsql("tipo")
		if tipo  = "SOLICITA" then
			Solicitado= "Solicitado"
			ID_PRO_SOLICITA= proceso
			 'response.Write(ID_PRO_SOLICITA)
		end if 
		
        if tipo = "MODIFICA" then
			Solicitado="Modificado"
			ID_PRO_MODIFICA= proceso
			'response.Write(ID_PRO_MODIFICA)
		end if 
		
        if tipo = "MODIFICA2" then
			Solicitado ="2° Modificacion"
			ID_PRO_MODIFICA2= proceso
			'response.Write(ID_PRO_MODIFICA2)
		end if
		
        if tipo = "EXCEPCION" then
			Solicitado ="Excepcion"
			ID_PRO_EXCEPCION= proceso
			'response.Write(ID_PRO_EXCEPCION)
		end if		
	    
  end if 
  'response.End() 
  
function BuscaRespuesta(codramo, codsecc)
Dim rstResp, sql
   sql= "Select coalesce(CodseccPermuta,0) as CodseccPermuta from ra_carga_log" 
   sql = sql & " where RAMOEQUIV = '" & codramo & "'"
   sql = sql & " and Ano = '" & Ano & "' "
   sql = sql & " and Periodo = '" & Periodo & "' " 
   sql = sql & " and codcli = '" & Codcli &"'"'
   sql = sql & " and inscrito = 'S' "
   sql = sql & " and codsecc ='"&codsecc&"' "
   sql = sql & " and id_proceso= '"& Proceso &"'"
   
   'response.write("sql")
   'response.end()
   if bcl_ado(Sql,rstResp) then
	if ValNulo(rstResp("CodSeccPermuta"),Num_)= 0 then
		BuscaRespuesta = "" 
	else 
		BuscaRespuesta = rstResp("CodSeccPermuta")
	end if 

  end if	  
end function 	

' desde la tabla mt_parame obtener el año y periodo
strRamosDebe = " SELECT b.codramo, b.nombre, a.inscrito, a.codsecc, a.Ramoequiv, c.nombre as NombreReal, tipo.tipocurso,"
strRamosDebe = strRamosDebe & " tipo.fechacert1, tipo.fechacert2, tipo.fechacert3, tipo.fechaexamen, tipo.fechaexamener, a.prioridadsol"
strRamosDebe = strRamosDebe & " FROM ra_carga a, ra_ramo b, ra_ramo c , ra_seccio tipo" 
strRamosDebe = strRamosDebe & " WHERE a.ramoequiv = b.codramo and " 
strRamosDebe = strRamosDebe & " a.ramoequiv *= c.codramo and " 
strRamosDebe = strRamosDebe & " a.preinscrito = 'S' and " 
strRamosDebe = strRamosDebe & " a.inscrito = 'S' and " 
strRamosDebe = strRamosDebe & " TIPO.TIPOCURSO = 'T' and "
strRamosDebe = strRamosDebe & " a.codcli ='" & CodCli & "' and " 
strRamosDebe = strRamosDebe & " a.ano = '" & ano & "' and " 
strRamosDebe = strRamosDebe & " a.periodo = '" & periodo & "'" 
strRamosDebe = strRamosDebe & " and b.codramo=tipo.codramo and a.codsecc=tipo.codsecc "
strRamosDebe = strRamosDebe & " and tipo.ano='"& ano &"'"
strRamosDebe = strRamosDebe & " and tipo.periodo='" & periodo & "'"

strRamosDebe = strRamosDebe & " Union "

strRamosDebe = strRamosDebe & " SELECT b.codramo, b.nombre, a.inscrito, a.codsecc, a.Ramoequiv, c.nombre as NombreReal,tipo.tipocurso, " 
strRamosDebe = strRamosDebe & " tipo.fechacert1, tipo.fechacert2, tipo.fechacert3, tipo.fechaexamen, tipo.fechaexamener, a.prioridadsol"
strRamosDebe = strRamosDebe & " FROM ra_cargaactividad a, ra_ramo b, ra_ramo c ,ra_seccio tipo " 
strRamosDebe = strRamosDebe & " WHERE a.ramoequiv = b.codramo and " 
strRamosDebe = strRamosDebe & " a.ramoequiv *= c.codramo and " 
strRamosDebe = strRamosDebe & " a.preinscrito = 'S' and " 
strRamosDebe = strRamosDebe & " a.inscrito = 'S' and "
strRamosDebe = strRamosDebe & " TIPO.TIPOCURSO = 'T' and " 
strRamosDebe = strRamosDebe & " a.codcli ='" & CodCli & "' and " 
strRamosDebe = strRamosDebe & " a.ano = '" & ano & "' and " 
strRamosDebe = strRamosDebe & " a.periodo = '" & periodo & "'" 
strRamosDebe = strRamosDebe & " and b.codramo=tipo.codramo and a.codsecc=tipo.codsecc "
strRamosDebe = strRamosDebe & " and tipo.ano='"& ano &"'"
strRamosDebe = strRamosDebe & " and tipo.periodo='" & periodo & "'"

'response.write(strRamosDebe)
'response.End()			     
				 
				' Order By a.prioridad "
'agregar el año y periodo a esta query
RamosDebe.Open strRamosDebe, Session("Conn")

strRamosSolici = "SELECT b.codramo, b.nombre, a.estado, a.codsecc, a.Ramoequiv, c.nombre as NombreReal, a.CodSeccOri " & _
                " FROM ra_solici a, ra_ramo b, ra_ramo c " & _
                " WHERE a.ramoequiv = b.codramo and " & _
                " a.ramoequiv *= c.codramo and " & _
                " a.codcli ='" & CodCli & "' and " & _
                " a.ano = '" & ano & "' and " & _
                " a.PerInscrip = " & SESSION("PER_ID") & " and " & _
                " a.periodo = '" & periodo & "' "

'agregar el año y periodo a esta query
'RESPONSE.WRITE(strRamosSolici)
'RESPONSE.End()

RamosSolici.Open strRamosSolici, Session("Conn")
dim paso 
paso=0
do while not RamosSolici.eof and paso=0
    if RamosSolici("codsecc")<>RamosSolici("CodSeccori") then
	 paso=1	  
    end if
   RamosSolici.movenext
loop
RamosSolici.close


strRamosSolici = "SELECT b.codramo, b.nombre, a.estado, a.codsecc, a.Ramoequiv, c.nombre as NombreReal, CodSeccOri " & _
                " FROM ra_solici a, ra_ramo b, ra_ramo c " & _
                " WHERE a.ramoequiv = b.codramo and " & _
                " a.ramoequiv *= c.codramo and " & _
                " a.codcli ='" & CodCli & "' and " & _
                " a.ano = '" & ano & "' and " & _
                " a.PerInscrip = " & SESSION("PER_ID") & " and " & _
                " a.periodo = '" & periodo & "' "
'agregar el año y periodo a esta query
'RESPONSE.WRITE(strRamosSolici)
'RESPONSE.End
RamosSolici.Open strRamosSolici, Session("Conn")

%>
<Script Language="JavaScript">
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
</Script><html>
<head> 
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html;">
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"> 

<%' if not InscritoNormal() then %>
  <script>
  //   alert("No ha inscrito aún vía solicitud Normal ...");
  </script>   
<%'end if%>
<table border="0" cellpadding="0" cellspacing="0"  align="center">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right" ><% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td width="825" valign="top">
			<% CargarTop1()%><% SubMenu()%>	
				
				  <form name="form1" method="post" action="per-actualizar.asp">
					<input type = "hidden" name = "Ramos" value = "">
					<table border="0" cellpadding="0" cellspacing="15" width="730">
					  <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
					  <tr valign="top" > 
						<td ><img name="fder_r2_c1" src="Imagenes/titulos/T-ver_cursos_insc.gif" width="400" height="38" border="0"></tr>
						<tr valign="top" > 
						<td  align="right"><a href="javascript:imprime()"><img src="imag/f-der/impresion.gif" width="36" height="36" border="0" align="right"></a></tr>
					  <tr valign="top">
					  	<td height="1" colspan=3><span class="Tit-celdas"><b class="Tit-celdas"></b></span><span style="font-size: 13px; font-weight: bold; font-family: Verdana, Arial, Helvetica, sans-serif;"><b class="Tit-celdas"><font face="Verdana, Arial, Helvetica, sans-serif">Resultado 
						  de Inscripción de : &nbsp;&nbsp;&nbsp;</font></b></span><span class="Tit-celdas"><b class="Tit-celdas"></b><font class="Tit-celdas"><%=GetNombrealumno(Codcli)%> </font></span></td>
				      </tr>
					  <tr valign="bottom"> 
						<td width="688" class="text-normal-celdas"><font class="text-normal-celdas">Fecha:&nbsp;&nbsp;<%=date()%></font></span></b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="tex-totales-celda"><font class="text-normal-celdas">Hora:&nbsp;&nbsp;<%=time()%></font></span></b></span></td>
					
						<tr valign="top">
					  	<td height="1" colspan=3 align="right"><span class="Tit-celdas"><b class="Tit-celdas"></b></span><a target="_top" href="modifica-toma-de-ramos.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','Imagenes/botones/volver-on.gif',1)"><img src="Imagenes/botones/volver-of.gif" width="178" height="45" name="Image1" border="0"></a></td>
				      </tr>
					  <tr valign="top"> 
						<td colspan="3" height="80"> 
						  <table width="741" cellspacing="1" cellpadding="0" height="59" border="0" bordercolor="#FFFFFF">
							<td height="20" colspan="8"> <div align="left"> 
								<p class="text-normal-celdas"><font face="Verdana, Arial, Helvetica, sans-serif" size="2" color="#000000">Ud. 
								  Tiene inscritas las siguientes asignaturas:</font></p>
							  </div></td>
							</tr>
							<tr background="imagenes/fdo-cabecera-cel22.jpg"> 
							  <td width="68" height="30"  background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center" ><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">C&oacute;digo</font></b></font></div></td>
							  <td width="312" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Asignatura</font></b></font></div></td>
							  <td width="65" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">S</font><font size="1">ecci&oacute;n</font></b></font></div> <div align="center"></div></td>
							  <td width="45" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Tipo</font></b></font></div></td>
							  <td width="112" height="30" background="imagenes/fdo-cabecera-cel22.jpg" align="center"><span style="font-weight: bold"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Permutar</font></span></td>
							  <td width="65" height="30" background="imagenes/fdo-cabecera-cel22.jpg" align="center"><span style="font-weight: bold"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">Eliminar</font></span></td>
						    </tr>
							<% 
							TabInx = 1
						  If RamosDebe.Eof Then
						  %>
							<%
						  Else
						   While Not RamosDebe.Eof
						      'response.write(RamosDebe("CodSecc"))
							  'response.End()
							 if Ucase(RamosDebe("inscrito")) = "S" then
							 'if Ucase(RamosDebe("Preinscrito")) = "S" then
							   'CodSecc = RamosDebe("CodSecc")
						      'response.write(CodSecc)
							  'response.End()						
							   Horario = GetHor2(RamosDebe("RamoEquiv"), CodSede, RamosDebe("CodSecc"), Ano, Periodo)
							 else
							   CodSecc = ""
							   Horario = ""
							 end if
						   %>
						   	<tr bgcolor="#ffe3e3" height="25"> 
							  <td width="68" height="30" align="center" ><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("CodRamo")%></font></td>
							  <td width="312" height="30" align="left" bgcolor="#ffe3e3"><font face="Arial, Helvetica, sans-serif" size="1">&nbsp;&nbsp;&nbsp;<%=RamosDebe("Nombre")%></font></td>
							  <td height="30" align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=RamosDebe("CodSecc")%></font></td>
							  <td width="45" height="30" align="center"><font size="1" face="Arial, Helvetica, sans-serif"> 
								<%if RamosDebe("Tipocurso")="T"  then
																																 response.write("Teorico")
																															  end if	 
																															  if  RamosDebe("Tipocurso")="L" then
																																  response.write("Laboratorio")	
																															  end if
																															   if RamosDebe("Tipocurso")="P" then
																																  response.write("Practico")
																															   end if
																															   
																															   
							%>
                            
                   <%
									Str1 = "Select distinct coalesce(lo.CodseccPermuta,0) as CodseccPermuta, coalesce(LO.TIPOSOLICITUD,'N') as tiposolicitud"
									Str1 = Str1 & " from ra_seccio e, ra_profes p, ra_ramo r, ra_carga_log lo "
									Str1 = Str1 & " where e.CodRamo = '" & RamosDebe("Codramo") & "' and e.CodSede = '" & CodSede & "' " 
									Str1 = Str1 & " and e.Ano = " & ano & " And e.Periodo = " & Periodo & ""
									Str1 = Str1 & " and e.CodProf *= p.CodProf "
									If FiltraJornada = "S" Then
									   Str1 = Str1 & " and e.Jornada = '" & Session("Jornada") & "'"
									end if                  
									Str1 = Str1 & " and lo.codseccpermuta is not null" 
									Str1 = Str1 & " and e.codramo = r.codramo" 
									Str1 = Str1 & " and e.codramo = lo.codramo "
									Str1 = Str1 & " and e.codsecc not in ( select lo.CODSECC from RA_CARGA_LOG lo "
									Str1 = Str1 & "where e.CODRAMO = lo.RAMOEQUIV and lo.CODCLI = '" & trim(session("codcli")) & "' ) "  
									Str1 = Str1 & " and e.codramo not in (select n.ramoequiv from ra_nota n "
									Str1 = Str1 & " where n.codcli='" & trim(session("codcli")) & "' and n.estado in ('A','E','I'))"
									Str1 = Str1 & " Order by lo.CodSeccPermuta "
									
									'response.Write(str1)
									'response.End()
				   %>         
				 </font></td>
				<td width="112" height="30" align="center">
				 <%
				Micontrol="Combo" & int(control)
                Mitexto="Texto" & int(control)
                Mivariable="P"  & int(control)
				%>
                <select name="<%=Micontrol%>" id="micontrol" size="1" style="width: 60px" onChange="javascript:Traspasa(this)" onClick="javascript:Traspasa(this)" >
				<%  
				  dim str 
				  dim rst 
				  dim Opcion
						
						Str = "Select distinct e.CodSecc "
						Str = Str & " from ra_seccio e, ra_profes p, ra_ramo r, ra_carga_log lo "
						Str = Str & " where e.CodRamo = '" & RamosDebe("Codramo") & "' and e.CodSede = '" & CodSede & "' " 
						Str = Str & " and e.Ano = " & ano & " And e.Periodo = " & Periodo & ""
						Str = Str & " and e.CodProf *= p.CodProf "
						If FiltraJornada = "S" Then
						   Str = Str & " and e.Jornada = '" & Session("Jornada") & "'"
						end if                  
						Str = Str & " and e.codramo = r.codramo" 
						Str = Str & " and e.codramo = lo.codramo "
						Str = Str & " and e.codsecc not in ( select lo.CODSECC from RA_CARGA_LOG lo "
						Str = Str & " where e.CODRAMO = lo.RAMOEQUIV and lo.CODCLI = '" & trim(session("codcli")) & "' ) "  
						Str = Str & " and e.codramo not in (select n.ramoequiv from ra_nota n "
						Str = Str & " where n.codcli='" & trim(session("codcli")) & "' and n.estado in ('A','E','I'))"
						Str = Str & " Order by e.CodSecc "
						
						SeccionPermuta = 0	
						TIPOSOLICITUD = "N"
						 if bcl_ado(Str,rst) then
							
							
											
							Micontrol= Micontrol & "<option value=>"														
							while not rst.EOF
									'id_Proceso= rst("id_proceso")
									SeccionVariable= rst("codsecc")	
									'if cdbl(RamosDebe("CodSecc"))=cdbl(rst("codsecc")) then
									if cdbl(RamosDebe("CodSecc"))=cdbl(SeccionVariable) then
										sel="selected "
									else					
										sel="" 
									end if
				
									Str1 = "Select distinct coalesce(lo.CodseccPermuta,0) as CodseccPermuta, coalesce(LO.TIPOSOLICITUD,'N') as tiposolicitud"
									Str1 = Str1 & " from ra_seccio e, ra_profes p, ra_ramo r, ra_carga_log lo "
									Str1 = Str1 & " where e.CodRamo = '" & RamosDebe("Codramo") & "' and e.CodSede = '" & CodSede & "' " 
									Str1 = Str1 & " and e.Ano = " & ano & " And e.Periodo = " & Periodo & ""
									Str1 = Str1 & " and e.CodProf *= p.CodProf "
									If FiltraJornada = "S" Then
									   Str1 = Str1 & " and e.Jornada = '" & Session("Jornada") & "'"
									end if                  
									'Str1 = Str1 & " and lo.codseccpermuta is not null" 
									Str1 = Str1 & " and e.codramo = r.codramo" 
									Str1 = Str1 & " and e.codramo = lo.codramo "
									Str1 = Str1 & " and e.codsecc not in ( select lo.CODSECC from RA_CARGA_LOG lo "
									Str1 = Str1 & "where e.CODRAMO = lo.RAMOEQUIV and lo.CODCLI = '" & trim(session("codcli")) & "' ) "  
									Str1 = Str1 & " and e.codramo not in (select n.ramoequiv from ra_nota n "
									Str1 = Str1 & " where n.codcli='" & trim(session("codcli")) & "' and n.estado in ('A','E','I'))"
									Str1 = Str1 & " Order by lo.CodSeccPermuta "
									

									if bcl_ado(Str1,rst1) then
									    
										IF trim(Rst1("CodSeccPermuta")) = 0 OR trim(Rst1("CodSeccPermuta")) = "" THEN 
											SeccionPermuta = ""	
										END IF 
										
										if trim(Rst1("CodSeccPermuta")) = 0 AND trim(Rst1("tiposolicitud")) ="N" then 
										    SeccionPermuta = ""											  
										else 
								      		SeccionPermuta = Rst1("CodSeccPermuta")
										end if 

									  IF UCASE(Rst1("tiposolicitud"))="ELIMINAR" THEN 
										  TIPOSOLICITUD = "S"
									  ELSE
										  TIPOSOLICITUD = "N"
									  END IF 
									
									end if 
								
								Micontrol= Micontrol & "<option value=""" & rst("codsecc") & """ " & sel & " >"
								Micontrol=	Micontrol & rst("codsecc") & "</option>" & vbcrlf 
															
								rst.MoveNext 
							wend
					
					end if%>
			  <%=Micontrol%>
                                   
			</select>
        <input type="text" name="<%=Mitexto%>" maxlength="1" size="1" disabled="false" onChange="Javascript:Rango(this) " 
        value=<%=BuscaRespuesta(RamosDebe("Codramo"),RamosDebe("Codsecc"))%> >      </td>
        
        
       <td height="30" align="center"><% if TIPOSOLICITUD = "S" then %>
							  </font><font face="Verdana" size="1">
								<input type="checkbox" name="elimina" checked="TRUE" value="1" OnClick="Javascript:ActualizaPermuta(this,'<%=RamosDebe("CodRamo")%>','<%=RamosDebe("Codsecc")%>','<%=SeccionPermuta%>','1')">
							  <% else %>
								<input type="checkbox" name="elimina" value="0" OnClick="Javascript:ActualizaPermuta(this,'<%=RamosDebe("CodRamo")%>','<%=RamosDebe("Codsecc")%>','<%=SeccionPermuta%>','2')">
							  <% end if %></span></td>
						   </tr>
						   <%
							I = I + 1
							control=int(control) + 1
							 RamosDebe.MoveNext
						   Wend
						 End If %>
						  </table>
						  <br>
                          <table>
                           <tr> 
      						<td colspan="2"><a href="javascript:Mienvio()" onMouseOver="MM_swapImage('Image2','','Imagenes/botones/grabar-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="Imagenes/botones/grabar-of.gif" width="178" height="45" name="Image2" border="0"></a></td>
   						 </tr>
                         </table>
						  <%
X=REQUEST("X")
IF X=1 THEN
	 
	 %>
	
	 <script>
	 alert("Esta sección no corresponde a la asignatura teórica seleccionada")
	 </script>
	
	 <%
END IF


'Dim strRamosDebe, strRamosPuede, Rut, CodCli
Dim RstHorario

CodCli = Session("CodCli")
IF CodCli = "" then
   'Response.Redirect "pagina de login"
   'Response.Write( "Me voy : " + CodCli + " " + session("CodCli") )
end if
Ano = Session("Ano")
Periodo = Session("Periodo")
CodSede = Session("CodSede")
P = REQUEST("P")
P = "S"
M = Request("M")

Dim MtrHorario(20, 7)
Dim MtrHorSt(20, 7)
Dim MtrHorRamo(20, 7)

If P = "S" then
  ConPreInscrito = 1
else
  ConPreInscrito = 0
end if
strHorario =""
T = ucase(REQUEST("T"))


' Analizar según el procedimiento TopeHorario
if ConPreInscrito = 1 then

  strHorario = " SELECT h.codramo, h.codsecc, h.codsal, h.dia, h.codmod, c.Inscrito, c.Preinscrito, c.codramo as ramomalla ,m.orden"
  strHorario = strHorario & " FROM ra_horari h, ra_carga_log c, ra_modulo m  "
  strHorario = strHorario & " WHERE c.codcli ='" & CodCli & "' and " 
  strHorario = strHorario & " c.ramoequiv_i = h.codramo and " 
  strHorario = strHorario & " c.codsecc_i = h.codsecc and  " 
  strHorario = strHorario & " c.ano = h.ano and  " 
  strHorario = strHorario & " h.codmod = m.codmod and  " 
  strHorario = strHorario & " c.periodo = h.periodo and  "
  strHorario = strHorario & " h.ano = '" & ano & "' and "
  strHorario = strHorario & " h.periodo = '" & periodo & "' and "
  strHorario = strHorario & " h.CodSede = '" & Codsede & "' and "
  strHorario = strHorario & " c.preinscrito = 'S' "
			   
  strHorario = strHorario & " UNION SELECT h.codramo, h.codsecc, h.codsal, h.dia, h.codmod, c.Inscrito, c.Preinscrito, c.codramo as ramomalla, m.orden "
  strHorario = strHorario & " FROM ra_horari h, ra_cargaactividad c, ra_modulo m  "
  strHorario = strHorario & " WHERE c.codcli ='" & CodCli & "' and "
  strHorario = strHorario & " c.ramoequiv = h.codramo and "
  strHorario = strHorario & " c.codsecc_i = h.codsecc and  "
  strHorario = strHorario & " c.ano = h.ano and  "
  strHorario = strHorario & " h.codmod = m.codmod and  " 
  strHorario = strHorario & " c.periodo = h.periodo and  "
  strHorario = strHorario & " h.ano = '" & ano & "' and "
  strHorario = strHorario & " h.periodo = '" & periodo & "' and "
  strHorario = strHorario & " h.CodSede = '" & Codsede & "' "
  strHorario = strHorario & " order by h.dia, h.codmod"
  
else
  strHorario = " SELECT h.codramo, h.codsecc, h.codsal, h.dia, h.codmod, c.Inscrito, c.Preinscrito, c.codramo as ramomalla , m.orden "
  strHorario = strHorario & " FROM ra_horari h, ra_carga_log c, ra_modulo m "
  strHorario = strHorario & " WHERE c.codcli ='" & CodCli & "' and "
  strHorario = strHorario & " c.ramoequiv = h.codramo and "
  strHorario = strHorario & " c.codsecc = h.codsecc and  "
  strHorario = strHorario & " c.ano = h.ano and  "
  strHorario = strHorario & " h.codmod = m.codmod and  " 
  strHorario = strHorario & " c.periodo = h.periodo and  "
  strHorario = strHorario & " h.ano = '" & ano & "' and "
  strHorario = strHorario & " c.inscrito = 'S' and "
  strHorario = strHorario & " h.periodo = '" & periodo & "' and "
  strHorario = strHorario & " h.CodSede = '" & Codsede & "' "
			   
  strHorario = strHorario & " UNION SELECT h.codramo, h.codsecc, h.codsal, h.dia, h.codmod, c.Inscrito, c.Preinscrito, c.codramo as ramomalla, m.orden "
  strHorario = strHorario & " FROM ra_horari h, ra_cargaactividad c , ra_modulo m "
  strHorario = strHorario & " WHERE c.codcli ='" & CodCli & "' and "
  strHorario = strHorario & " c.ramoequiv = h.codramo and "
  strHorario = strHorario & " c.codsecc = h.codsecc and  "
  strHorario = strHorario & " c.ano = h.ano and  "
  strHorario = strHorario & " h.codmod = m.codmod and  " 
  strHorario = strHorario & " c.periodo = h.periodo and  "
  strHorario = strHorario & " h.ano = '" & ano & "' and "
  strHorario = strHorario & " c.inscrito = 'S' and "
  strHorario = strHorario & " h.periodo = '" & periodo & "' and "
  strHorario = strHorario & " h.CodSede = '" & Codsede & "' "
  strHorario = strHorario & " order by h.dia, h.codmod"
end if

'RstHorario.Open strHorario, Session("Conn")
'Response.Write(StrHorario)
'Response.End

set RstHorario = Session("Conn").execute(strHorario)

do while not RstHorario.eof
  j = GetDia(RstHorario("dia"))
  i = RstHorario("orden")
  if ConPreinscrito then
    if RstHorario("Preinscrito") = "S" then
	'if RstHorario("inscrito") = "S" then
       'Response.Write("Ramo = " & RstHorario("CODRAMO") & " " & rstHorario("RamoMalla") & "<br>" ) 
       MtrHorario(i, j) = RstHorario("CODRAMO") & "/" & RstHorario("CodSecc") & " " & RstHorario("CodSal")
       if Trim(MtrHorSt(i,j)) = "" or Trim(MtrHorSt(i,j)) = "T" then
         MtrHorSt(i,j) = "P"
       else
         MtrHorSt(i,j) = "E"
       end if  
       MtrHorRamo(i, j) = RstHorario("RamoMalla")
         
       'MtrHorario(i, j, 2) = "P"
    else
      'if RstHorario("Inscrito") = "S" then
      '   MtrHorario(i, j) = RstHorario("CODRAMO") & "/" & RstHorario("CodSecc") & " " & RstHorario("CodSal")
      '   if Trim(MtrHorSt(i,j)) <> "" then
      '     MtrHorSt(i,j) = "I"
      '   else
      '     MtrHorSt(i,j) = "E"
      '   end if  
         'MtrHorario(i, j, 2) = "I"
      'end if
    end if
  else
    if trim(RstHorario("PreInscrito")) = "S" then
	'if trim(RstHorario("Inscrito")) = "S" then
       MtrHorario(i, j ) = RstHorario("CODRAMO") & "/" & RstHorario("CodSecc") & " " & RstHorario("CodSal")
       if Trim(MtrHorSt(i,j)) = "" or Trim(MtrHorSt(i,j)) = "T" then
         MtrHorSt(i,j) = "I"
       else
         MtrHorSt(i,j) = "E"
       end if  
       MtrHorRamo(i, j) = RstHorario("RamoMalla")
       'MtrHorario(i, j, 2) = "I"
    end if
  end if
  
  'Response.Write(MtrHorario(i, j))
  RstHorario.movenext
loop

'RstHorario.close()



%>
						  <table align="center">
						  	<tr>
							 
							</tr>
						  </table>
						  
						  <p></p></td>
					  </tr><tr valign="top"> 
        <td height="1" colspan="2" class="Tit-celdas"><b><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Horario 
          de asignaturas:&nbsp;&nbsp;&nbsp;</font></b><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><%=GetNombrealumno(Codcli)%></font></td> 
        </tr>
      <tr valign="top" > 
        <td colspan="2" height="2"> 
        <table width="741" cellspacing="1" cellpadding="0" height="30" border="0" bordercolor="#FFFFFF" >
          <tr background="Imagenes/fdo-cabecera-cel22.jpg"> 
              <td width="24" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Mod</font></b></font></div>            </td>
              <td width="98" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Hora</font></b></font></div>            </td>
              <td width="94" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Lunes</font></b></font></div>            </td>
              <td width="94" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Martes</font></b></font></div>            </td>
              <td width="94" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Mi&eacute;rcoles</font></b></font></div>            </td>
              <td width="94" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Jueves</font></b></font></div>            </td>
              <td width="89" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Viernes</font></b></font></div>            </td>
              <td width="110" height="30" background="imagenes/fdo-cabecera-cel22.jpg"> 
                <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">S&aacute;bado</font></b></font></div>            </td>
          </tr>
			<%
			dim DiaMaxModulo
			Set RstModulos=Session("Conn").execute("select top 1 dia from ra_modulo where dia='LUNES' and orden in (select max(convert(int,orden)) as orden from ra_modulo )")
			DiaMaxModulo=RstModulos(0)
'			response.Write(DiaMaxModulo)
			RstModulos.close 	
			%>
        <% strModulos = "Select orden, Hor_Ini, Hor_Fin, codmod from ra_modulo where CodSede = '" & CodSede & "' and dia = '"& trim(DiaMaxModulo) &"' Order By convert(numeric,Orden)" 
           set RstModulos = Session("Conn").execute(strModulos)
           do while not RstModulos.eof
             i = RstModulos("Orden")
        %>
		  <div id="tooltip" align="center" style="position:absolute;visibility:hidden;border:1px solid black;font-size:12px;layer-background-color:lightyellow;background-color:lightyellow;padding:1px" ></div>
          <tr bgcolor="#ffe3e3"> 
             
            <td width="24" height="17" align="center"><font face="Arial" size="1"><%=RstModulos("Orden")%></font></td>
             <td width="98" height="17" align="center"><font face="Arial" size="1"><%=RstModulos("Hor_Ini") & " - " & RstModulos("Hor_Fin")%></font></td>
          <% for j = 1 to 6 %>
             
             <% if trim(MtrHorSt(i, j)) = "" then 
                  color = "#ffe3e3"
             %>
              <td width="94" height="17"><font face="Verdana" size="1">  <br> &nbsp 
             <% 
                else
                   Select case  MtrHorSt(i,j) 
                     Case "I": Color = "#00FF00"
                     Case "E": Color = "#FF0000"
                     Case "P": Color = "#FFCC66"
                     Case "T": Color = "#ffe3e3"
                     Case "S": Color = "#FFCC66"
                   End Select
              %>                
                 <td bgcolor="<%=color%>" width="94" height="17" onMouseOver="javascript:showtip(this, event, '<%=MtrHorRamo(i, j)%>')" onMouseout="javascript:hidetip()"><font face="Verdana" size="1"><%=MtrHorario(i, j) %> 
             <% end if%>
            </font></td>
          <% next %>
          </tr>
        <%
           RstModulos.Movenext 
           loop 
         %>    
          </table>        </td>
      </tr>
      <tr valign="middle"> 
        <td colspan="2" height="10"> 
          <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#666666">(Este 
            documento NO constituye certificado)</font></div>        </td>
      </tr>
	</table>
	</form></table>
	</td>
  </tr>
</table>

</body>

<script languaje = "javascript">

function ConfirmaInscrip()
{
  document.form1.action = "solic-confirmacion.asp"
 // alert("Hola");
  document.form1.submit();
} 
function visualiza() {
  var x=window.open ("tex-resumenInscrip.htm","ResumenInscripción","width=500,height=400,resizable=yes,toolbar=yes");
 }

function imprime()
{
  //parent.focus();
   window.print() 
  //window.print();
  //parent.Horario.focus();
  //parent.Horario.print(); 
}

<%if trim(thisID)<>"" then%>
	$('<%=thisID%>').activate();
<%end if%>	

</script>
</html>
<%
RamosDebe.Close()
RamosSolici.Close()
%>
<!--#INCLUDE file="include/desconexion.inc" -->
