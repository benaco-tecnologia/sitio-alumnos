<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%

dim strcarrera
dim Nombrecarrera
strcarrera = request.form("cmbCarrera")
Nombrecarrera=GetNombreCarrera(strcarrera)
Accionvar= request.form("Accion")
Rut=request(Rut)
RutCli=session("RutCli")

set rs=Server.CreateObject ("ADODB.Recordset")
	
	sql="select b.codcarr,b.nombre_l,estacad "
	sql=sql & " from mt_alumno a, mt_carrer b "
	sql=sql & " where a.codcarpr=b.codcarr"
	sql=sql & " and a.rut='" & Rutcli & "'"

set rs=Conn.Execute (sql)
if not rs.EOF then
	cmbCarrera=""
	pricarr=true	
	while not rs.EOF
		if strcarrera ="" and pricarr then
			'sel=" selected"
			pricarr=false
			strcarrera=rs("codcarr") & " " & rs("Nombre_l") 
		else
			if strcarrera=rs("codcarr") then
				sel=" selected"
			else
				sel=""
			end if
		end if
		
		cmbCarrera= cmbCarrera & "<option value=""" & rs("codcarr") & """ " & sel & ">"
		cmbCarrera=	cmbCarrera & rs("CodCarr") & " " & rs("nombre_l") & " " & rs("estacad") & "</option>" & vbcrlf
		rs.MoveNext 

	wend
end if
%>
<% 
if AccionVar="Enviar" then
  session("CarreraAlumno")=strcarrera
  Accionvar=""
  response.Redirect "ValidaClave.asp"
end if
if AccionVar="Seleccionar" then
	NombreCarrera=GetNombreCarrera(strcarrera)
	AccionVar=""
end if
%>

<html>
<head>
<title>Eleccion de Carrera</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/css/parrafo.css" rel="stylesheet" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
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
//-->
</script>
</head>
<body bgcolor="#FFFFFF" background="ima/iso-uhc.gif" text="#CC3300" vlink="#0033CC" onLoad="MM_preloadImages('ima/bot/aceptar-on.gif')">
<form action="EligeCarrera.asp" method="post" name="TheForm" id="TheForm">
  <p>&nbsp;</p>
  <table width="700" height="100" border="0" cellspacing="15">
    <tr> 
      <td height="68"><div align="justify"><font face="Arial, Helvetica, sans-serif">Usted 
          se encuentra en m&agrave;s de una carrera en nuestros registros. Deber&agrave; 
          elegir una de las carreras que ha cursado para ingresar a nuestro sitio..</font></div></td>
    </tr>
  </table>
  <p>
    <input type="hidden" name="Accion" value="<%=AccionVar%>">
  </p>
  <table width="700" height="189" border="0" cellspacing="15">
    <tr> 
      <td width="175" height="24"><font color="#000000" face="Arial, Helvetica, sans-serif"><strong>Seleccione 
        la Carrera</strong></font></td>
      <td width="515"><select name="cmbCarrera" onChange="javascript:Actualizar();">
            <option value="">SELECCIONE </option>
			<%=cmbCarrera%>
		     </select></td>
    </tr>
    <tr> 
      <td height="21"><font color="#000099" size="2" face="Arial, Helvetica, sans-serif"><strong>Codigo 
        Carrera :</strong></font></td>
      <td><%=strcarrera%></td>
    </tr>
    <tr> 
      <td height="21"><font color="#000099" size="2" face="Arial, Helvetica, sans-serif"><strong>Nombre 
        Carrera :</strong></font></td>
      <td><%=NombreCarrera%></td>
    </tr>
    <tr> 
      <td colspan="2"><div align="center"> <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('ImagenEligeCarrera','','ima/bot/aceptar-on.gif',1)"> 
          </a> </div></td>
    </tr>
  </table>
  <p align="center"> 
    <input name="Emviar" type="button" id="Emviar2" style="color: #000000;FONT-SIZE: 10px; height:20; font-weight: bold;width:60" onClick="javascript:Enviar();" value="Enviar">
  </p>
  <p>&nbsp;</p>
</form>
</body>
</html>
<script>
function Actualizar() {
  document.TheForm.Accion.value="Seleccionar";
  document.TheForm.submit();
}
function Enviar() {
     document.TheForm.Accion.value="Enviar";
     document.TheForm.submit();
}
</script>
