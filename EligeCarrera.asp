<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenuexterno.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%
'response.write("este es rut en elige carrera "& session("RutAlum"))
'response.end
dim strcarrera , i
dim Nombrecarrera , codcarrX, carrAux,estacadAux
session("RX")=SESSION("RutAlum")
Nombrecarrera=GetNombreCarrera(strcarrera)
Accionvar= request.form("Accion")
Rut=request(Rut)
RutCli=session("RutCli")

Valores = request.form("cmbCarrera")
if Valores<> "" then
	Contador=0
	for i=1 to len(Valores)
		letra=mid(Valores,i,1)
		if letra = ";" then
			contador = contador + 1
		else
			if contador = 0 then
				strcarrera = strcarrera &	letra
			elseif  contador = 1 then
				estacad = estacad &	letra
			end if 
		end if
	next
end if 

set rs=Server.CreateObject ("ADODB.Recordset")
'	sql="select b.codcarr,b.nombre_l "
'	sql=sql & " from mt_alumno a, mt_carrer b "
'	sql=sql & " where a.codcarpr=b.codcarr"
'	sql=sql & " and a.rut='" & Rutcli & "'"
	sql="select b.codcarr,b.nombre_l,a.estacad "
	sql=sql & " from mt_alumno a, mt_carrer b "
	sql=sql & " where a.codcarpr=b.codcarr"
	sql=sql & " and a.rut='" & Rutcli & "'"
	'sql=sql & " and a.estacad in ('VIGENTE','EGRESADO')"	
	if strcarrera<>"" then
		sql=sql & " and a.codcarpr ='" & strcarrera &"' and a.estacad ='"& estacad &"'"
	end if 
	sql=sql & "order by a.ano desc , a.periodo desc"


set rs=Session("Conn").Execute (sql)

  i = 0
 
if not rs.EOF then
	cmbCarrera=""
	pricarr=true
	estacadAux = rs("estacad")	
	carrAux = rs("codcarr")
	session("EstadoCarrera") = rs("estacad")	
	codcarrX = session("CarreraAlumno")=strcarrera
	while not rs.EOF
		if strcarrera ="" and pricarr then
			'sel=" selected"
			session("EstadoCarrera") = rs("estacad")
			pricarr=false
			strcarrera=rs("codcarr") '& " " & rs("Nombre_l")
			Nombrecarrera=GetNombreCarrera(strcarrera)
		else
			if strcarrera=rs("codcarr") then
				sel=""
			else
				sel=""
			end if
		end if
		cmbCarrera= cmbCarrera & "<option value=""" & rs("codcarr") & ";" & rs("estacad") & """" & sel & ">"
		cmbCarrera=	cmbCarrera & rs("CodCarr") & " " & rs("nombre_l") & " (" & rs("estacad") & ") " & "</option>" & vbcrlf
		rs.MoveNext 
       i = i + 1
	wend
end if



if i = 1 then   
	session("EstadoCarrera") = estacadAux
	session("CarreraAlumno") = carrAux
	Session("usrNWRC")=session("RX")
	Session("usrNWRC") = Replace(Session("usrNWRC"), ".", "")
	Session("usrNWRC") = Replace(Session("usrNWRC"), "-", "")	
	response.Redirect("ValidaClave.asp")
end if

%>
<% 
if AccionVar="Enviar" then
'response.Write(strcarrera)
'response.end()
  session("CarreraAlumno")=strcarrera
 'session("CarreraAlumno") = codcarrX
  Accionvar=""
 'response.Write("VUELVE A VALIDA")
 'response.End()
 
  response.Redirect("ValidaClave.asp")

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
<link rel="stylesheet" href="css/tablas.css" type="text/css">
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

<body bgcolor="#FFFFFF" background="ima/iso-uhc.gif" text="#CC3300" vlink="#0033CC" onLoad="MM_preloadImages('ima/bot/aceptar-on.gif')">
<table border="0" cellpadding="0" cellspacing="0"  align="left">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right"><% CargarTop()%><% 
				CargarMenu2() 
			    %></td>
			<td valign="top">
			<% CargarTop1()%><% SubMenu()%>
                <form action="EligeCarrera.asp" method="post" name="TheForm" id="TheForm">
                 <table width="700" height="10" border="0" cellspacing="15">
                    <tr> 
                      <td><div align="justify"><font face="Arial, Helvetica, sans-serif"> Deber&aacute; 
                      elegir la  carrera que ha cursado para ingresar a nuestro sitio..</font></div></td>
                    </tr>
                  </table>
                  <p>
                    <input type="hidden" name="Accion" value="<%=AccionVar%>">
                    
                  </p>
                  <table width="700" height="18" border="0" cellspacing="15">
                    <tr> 
                      <td width="175" height="24"><font color="#000000" face="Arial, Helvetica, sans-serif"><strong>Seleccione 
                        la Carrera</strong></font></td>
                      <td width="515"><select name="cmbCarrera" onChange="javascript:Actualizar();">
                            <option value="">SELECCIONE </option>
                            <%=cmbCarrera%>
                             </select>
                </td>
                    </tr>
                    <tr> 
                      <td height="21">&nbsp;</td>
                      <td>&nbsp;</td>
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
                      <td height="21">&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                    <tr> 
                      <td height="21" colspan="2">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td colspan="2"><div align="center"> <a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('ImagenEligeCarrera','','ima/bot/aceptar-on.gif',1)"> 
                          </a> </div></td>
                    
                    </tr>
                  </table>
                  <p align="center"  style="display:none"> 
                    <input name="Emviar" type="button" id="Emviar2" style="color: #FFFFFF;FONT-SIZE: 10px; height:20; font-weight: bold;width:60" onClick="javascript:Enviar();" value="Enviar">
                  </p>
                  <p>&nbsp;</p>
                </form>
			</table>
			</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>                
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