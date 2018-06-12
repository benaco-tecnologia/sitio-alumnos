<%
    Response.Expires = -1
    Session.Timeout = 999
    Response.Buffer = false
    Server.ScriptTimeout = 99999
%>
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%
  dim Valor
  dim strOpcionTexto
  strPregunta=trim(request("Pregunta"))
  strNumero=trim(request("NumPre"))
  Accion=request.form("hidAccion")
  Valor=trim(request("CmbValor"))
  strOpcion=trim(request("CmbValor"))
  strOpcionTexto= request.form("txtOpcion")
  if trim(request.form("txtOpcion")) = "" then 
    strOpcion=request.form("CmbValor")
  else  
     strOpcion= request.form("txtOpcion")
  end if
  'response.Write(strOpcion)
  
     Dim numPreg
	'Total Pregunta Encuesta
     sql="Select count(*) from ra_preguntas"
     'set rst=conn.execute(sql)
     'if not rst.eof then
	 if bcl_ado(sql,rst) then
        numPreg=valnulo(rst(0),num_)
     else
        numPreg=0
     end if
  if Accion="Actualizar" then
      if valor="OTRO VALOR" then
	  strOpcion= request.form("txtOpcion")
	  	session("respuesta")=strOpcion
	      else
		  if trim(request.form("txtOpcion")) = "" then 
		    strOpcion=request.form("CmbValor")
		    else  
		    strOpcion= request.form("txtOpcion")
		end if
 			session("respuesta")=strOpcion
	  end if 
  end if
   
  if Accion= "Grabar" then
	 strOpcion=session("respuesta")
	 if strOpcion="" then
	    strOpcion=strOpcionTexto
	 end if 
      if strOpcion <> "" then
	     sql="Select * from ra_resultados where Ano='" & session("Ano_Ed") & "'"
		 sql=sql & " and Periodo='" & session("Periodo_Ed") & "' and CodCli='" & session("CodCli") & "'"
		 sql=sql & " and Rut='" & session("Rut") & "' and CodRamo='" & session("CodRamo") & "'"
		 sql=sql & " and CodSecc='" & session("CodSecc") & "' and CodProf='" & session("CodProf") & "'"
		 sql=sql & " and TipoDocente='" & session("TipoDoc") & "' and NroPregunta=" & strNumero & ""
		 'set rs = Conn.execute (sql)
		 if bcl_ado(sql,rs) then
		 'if rs.eof then         'Inserta
               sql="Insert Into ra_Resultados (Ano,Periodo,CodCli,Rut,"
	           sql=sql & "CodRamo,CodSecc,CodProf,TipoDocente,NroPregunta,Respuesta) values("
	           sql=sql & "'" & session("Ano_Ed") & "','" & session("Periodo_Ed") & "' ,'" & session("CodCli") & "','" & session("Rut") & "',"
	           sql=sql & "'" & session("CodRamo") & "','" & session("CodSecc") & "' ,'" & session("CodProf") & "',"
 	           sql=sql & "'" & session("TipoDoc") & "'," & strNumero & " ," & valnulo(strOpcion,num_) & ")"
	           'response.Write(sql)
	     else                 'Actualiza
		       sql="Update ra_resultados set Respuesta=" & valnulo(strOpcion,num_) & " where Ano='" & session("Ano_Ed") & "'"
		       sql=sql & " and Periodo='" & session("Periodo_Ed") & "' and CodCli='" & session("CodCli") & "'"
		       sql=sql & " and Rut='" & session("Rut") & "' and CodRamo='" & session("CodRamo") & "'"
		       sql=sql & " and CodSecc='" & session("CodSecc") & "' and CodProf='" & session("CodProf") & "'"
		       sql=sql & " and TipoDocente='" & session("TipoDoc") & "' and NroPregunta=" & strNumero & ""		     
		 end if 		   
		 Session("Conn").execute (sql)
		 
         ' Si total de Respuesta es igual al numero de Preguntas , Actualiza a Profesor Encuestado
		 sql="Select Count(*) from ra_resultados where Ano='" & session("Ano_Ed") & "'"
		 sql=sql & " and Periodo='" & session("Periodo_Ed") & "' and CodCli='" & session("CodCli") & "'"
		 sql=sql & " and Rut='" & session("Rut") & "' and CodRamo='" & session("CodRamo") & "'"
		 sql=sql & " and CodSecc='" & session("CodSecc") & "' and CodProf='" & session("CodProf") & "'"
		 sql=sql & " and TipoDocente='" & session("TipoDoc") & "'"
		'response.Write(sql)
	     'set rs = Conn.execute (sql)
         'if not rs.eof then
		 if bcl_ado(sql,rs) then
            num_Respuestas=valnulo(rs(0),num_)	
			'response.Write(num_Respuestas & "<br>" & TotalRespuesta())					
			if num_Respuestas = TotalRespuesta() then
			    Strsql= " Update RA_ENCUESTADOS set Evaluado='S'"
				Strsql=Strsql & " WHERE  ANO = " & valnulo(session("Ano_Ed"),NUM_) & ""
                Strsql=Strsql & " AND  PERIODO = " & valnulo(session("Periodo_Ed"),NUM_) & " " 
                Strsql=Strsql & " AND  CODSEC = " & session("CodSecc") & "" 
                Strsql=Strsql & " AND  CODRAMO ='" & session("CodRamo") & "'"
                Strsql=Strsql & " AND  CODPROF ='" & session("CodProf") & "'"
				Strsql=Strsql & " AND  CODCLI = '" & session("CodCli") & "'"
				Session("Conn").execute (strsql)
			end if         
         end if
		 strPregunta=""
	            %>
				 <script>
				    parent.bottomFrame3.location.href="list-preguntas.asp";														    
				 </script>
				 <%
     end if				 
  end if
 
function GetRespuestaDes( Col,Row)
dim Rq
dim Sq

Sq="select descripcion from RA_DETALLERESPED "
Sq=Sq & " where coddescripcion="& row &""
Sq=Sq & " and respuesta="& col &"" 
'set Rq=conn.execute (Sq)
if bcl_ado(Sq, Rq) then 
	'if col=1 or col=2 then
	'	GetRespuestaDes=trim(Rq("Descripcion")) & " Hola"
	'else
		GetRespuestaDes=trim(Rq("Descripcion"))
	'end if
else
	GetRespuestaDes=""
end if
Rq.close
end function

 function GetNombreAsig(Ramo)
 Dim StrSql
 Dim Rst

 'Strsql="select nombre,paterno,materno from mt_client where codcli='" & codcli & "'"
 strsql="Select Nombre from ra_ramo where codramo='" & session("CodRamo") & "'"
 'Set Rst = Conn.Execute(StrSql)
  if bcl_ado(strsql,Rst) then
  'if not Rst.eof then
    GetNombreAsig = Ucase(Rst("nombre"))
  else
    GetNombreAsig = "Sin Nombre"
  end if
  Rst.close() 
 end function
%>
<%


'strvalor = request.form("CmbValor")

'set rs=Server.CreateObject ("ADODB.Recordset")
'	sql="select respuesta "
'	sql=sql & " from ra_respuestaed "

'set rs=Conn.Execute (sql)
'if Paso=0 then
'	if not rs.EOF then
'		CmbValor=""
'		pricarr=true	
'		while not rs.EOF
'			if CmbValor="" and pricarr then
'				'sel=" selected"
'				pricarr=false
'				strvalor=rs("respuesta") 
'			else
'				if strvalor=rs("respuesta") then
'					sel=" selected"
'				else
'					sel=""
'				end if
'			end if
'			CmbValor= CmbValor & "<option value=""" & rs("respuesta") & """ " & sel & ">"
'			CmbValor=	CmbValor & rs("respuesta") & "</option>" & vbcrlf
'			rs.MoveNext 
	
'		wend
'	end if
'end if
'paso=1
%>
<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="ed.css" type="text/css">
<script language="JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
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

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}
//-->
</script>
<link href="css/tablas.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="0" onLoad="MM_preloadImages('/ima/bot/aceptar-on.gif','ima/bot/ayuda-on.gif')">
<table border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td valign="top" >
			<% CargarTop1()%><% SubMenu()%>
				<form action="" method="post" name="TheForm" class="tex">
				  <table width="800" border="0" cellspacing="0">
					<tr> 
					  <td width="33%"><font class="Tit-celdas">Informaci&oacute;n General:</font> 
						<input type="hidden" name="hidAccion" value="<%=Accion%>"> </td>
					  <td width="16%">&nbsp;</td>
					  <td width="25%">&nbsp;</td>
					  <td width="12%"> 
						<!--<div align="right"><a href="javascript:Intruc();" onMouseOver="MM_swapImage('Image3','','Imagenes/botones/ayuda-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="Imagenes/botones/ayuda-of.gif" width="178" height="45" name="Image3" border="0"></a></div> -->      </td> 
					  <td width="14%"><div align="right"></div>      </td>
					</tr>
				  </table>
				  <table width="800" height="15" border="0" cellpadding="0" cellspacing="1">
					<tr> 
					  <td width="8%" height="15" background="Imagenes/fdo-cabecera-cel.jpg" class="tex"> <div align="right" class="text-cabecera-celda"><font color="#FFFFFF" id="lblAsignatura">Asignatura</font></div></td>
					  <td width="29%" height="15" class="tex">&nbsp;<font size="2" face="Arial, Helvetica, sans-serif" class="tex-totales-celda"><%=GetNombreAsig(session("CodRamo"))%></font></td>
					  <td width="6%" height="15" background="Imagenes/fdo-cabecera-cel.jpg" class="tex"> <div align="right" class="text-cabecera-celda"><font color="#FFFFFF">Seccion:</font></div></td>
					  <td width="8%" height="15" bgcolor="#F8F7EB" class="tex">&nbsp;<font size="2" face="Arial, Helvetica, sans-serif" class="tex-totales-celda"><%=session("CodSecc")%></font></td>
					  <td width="11%" height="15" background="Imagenes/fdo-cabecera-cel.jpg" class="tex" > <div align="right" class="text-cabecera-celda"><font color="#FFFFFF" class="text-cabecera-celda" id="lblProfe" >Profesor:</font></div></td>
					  <td width="38%" height="15" bgcolor="#F8F7EB" class="tex"><font size="2" face="Arial, Helvetica, sans-serif" class="tex-totales-celda"><%=GetNombreProfe(Session("CodProf"))%></font></td>
					</tr>
					<%dim st 
				  dim Rtss
				  Dim Rd
				  dim i 
				  dim j
				  dim SS 
				  dim Rr
				  dim Cantidad
				  dim Tama
				  vacio=""
'				  st="select max(respuesta) as Respuesta from ra_respuestaed where respuesta<>'0'"				  
'				  if bcl_ado(st,Rd) then				  
'					cantidad=Rd("Respuesta")
'				  else
'					Cantidad=7
'				  end if
'				  st="select distinct respuesta as respuesta,descripcion from ra_respuestaed with (nolock) where respuesta<>'0'"
'				  set Rtss=Session("Conn").execute(st)
				  'if bcl_ado(sql,rs) then
				  %>
				  </table>
				  <script>
				function Grabar(valor) {
				   if (valor=="") {
					  alert("Debe Seleccionar una Pregunta de la Lista");
					  return;
					 } 	  
					  
				   if (document.TheForm.txtOpcion.value!=""){ 
								if ((isNaN(document.TheForm.txtOpcion.value))==true)
								   {	alert("Valor debe ser Númerico(Entero)");
										document.TheForm.txtOpcion.select();
										document.TheForm.txtOpcion.focus();
										return;
								}else {
								if (document.TheForm.txtOpcion.value < 1 || document.TheForm.txtOpcion.value > 7) {
									alert ("Valor fuera de rango");
									return;
								   }
								} 
						}	
				  document.TheForm.hidAccion.value="Grabar";
				  document.TheForm.submit();
				}
				function Intruc()
				{ 
				var x=window.open ("Instrucciones.htm","HelpWindow","width=400,height=250,resizable=yes,scrollbars=yes,menubar=no,status=no")
				}
				
				function ActualizarValor() {
				  document.TheForm.hidAccion.value="Actualizar";
				  document.TheForm.submit();
				}
					  </script>
				</form>
	  			</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>
</body>
<%ObjetosLocalizacion("pregunta-opc.asp")%>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->