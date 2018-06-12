<%  response.buffer = false 
    Response.Expires = -1
%>

<%
If Session("NomAlum") = "" then%> 
	
	<script>
		alert("Expiro la sesión");
		window.close();
	</script>
   
<% response.end()
 end if %>
<!-- saved from url=(0022)http://internet.e-mail -->
<!--<OBJECT RUNAT=server PROGID=ADODB.Recordset id=malla name=malla></OBJECT> -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--<OBJECT RUNAT=server PROGID=ADODB.Recordset id=Detalle name=Detalle > </OBJECT>-->
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=Datos name=Datos></OBJECT>
<%
Dim Rut, strSql, Estado, GlosaError,strsql2, Imprime
Rut = RutSinDV(Session("RutAlum"))
codcli=session("codcli")
rutcompleto=trim((Session("logrut")))
Session("Imprime") = 0

if Session("logrut") = "" then
  'rutcompleto =(Session("RutCli")
   rutcompleto= Session("RutCliente")
end if  

If Session("NomAlum") = "" then 
	response.Redirect("salir.asp")
	response.End()
end if 


if session("CarreraAlumno") <> "" then 
            Sql30 = ""
			Sql30 = "Select estacad from mt_alumno where rut = '" & Session("RutCli")& "'"
			Sql30 =  Sql30 & " and codcarpr ='" & session("CarreraAlumno") & "' "
			'response.Write(sql30)
			'response.End()
			
			Set Rst30 = Session("Conn").Execute(Sql30)
			if not rst30.eof then 
				if rst30("estacad") <>"" then 
				 GlosaError = rst30("estacad")
				 'else 
				 'GlosaError = "Año Actual"
			   end if 
			end if
else
            Sql30 = "Select estacad from mt_alumno where rut = '" & Rut & "'"
			'Sql30 = sql30 & "matriculado ='S' "
'			'response.Write(sql30)
			'response.End()
			
			Set Rst30 = Session("Conn").Execute(Sql30)
			if not rst30.eof then 
				if rst30("estacad") <>"" then 
				 GlosaError = rst30("estacad")
				 'else 
'				 GlosaError = "Año Actual"
			   end if 
			end if
end if 			

'Datos.Open strSql, Conn
'Estado = Trim(Datos(0))
'Select Case Estado
'	Case "ELIMINADO":
'	     GlosaError = "Ud. se encuentra Eliminado de los registros de la  Universidad."
'	Case "EGRESADO"	 
'	     GlosaError = "Su estado actual es de Egresado en nuestros registros."
'	Case "TITULADO"	 
'	     GlosaError = "Su estado actual es de Titulado en nuestros registros."
'	Case "SUSPENDIDO"	 		
'	     GlosaError = "Ud. se encuentra suspendido de la Universidad."
'	Case else
'		 GlosaError = "Vigente/Alumno Regular"
'End Select
'Datos.Close()
'Conn.Close()

'dim strsql,strsql2

'rut=trim(Session("RutAlum"))
Function ValNulo(varDum, intTip) 
    If VarType(varDum) = vbNull Or varDum = "" Then
        Select Case intTip
            Case STR_
                ValNulo = ""
            Case NUM_
                ValNulo = 0
            Case DAT_
                ValNulo = 0 'Null
            Case Else
                ValNulo = "0"
        End Select
    Else
        If intTip = NUM_ Then
            ValNulo = cDbl(varDum)
        Else
            ValNulo = varDum
        End If
    End If
End Function

'codcli=200011001
'response.Write(codcli)
'response.end

strsql2 = ""
strsql2 =strsql2 & "Select CORRELATIVO_FOLIO_CUPO from mt_parame"
'strsql2 =strsql2 & "select b.nombre,a.ctadocnum as Numero,a.fecven,a.monto,a.saldo "
'strsql2= strsql2 & "from mt_ctadoc a, mt_docum b,mt_estdoc c, mt_item i, mt_detubi d "
'strsql2= strsql2 & "where a.codcli = '" & trim(Session("RutCli")) & "' "
'strsql2= strsql2 & "and a.ctadoc = b.tipodoc and a.estado = c.estado  "
'strsql2= strsql2 & "and a.item*=i.coditem and a.ubicacion=d.codubifis "
'strsql2= strsql2 & "and a.saldo > 0 and a.estado = 2 and convert(varchar,a.fecven,112) < convert(varchar,getdate(),112) "

'strsql2= strsql2 & "order by a.fecven asc "
'RESPONSE.Write(STRSQL2)
'RESPONSE.End()

Set Rst = Session("Conn").Execute(StrSql2)

if not rst.eof then  
       Numoperacion = valnulo(rst("CORRELATIVO_FOLIO_CUPO"),num_) 
	    
	 else
	   Numoperacion = "1" 				
end if

%>
<html>
<head>
<title>Cup&oacute;n de Impresi&oacute;n Matr&iacute;cula</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
<link rel="stylesheet" href="file://///Gotzilla/uddweb/admalumnos/css/tex-normales.css" type="text/css">
<link href="css/tex-normales.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.Estilo19 {
	color: #FFFFFF;
	font-weight: bold;
}
.Estilo20 {color: #FFFFFF; font-weight: bold; font-size: 16px; }
-->
</style>
<link href="css/textos.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.Estilo21 {color: #000000}
.Estilo22 {font-size: 10px}
-->
</style>
<script language="JavaScript" type="text/javascript">

function imprimir()
{ if ((navigator.appName == "Netscape")) { window.print() ;
}
else
{ var WebBrowser = '<OBJECT ID="WebBrowser1" WIDTH=0 HEIGHT=0 CLASSID="CLSID:8856F961-340A-11D0-A96B-00C04FD705A2"></OBJECT>';
document.body.insertAdjacentHTML('beforeEnd', WebBrowser); WebBrowser1.ExecWB(6, -1); WebBrowser1.outerHTML = "";
}
}

function CargaImpresion(){
	
	window.print();
	//alert(<%=Session("Imprime")%>);
	 <%
	 'if Session("Imprime") = 1 then
	 Numoperacion = Numoperacion + 1 
	 Sql = " Update mt_parame set CORRELATIVO_FOLIO_CUPO = '" & trim(Numoperacion) & "' "
	 Session("Conn").Execute Sql
	'end if 
	 %>

	window.location.href='cuponera.asp' ;
	
}

function Imprime(){

	<%  Session("Imprime") = 1  %>
	CargaImpresion();
}

</script>

</head>
<body >
<div align="left"> 
  <table border="0" cellpadding="0" cellspacing="0" height="308" align="left" width="336">
    <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
    <tr>
      <td colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td width="233">&nbsp;</td>
      <td width="233">&nbsp;</td>
      <td width="438">&nbsp;</td>
    </tr>
    <tr> 
      <td colspan="3" class="Estilo19" align="center"><span class="titulo">CUP&Oacute;N DE EMISI&Oacute;N</span></td>
    </tr>
    <tr valign="bottom"> 
      <td colspan="3" height="30" class="tex-normales">&nbsp;</td>
    </tr>
    <tr valign="top"> 
      <td colspan="3" height="8"><table width="334" height="496" border="0" cellpadding="3" cellspacing="0" >
        <tr bgcolor="#D6D6D6">
          <td height="47" colspan="2" bgcolor="#ECFBFF" class="textos"><img src="Imagenes/LogoCupon/izquierda.jpg" width="76" height="51"></td>
          <td colspan="7" align="center" bgcolor="#ECFBFF" class="textos"><strong>UNIVERSIDAD METROPOLITANA DE CIENCIAS DE LA EDUCACI&Oacute;N</strong></td>
          <td width="21" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          </tr>
        <tr>
          <td height="43" colspan="3" bgcolor="#ECFBFF" class="textos">N&deg; CONVENIO BCI</td>
          <td width="30" bgcolor="#ECFBFF" class="textos"><strong>961</strong></td>
          <td width="5" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td colspan="4" bgcolor="#ECFBFF" class="textos">RUT: <span class="textos" ><strong> <%=rutcompleto%> </strong></span></td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          </tr>
        <tr>
          <td height="36" colspan="2" bgcolor="#ECFBFF" class="textos">NOMBRE:</td>
          <td colspan="8" bgcolor="#ECFBFF" class="textos"> <strong> <%=Session("NomAlum")%> </strong></td>
          </tr>
        <tr>
          <td height="39" colspan="3" bgcolor="#ECFBFF" class="textos">N&deg; OPERACI&Oacute;N:</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td width="99" bgcolor="#ECFBFF" class="textos"><strong><%=Numoperacion%></strong></td>
          <td width="4" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td width="4" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td width="5" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          </tr>
        <tr>
          <td width="41" height="39" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td width="8" bgcolor="#ECFBFF" class="textos"><input type="checkbox" disabled></td>
          <td colspan="2" bgcolor="#ECFBFF" class="textos">EFECTIVO</td>
          <td bgcolor="#ECFBFF" class="textos"><input type="checkbox" disabled></td>
          <td bgcolor="#ECFBFF" class="textos">DOCUMENTO </td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          </tr>
        <tr>
          <td height="35" bgcolor="#ECFBFF" class="textos">MONTO</td>
          <td colspan="4" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos"><strong>$</strong>&nbsp;<strong><%=session("GetMatricula")%></strong></td>
          <td colspan="4" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td width="35" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td colspan="4" rowspan="3" bgcolor="#ECFBFF" class="textos" height="100" ><textarea cols="15" rows="5" disabled ></textarea></td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td colspan="4" bgcolor="#ECFBFF" class="textos" align="center">Timbre Cajero</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td colspan="7" bgcolor="#ECFBFF" class="textos">Cup&oacute;n Pago Matr&iacute;cula - Copia Banco</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        </table>
      </td>
    </tr>
    <tr valign="top" align="left" > 
      <td colspan="3" height="19"> <div align="center" class="textos">
        Nota: informacion sujeta a confirmaci&oacute;n.
      </div>      </td>
    </tr>
    <tr valign="top"> 
      <td colspan="3" height="2"><p>&nbsp;</p>
      <p>
        <input type="button" name="imprimir" value="Imprimir" onClick="Javascript:Imprime()" >
      </p>      </td>
    </tr>
  </table>
    <table border="0" cellpadding="0" cellspacing="0" height="308" align="left" width="336">
    <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
    <tr>
      <td colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td width="233">&nbsp;</td>
      <td width="233">&nbsp;</td>
      <td width="438">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="3" class="Estilo19" align="center"><span class="titulo">CUP&Oacute;N DE EMISI&Oacute;N</span></td>
    </tr>
    <tr valign="bottom">
      <td colspan="3" height="30" class="tex-normales">&nbsp;</td>
    </tr>
    <tr valign="top">
      <td colspan="3" height="8"><table width="334" height="496" border="0" cellpadding="3" cellspacing="0">
        <tr>
          <td height="47" colspan="2" bgcolor="#ECFBFF" class="textos"><img src="Imagenes/LogoCupon/izquierda.jpg" width="76" height="51"></td>
          <td colspan="7" bgcolor="#ECFBFF" class="textos" align="center"><strong>UNIVERSIDAD METROPOLITANA DE CIENCIAS DE LA EDUCACI&Oacute;N</strong></td>
          <td width="21" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td height="43" colspan="3" bgcolor="#ECFBFF" class="textos">N&deg; CONVENIO BCI</td>
          <td width="30" bgcolor="#ECFBFF" class="textos"><strong>961</strong></td>
          <td width="5" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td colspan="4" bgcolor="#ECFBFF" class="textos">RUT : <strong><%=rutcompleto%></strong> </td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td height="36" colspan="2" bgcolor="#ECFBFF" class="textos">NOMBRE:</td>
          <td colspan="8" bgcolor="#ECFBFF" class="textos"><strong><%=Session("NomAlum")%></strong></td>
        </tr>
        <tr>
          <td height="39" colspan="3" bgcolor="#ECFBFF" class="textos">N&deg; OPERACI&Oacute;N:</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td width="99" bgcolor="#ECFBFF" class="textos"><strong><%=Numoperacion%></strong></td>
          <td width="4" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td width="4" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td width="5" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td width="41" height="39" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td width="8" bgcolor="#ECFBFF" class="textos"><input type="checkbox" disabled></td>
          <td colspan="2" bgcolor="#ECFBFF" class="textos">EFECTIVO</td>
          <td bgcolor="#ECFBFF" class="textos"><input type="checkbox" disabled></td>
          <td bgcolor="#ECFBFF" class="textos">DOCUMENTO </td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td height="35" bgcolor="#ECFBFF" class="textos">MONTO</td>
          <td colspan="4" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos"><strong>$</strong>&nbsp;<strong><%=session("GetMatricula")%></strong></td>
          <td colspan="4" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td width="35" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td colspan="4" rowspan="3" bgcolor="#ECFBFF" class="textos"><textarea cols="15" rows="5" disabled></textarea></td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td colspan="4" bgcolor="#ECFBFF" class="textos" align="center">Timbre Cajero</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td colspan="7" bgcolor="#ECFBFF" class="textos">Cup&oacute;n Pago Matr&iacute;cula - Copia Alumno</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
      </table></td>
    </tr>
    <tr valign="top" align="left" >
      <td colspan="3" height="19"><div align="center" class="textos"> Nota: informacion sujeta a confirmaci&oacute;n. </div></td>
    </tr>
    <tr valign="top">
      <td colspan="3" height="2"><p>&nbsp;</p>
        <p></p></td>
    </tr>
  </table>
  <table border="0" cellpadding="0" cellspacing="0" height="308" align="left" width="336">
    <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
    <tr>
      <td colspan="3">&nbsp;</td>
    </tr>
    <tr>
      <td width="233">&nbsp;</td>
      <td width="233">&nbsp;</td>
      <td width="438">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="3" class="Estilo19" align="center"><span class="titulo">CUP&Oacute;N DE EMISI&Oacute;N</span></td>
    </tr>
    <tr valign="bottom">
      <td colspan="3" height="30" class="tex-normales">&nbsp;</td>
    </tr>
    <tr valign="top">
      <td colspan="3" height="8"><table width="334" height="496" border="0" cellpadding="3" cellspacing="0">
        <tr>
          <td height="47" colspan="2" bgcolor="#ECFBFF" class="textos"><img src="Imagenes/LogoCupon/izquierda.jpg" width="76" height="51"></td>
          <td colspan="7" bgcolor="#ECFBFF" class="textos" align="center"><strong>UNIVERSIDAD METROPOLITANA DE  CIENCIAS DE LA EDUCACI&Oacute;N</strong></td>
          <td width="21" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td height="43" colspan="3" bgcolor="#ECFBFF" class="textos">N&deg; CONVENIO BCI</td>
          <td width="30" bgcolor="#ECFBFF" class="textos"><strong>961</strong></td>
          <td width="5" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td colspan="4" bgcolor="#ECFBFF" class="textos">RUT: <strong><%=rutcompleto%></strong></td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td height="36" colspan="2" bgcolor="#ECFBFF" class="textos">NOMBRE:</td>
          <td colspan="8" bgcolor="#ECFBFF" class="textos"><strong><%=Session("NomAlum")%></strong></td>
        </tr>
        <tr>
          <td height="39" colspan="3" bgcolor="#ECFBFF" class="textos">N&deg; OPERACI&Oacute;N:</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td width="99" bgcolor="#ECFBFF" class="textos"><strong><%=Numoperacion%></strong></td>
          <td width="4" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td width="4" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td width="5" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td width="41" height="39" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td width="8" bgcolor="#ECFBFF" class="textos"><input type="checkbox" disabled></td>
          <td colspan="2" bgcolor="#ECFBFF" class="textos">EFECTIVO</td>
          <td bgcolor="#ECFBFF" class="textos"><input type="checkbox" disabled></td>
          <td bgcolor="#ECFBFF" class="textos">DOCUMENTO </td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td height="35" bgcolor="#ECFBFF" class="textos">MONTO</td>
          <td colspan="4" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos"><strong>$</strong>&nbsp;<strong><%=session("GetMatricula")%></strong></td>
          <td colspan="4" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td width="35" bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td colspan="4" rowspan="3" bgcolor="#ECFBFF" class="textos"><textarea cols="15" rows="5" disabled></textarea></td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td colspan="4" bgcolor="#ECFBFF" class="textos" align="center">Timbre Cajero</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td colspan="7" bgcolor="#ECFBFF" class="textos">Cup&oacute;n Pago Matr&iacute;cula - Matr&iacute;cula</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
        <tr>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
          <td bgcolor="#ECFBFF" class="textos">&nbsp;</td>
        </tr>
      </table></td>
    </tr>
    <tr valign="top" align="left" >
      <td colspan="3" height="19"><div align="center" class="textos"> Nota: informacion sujeta a confirmaci&oacute;n. </div></td>
    </tr>
    <tr valign="top">
      <td colspan="3" height="2"><p>&nbsp;</p>
        <p>&nbsp;</p></td>
    </tr>
  </table>
</div>
</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->