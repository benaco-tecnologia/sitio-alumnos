<%
Dim Direc_Ini
Dim var_servidor


Direc_Ini = "C:/simulador_bpe/simuladorBPE.ini"

var_servidor = LeeIni(Direc_Ini,"servidor")

'------------------------------------------------------------------------------------------------------------------
' Lee ini
'------------------------------------------------------------------------------------------------------------------
	Function LeeIni(archivo,linea)
		Const fsoLectura = 1
  		Dim objFSO
		Dim objTextStream
		Dim Key


		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

		If objFSO.FileExists(archivo) Then
		   Set objTextStream = objFSO.OpenTextFile(archivo, fsoLectura,true)
			Do While not objTextStream.AtEndOfStream
     			Key = Split(objTextStream.ReadLine, "=")
				If Key(0) = linea Then
				 	LeeIni = Key(1)
					Exit Do
				End If
   			loop
   				objTextStream.Close
   			Set objTextStream = Nothing
		Else
	    	LogTransac "Archivo Servidor", "Archivo de Token de direccion, no existe"
			LeeIni = "No Existe Archivo"
		End if
	End Function
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Simulador Boton de Pago Estandar</title>
<link href="estilos.css" rel="stylesheet" type="text/css" />
<script language="javascript" type="text/javascript">
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
<script language="javascript" type="text/javascript">
<!---

function enviar()
{

	if (document.datos.canal.value  == '')
  		{
  			alert("Debe ingresar un valor en el campo Canal.");
  			return;
  		}
	if (document.datos.identificador.value  == '')
  		{
  			alert("Debe ingresar un valor en el campo Identificador.");
  			return;
  		}

	if (document.datos.subtotal.value  == '')
  		{
  			alert("Debe ingresar un valor en el campo subtotal.");
  			return;
  		}

	var props = "height=425,width=675,top=100,left=100,resizable=no,menubar=no,location=no,toolbar=no,status=no,scrollbars=no,directories=no";
	xml = "?xml=<?xml version='1.0' encoding='ISO8859-1'?><Servipag><Header><FirmaEPS>r5mOOrC6YZcy2SLtwilP1hIugm802xu/TRNjuHwanlVUfhqPtCUX9B5LDuqfHGiRvwvYKFuahQwpP9jxzNThuLBosKoKvblH3GFBT3ek2HiKmsf61Sv0HoFxYcWkXK4nVFknCTFpmaMaQ9CoZRtkdtqZAW1nB/Jc1C37ru10WO4=</FirmaEPS><CodigoCanalPago>";
	xml1 ="</CodigoCanalPago><IdTxCliente>696969</IdTxCliente><FechaPago>20070205</FechaPago><MontoTotalDeuda>";
	xml2 ="</MontoTotalDeuda><NumeroBoletas>1</NumeroBoletas></Header><Documentos><IdSubTrx>1</IdSubTrx><CodigoIdentificador>";
	xml3 ="</CodigoIdentificador><Boleta>88888</Boleta><Monto>";
	xml4 ="</Monto><FechaVencimiento>20070205</FechaVencimiento></Documentos></Servipag>";
	window.open('https://<%=var_servidor%>/bpe/BPE_Inicio.asp' + xml+document.datos.canal.value+xml1+document.datos.subtotal.value+xml2+document.datos.identificador.value+xml3+document.datos.subtotal.value+xml4, 'Simulador', props);
}

// --->
</Script>
</head>

<body onload="MM_preloadImages('img/boton-a.gif')">
<form name="datos" action="/browse.asp" method="post">
<table width="772" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="img/top-right.gif"><img src="img/spacerhoriz.gif" width="15" height="10" /></td>
    <td background="img/top.gif"><img src="img/spacerhoriz.gif" width="16" height="10" /></td>
    <td background="img/top-left.gif"><img src="img/spacerhoriz.gif" width="15" height="15" /></td>
  </tr>
  <tr>
    <td width="5" background="img/left.gif">&nbsp;</td>
    <td bordercolor="0"><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFCC00">
      <tr>
        <td height="24" colspan="5" background="img/top-amarillo.gif">&nbsp;</td>
        </tr>
      <tr>
        <td width="4" height="100" rowspan="11" bgcolor="#717274"><img src="img/spacerhoriz.gif" width="4" height="10" /></td>
        <td width="100%" height="27" colspan="3" background="img/titulo.gif" bgcolor="#E1E6EA"><span class="titulo_blanco_bold">SIMULADOR BOTON DE PAGO ESTANDAR </span></td>
        <td width="4" rowspan="11" bgcolor="#717274"><img src="img/spacerhoriz.gif" width="4" height="10" /></td>
      </tr>
      <tr>
        <td colspan="3" bgcolor="#FFFFFF"><img src="img/spacerhoriz.gif" width="15" height="10" /></td>
      </tr>
      <tr>
        <td height="20" colspan="3" bgcolor="#E1E6EA"><span class="subtitulo">INFORMACI&Oacute;N PERSONAL </span></td>
      </tr>
      <tr>
        <td width="30" align="left" bgcolor="#E1E6EA">&nbsp;</td>
        <td width="100%" align="left" bgcolor="#E1E6EA"><table width="500" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" id="datos">
          <tr>
            <td width="100" bgcolor="#FACC21" class="titulo_blanco_tablas">Nombre:</td>
            <td><input disabled="true" name="nombre" type="text" class="titulo_gris_tablas" id="nombre" value="Jaime Andr&eacute;s Silva Hern&aacute;ndez" size="50" /></td>
          </tr>
          <tr>
            <td bgcolor="#FACC21" class="titulo_blanco_tablas">Tel&eacute;fono:</td>
            <td><input disabled="true" name="telefono" type="text" class="titulo_gris_tablas" id="telefono" value="2425638" size="50" /></td>
          </tr>
          <tr>
            <td bgcolor="#FACC21" class="titulo_blanco_tablas">Celular:</td>
            <td><input disabled="true" name="celular" type="text" class="titulo_gris_tablas" id="celular" value="09-3556248" size="50" /></td>
          </tr>
        </table></td>
        <td width="30" height="30" align="left" bgcolor="#E1E6EA">&nbsp;</td>
      </tr>
      <tr>
        <td height="20" colspan="3" bgcolor="#E1E6EA"><span class="subtitulo">CUENTA A PAGAR  </span></td>
      </tr>
      <tr>
        <td width="7" align="left" bgcolor="#E1E6EA">&nbsp;</td>
        <td align="center" bgcolor="#E1E6EA"><table width="500" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" id="datos">
          <tr>
            <td width="100" align="left" bgcolor="#FACC21" class="titulo_blanco_tablas">Canal:</td>
            <td align="left"><input  name="canal" type="text" class="titulo_gris_tablas" id="canal" size="50" /></td>
          </tr>
          <tr>
            <td align="left" bgcolor="#FACC21" class="titulo_blanco_tablas">Identificador:</td>
            <td align="left"><input name="identificador" type="text" class="titulo_gris_tablas" id="identificador" size="50" /></td>
          </tr>
          <tr>
            <td align="left" bgcolor="#FACC21" class="titulo_blanco_tablas">Valor: $</td>
            <td align="left"><input  name="subtotal" type="text" class="titulo_gris_tablas" id="subtotal"  size="50" /></td>
          </tr>
        </table></td>
        <td width="7" align="left" bgcolor="#E1E6EA">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="3" align="center" bgcolor="#E1E6EA"><a onclick="enviar()" href="#" onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('boton','','img/boton-a.gif',1)"><img src="img/boton-b.gif" alt="servipag" name="boton" width="200" height="70" border="0" id="boton" /></a></td>
      </tr>
      <tr>
        <td colspan="3" align="center" bgcolor="#E1E6EA">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="3" bgcolor="#FFFFFF"><img src="img/spacerhoriz.gif" width="15" height="10" /></td>
      </tr>
      <tr>
        <td colspan="5" background="img/bottom-amarillo.gif"><img src="img/spacerhoriz.gif" width="15" height="10" /></td>
      </tr>
    </table></td>
    <td width="5" background="img/right.gif">&nbsp;</td>
  </tr>
  <tr>
    <td bordercolor="0" background="img/bottom-left.gif"><img src="img/spacerhoriz.gif" width="15" height="15" /></td>
    <td background="img/bottom.gif"><img src="img/spacerhoriz.gif" width="16" height="10" /></td>
    <td background="img/bottom-right.gif"><img src="img/spacerhoriz.gif" width="15" height="10" /></td>
  </tr>
</table>
</form>
</body>
</html>
