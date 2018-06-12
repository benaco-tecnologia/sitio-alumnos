<!--#INCLUDE FILE="include/conexion.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <title>Pago de cuotas</title>
    <link href="estilos.css" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" href="css/tablas.css" type="text/css" />
    
<script type="text/JavaScript">
	function cerrar() {
		window.open('', '_parent', '');
		window.close();  
	}
</script>

<style type="text/css">
<!--
.style1 {font-size: 24px}  
.style2 {
	font-size: 12px
}
-->
</style>

</head> 
<body style="font-family:arial; font:bold">

<table align="center" >
<tr>
<td>  
<fieldset style="border-radius:10px;border:3px solid #03F"> 
<center>
      <p>
        <img src="images/pago.gif" width="300" height="80" /><br />
        <b class="Tit-celdas style2"><font face="Verdana, Arial, Helvetica, sans-serif">Para consultar su nuevo estado de Pago dir&iacute;jase al
        Men&uacute; Mis Finanzas - Opci&oacute;n Consulta de cuentas. </font> </b>        
        
        <% 
		' Recupera número de Orden de Compra
		OC = Request.Form("TBK_ORDEN_COMPRA")
		
		strP=" SP_TRAE_DATOS_WEBPAY '"& OC &"'"

		if bcl_ado(strP,rstP) then
			TBK_ORDEN_COMPRA = rstP("TBK_ORDEN_COMPRA") 
			TBK_TIPO_TRANSACCION = rstP("TBK_TIPO_TRANSACCION") 
			TBK_RESPUESTA = rstP("TBK_RESPUESTA") 
			TBK_MONTO = rstP("TBK_MONTO")      
			TBK_CODIGO_AUTORIZACION = rstP("TBK_CODIGO_AUTORIZACION")
			TBK_FINAL_NUMERO_TARJETA = rstP("TBK_FINAL_NUMERO_TARJETA")
			TBK_FECHA_CONTABLE = rstP("TBK_FECHA_CONTABLE")
			TBK_FECHA_TRANSACCION = rstP("TBK_FECHA_TRANSACCION")
			TBK_HORA_TRANSACCION = rstP("TBK_HORA_TRANSACCION")
			TBK_ID_SESION = rstP("TBK_ID_SESION") 
			TBK_ID_TRANSACCION = rstP("TBK_ID_TRANSACCION")
			TBK_TIPO_PAGO = rstP("TBK_TIPO_PAGO")
			TBK_NUMERO_CUOTAS = rstP("TBK_NUMERO_CUOTAS")
			'TBK_TASA_INTERES_MAX = rstP("TBK_TASA_INTERES_MAX")
			TBK_VCI = rstP("TBK_VCI")
			TBK_MONTO = TBK_MONTO /100
			TBK_TIPO_CUOTAS = rstP("TBK_TIPO_CUOTAS")
			TBK_NOMBRE_COMERCIO = rstP("TBK_NOMBRE_COMERCIO")
			TBK_URL_COMERCIO = rstP("TBK_URL_COMERCIO") 
			TBK_NOMBRE_COMPRADOR = rstP("TBK_NOMBRE_COMPRADOR") 
			TBK_DESCRIPCION_SERVICIO = rstP("TBK_DESCRIPCION_SERVICIO")
			TBK_NOMBRE_REEMBOLSO = rstP("TBK_NOMBRE_REEMBOLSO")
			TBK_MAIL_REEMBOLSO  = rstP("TBK_MAIL_REEMBOLSO")

			'actualizar estado de mt_webpay y llama al rebaja cuotas .exe		
			strpe ="SP_PAGO_EXITOSO_WEBPAY '"& OC &"','"& TBK_ID_TRANSACCION &"'"
			Session("Conn").execute(strpe)
		else
			'no encontro orden de compra en proceso, manda a pagina de fracaso
			Session("OC") = OC 
			response.Redirect("PagoFracasoTBK.asp")
			response.End()
		end if		
		%>
        
</center> 

<table align="center" width="619">
<hr width=98% align="center" style="border:1px solid #00C">

	<form name="form" method="post" action="grabainscripcion.asp" style="font-family:'Lucida Grande', 'Lucida Sans Unicode', 'Lucida Sans', 'DejaVu Sans', Verdana, sans-serif; font:bold" onSubmit="return validaDatos(this);" >
<tr>
<td width="104"></td>
  <td>
  Nombre de comercio</td>
  <td > 
  	:   	
	</td>
  <td align="left">
    <%=TBK_NOMBRE_COMERCIO%><br>
  </td>
</tr>
<tr>
<td width="104"></td>
  <td>
  URL de comercio</td>
  <td > 
  	:   	
	</td>
  <td align="left">
    <a href="<%=TBK_URL_COMERCIO%>"><%=TBK_URL_COMERCIO%></a><br>
  </td>
</tr>
<tr>
<td></td>
  <td>
  	Tipo Transacci&oacute;n</td>
    <td >    	
    :
	</td>
  <td>
    <%=TBK_TIPO_TRANSACCION%><br>
  </td>
</tr>
<tr>
<td width="104"></td>
  <td>
  N&uacute;mero de Orden</td>
  <td > 
  	:   	
	</td>
  <td align="left">
    <%=TBK_ORDEN_COMPRA%><br>
  </td>
</tr>
<tr>
<td></td>
  <td>
  Tipo de Pago</td>
    <td > 
    	:   	
	</td>
  <td>
    <%=TBK_TIPO_PAGO%><br> 
  </td>
</tr>
<tr>
<td></td>
  <td>
  Monto  </td>
  <td > 
  	:   	
	</td>
  <td>
    <%="$"& FormatNumber(TBK_MONTO, 0)%><br> 
  </td>
</tr>
<tr>
<td></td>
  <td>
  N&uacute;mero de Tarjeta </td>
  <td > 
  	:   	
	</td>
  <td>
    <%="************" & TBK_FINAL_NUMERO_TARJETA%><br>
  </td>
</tr>
<tr>
<td></td>
  <td>
  C&oacute;digo Autorizaci&oacute;n </td>
  <td >
  	:    	
	</td>
  <td>
  	<%=TBK_CODIGO_AUTORIZACION%><br>
  </td>
</tr>
<tr>
<td></td>
  <td width="205">  
  Tipo de Cuotas </td>
  <td width="20" > 
	  :   	
	</td>
  <td width="251">
  	<%=TBK_TIPO_CUOTAS%><br>
  </td>
</tr>
<tr>
<td></td>
  <td>
   Cuotas</td>   
  <td >
  	:    	
	</td>
  <td>
    <%=TBK_NUMERO_CUOTAS%><br>
  </td>
</tr>
<tr>
<td></td>
  <td>
  Fecha Transacci&oacute;n</td>
  <td >
  	:    	
	</td>
  <td>
  	<%=TBK_FECHA_TRANSACCION%><br> 
  </td>
</tr>
<tr>
<td></td>
  <td>
  Nombre del comprador</td>
  <td >
  	:    	
	</td>
  <td>
  	<%=TBK_NOMBRE_COMPRADOR%><br> 
  </td>
</tr>
<tr>
<td></td>
  <td>
  Descripci&oacute; Servicio</td>
  <td >
  	:    	
	</td>
  <td>
  	<%=TBK_DESCRIPCION_SERVICIO%><br> 
  </td>
</tr>
<tr>
	<td height="5" colspan="4" align="left">
	</td>
</tr>
<tr>
	<td height="5" colspan="4" align="left"><p class="Tit-celdas style2" align="justify">Devoluciones y Reembolsos<br />
    Los alumnos podran solicitar la devoluci&oacute;n del  monto cancelado dentro de los 10 d&iacute;as h&aacute;biles,&nbsp;
    a trav&eacute;s de una carta formal dirigida a la vicerrectoria zonal  correspondiente, indicando claramente el motivo de este,
    ante cualquier problema contactarse con <%=TBK_NOMBRE_REEMBOLSO%>, email : <%=TBK_MAIL_REEMBOLSO%><br /> 
    </p>	
	</td>
</tr>
<tr>
	<td align="left" colspan="3">    
		<a href="javascript: res=cerrar();"><img src="Imagenes/botones/continuar-of.gif" alt="Volver al Portal de Alumnos" name="boton_volver" width="149"
                                    height="49" border="0" id="boton_volver" onclick="return cerrar()" /></a>              
	</td>
    <td align="right" >
    <b><a href="javascript:imprime()"><img src="imag/f-der/impresion.gif" width="36" height="36" border="0" align="certer"></a></b>
    </td>
</tr>
</form>
</table>
</fieldset>
</td>
</tr>
</table>   

</html>
<script>
function imprime()
{
  parent.focus();
  parent.print(); 
}
</script>
