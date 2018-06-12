<%  response.buffer = false 
    Response.Expires = -1 
%>
<!-- saved from url=(0022)http://internet.e-mail -->
<!--<OBJECT RUNAT=server PROGID=ADODB.Recordset id=malla name=malla></OBJECT> -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<SCRIPT LANGUAGE="JavaScript" >
var total=0;

function manda_cero( ) {
         document.getElementById('id_TBK_RPTA').value = 0 ;
         formTBK.submit();
 	
}

function manda_error( ) {
         document.getElementById('id_TBK_RPTA').value = 3456789 ;
         formTBK.submit();
}

</script>
<!--<OBJECT RUNAT=server PROGID=ADODB.Recordset id=Detalle name=Detalle > </OBJECT>-->
<%

' PagoTabla.asp
' Descripcion: Suma campos y genera parametros para envio a WebPay
'
' Inputs via Form POST
'
' FILAS:	Numero de filas de la tabla de costos
' pagarX:	PAGAR si esa fila se pagara o vacio si no
' cuotaX:	valor en pesos sin decimales de la cuota a pagar
' TBK_ORDEN_COMPRA:	id o numero de la orden de compra para webpay
'
' Editar funcion guarda_orden para integrar con sistema actual
'
' Basado en ejemplo webpay, modificado por Eduardo Kaftanski (ekaftan@gmail.com)
'
OC=Request.Form("OC")
montoacancelar=Request.Form("cuota1")

RESPUESTA = Request.Form("TBK_RESPUESTA")
ID_SESION = Request.Form("TBK_ID_SESION")
ORDEN_COMPRA = Request.Form("TBK_ORDEN_COMPRA")

'********************     ACTION="/WEBPAY/TBK_BP_PAGO.CGI"
%>



<HTML> <HEAD> <TITLE>YO SOY TRANSBANK</TITLE> </HEAD>
<BODY BGCOLOR="#3069C6" TOPMARGIN="10" LEFTMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
<BR> <P ALIGN="CENTER">
<FONT FACE=ARIAL SIZE="5" COLOR="WHITE"> YO SOY TRANSBANK </FONT>
</P> <BR>
<FORM name="formTBK" METHOD="POST" ACTION="PagoWebRegistro.asp">
<TABLE BORDER="0" ALIGN="CENTER">
<TR>
<TD ALIGN="CENTER">
<FONT FACE=ARIAL SIZE="3" COLOR="WHITE">MONTO TRANSACCIÓN</FONT> <BR>
<INPUT TYPE="HIDDEN" NAME="TBK_MONTO" VALUE="<%= montoacancelar * 100%>"><%= montoacancelar %> <BR>
</TD>
<TD ALIGN="CENTER">
<FONT FACE="ARIAL" SIZE="3" COLOR="WHITE">Nº DE ORDEN</FONT> <BR>
<INPUT TYPE="HIDDEN" NAME="TBK_ORDEN_COMPRA" VALUE="<%=OC%>"> <%=OC%> <BR>
</TD>

<TD ALIGN="CENTER"> <BR>
<INPUT TYPE="button" NAME="botonexito" id="bexito" value="CASOEXITO" onClick="Javascript:manda_cero()">
<BR>
</TD>

<TD ALIGN="CENTER"> <BR>
<INPUT TYPE="button" NAME="botonfracaso" id="bfracaso" value="CASOFRACASO" onClick="Javascript:manda_error()">
 <BR>
</TD>
<INPUT TYPE="HIDDEN" NAME="TBK_RESPUESTA" id="id_TBK_RPTA"  > 
</TR>
</TABLE>

</FORM>
</BODY>
</HTML>
