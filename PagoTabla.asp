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



function guarda_orden () 

OC = Request.Form("TBK_ORDEN_COMPRA")

carpetaLogs = "c:\inetpub\wwwroot\webpay\log\"
'carpetaLogs = "c:\"
archivo_temporal = carpetaLogs & "Datos_" & OC & ".txt"
response.Write("***********")
response.Write(archivo_temporal)
response.Write("***********")
set filesys = CreateObject("Scripting.FileSystemObject")

'response.Write("***********")
'response.Write(archivo_temporal)
'response.Write("***********")
'response.end


set file = filesys.CreateTextFile(archivo_temporal)
file.write(Request.Form("TBK_ORDEN_COMPRA"))

end function

guarda_orden

dim i
dim t

for i = 1 To Request.Form("FILAS") Step 1
	response.write("pagar" &  cstr(i) & Request.Form("pagar"&cstr(i)) & "cuota" &  cstr(i) &Request.Form("cuota"&cstr(i))&"<br>")

	if Request.Form("pagar" &  cstr(i)) = "PAGAR" then
	response.write("sumando "&Request.Form("cuota"&cstr(i))&"<br>")
	t = t + Clng(Request.Form("cuota"&cstr(i)))
	end if
next
response.write("total "& cstr(t) &"<br>")

%>



<HTML> <HEAD> <TITLE>TIENDA ASP - PASO 2</TITLE> </HEAD>
<BODY BGCOLOR="#3069C6" TOPMARGIN="10" LEFTMARGIN="0" MARGINWIDTH="0" MARGINHEIGHT="0">
<BR> <P ALIGN="CENTER">
<FONT FACE=ARIAL SIZE="5" COLOR="WHITE"> TRANSACCION NORMAL - WINDOWS ASP - PASO 2</FONT>
</P> <BR>
<FORM METHOD="POST" ACTION="/WEBPAY/TBK_BP_PAGO.CGI">
<TABLE BORDER="0" ALIGN="CENTER">
<TR>
<TD ALIGN="CENTER">
<FONT FACE=ARIAL SIZE="3" COLOR="WHITE">MONTO TRANSACCIÓN</FONT> <BR>
<INPUT TYPE="HIDDEN" NAME="TBK_MONTO" VALUE="<%= t * 100%>"><%= t %> <BR>
</TD>
<TD ALIGN="CENTER"> <BR>
<INPUT TYPE="HIDDEN" NAME="TBK_TIPO_TRANSACCION" VALUE="TR_NORMAL"> <BR>
</TD>
</TR>
<TR>
<TD ALIGN="CENTER">
<FONT FACE="ARIAL" SIZE="3" COLOR="WHITE">Nº DE ORDEN</FONT> <BR>
<INPUT TYPE="HIDDEN" NAME="TBK_ORDEN_COMPRA" VALUE="<%=Request.Form("TBK_ORDEN_COMPRA")%>"> <%=Request.Form("TBK_ORDEN_COMPRA")%> <BR>
</TD>
<TD ALIGN="CENTER"> <BR>
<INPUT TYPE="HIDDEN" NAME="TBK_ID_SESION" VALUE="<%=OC%>"> <BR>
</TD>
</TR>
</TABLE>
<TABLE BORDER=0 ALIGN="CENTER">
<TR>
<TD ALIGN="CENTER"> <BR>
<INPUT TYPE="HIDDEN" NAME="TBK_URL_EXITO" SIZE=40
VALUE="HTTP://webpay.nn.cl/EXITO.ASP" SIZE="50"> <BR>
</TD>
<TD ALIGN="CENTER"> <BR>
<INPUT TYPE="HIDDEN" NAME="TBK_URL_FRACASO" SIZE=40
VALUE="HTTP://webpay.nn.cl/FRACASO.ASP" SIZE=50> <BR>
</TD>
</TR>
</TABLE>
<TABLE BORDER="0" ALIGN="CENTER">
<TR> <TD ALIGN="CENTER"> <BR>
<INPUT TYPE="SUBMIT" VALUE="PAGAR CON TARJETA DE CRÉDITO" SIZE=20> </BR>
</TD>
</TR>
</TABLE>
</FORM>
</BODY>
</HTML>
