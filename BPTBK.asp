<!--#INCLUDE FILE="include/conexion.inc" --> 
<% 
dim xml 
xml = Session("miXML")
ID = Session("IDPAGOWEB")
monto = Session("monto_a_pagar")
OC = "OC_" & ID

'VALIDA SI YA INGESO
strValida="SELECT TOP 1 1 FROM MT_WEBPAY WHERE TBK_ORDEN_COMPRA ='"& OC &"'"
if BCL_ADO(strValida, rstV) then 
	response.Redirect("PagoRepiteTBK.asp")
end if

NombreCliente=trim(Right(Session("NomAlum"),50))
RutCliente=Session("RutCliente")
RutCliente = trim(Right(Replace(RutCliente, ".", ""),10))

str="select mail from mt_client where codcli='"& Session("RutCli")  &"'"
set rst = Session("Conn").Execute(str)
if not rst.eof then
	if isnull(rst("mail")) then
		EmailCliente=""
	else
		EmailCliente=trim(Right(rst("mail"),50))
	end if 
else
	EmailCliente=""
end if 

'Nuevo campo actualiza monto en mt_pago_web 
StrActMonto="UPDATE MT_PAGO_WEB SET MONTO_PAGO='"& monto &"' WHERE ID='"& ID &"'" 
Session("Conn").execute(StrActMonto)

StrW="SP_INSERTA_WEBPAY '"& OC &"','"& session("codcli") &"'"
Session("Conn").execute(StrW)

strParame="SELECT dbo.Fn_ValorParame('TBK_URL_EXITO')TBK_URL_EXITO,dbo.Fn_ValorParame('TBK_URL_FRACASO')TBK_URL_FRACASO,dbo.Fn_ValorParame('TBK_URL_PROCESA')TBK_URL_PROCESA"
if BCL_ADO(strParame, rstParame) then
	TBK_URL_EXITO = trim(valnulo(rstParame("TBK_URL_EXITO"),STR_))
	TBK_URL_FRACASO = trim(valnulo(rstParame("TBK_URL_FRACASO"),STR_)) 
	TBK_URL_PROCESA = trim(valnulo(rstParame("TBK_URL_PROCESA"),STR_)) 
end if
%>
<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css" />
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1"> 
<head>
    <title>Boton de Pago Servipag</title>
</head>
<body>
    <form method="post" action="<%=TBK_URL_PROCESA%>">
      <div align="center">
        <p>
          <% if xml="" then response.write("Su sesi&oacuten ha expirado.  Deber&aacute volver a loguearse antes de continuar.") end if %>
            
          <br />
          <img src="images/T-confirmacion-pago.gif" width="257" height="38"><br />
        </p>
      </div>
      <table width="566" border="0" align="center" cellpadding="3" cellspacing="2">
<tr>
            <td width="127" height="25" background="Imagenes/fdo-cabecera-cel.jpg" class="text-cabecera-celda">
                <strong class="text-cabecera-celda">NOMBRE : </strong>            </td>
      <td width="421" height="25" bgcolor="#9FDBE8" class="tex-totales-celda">
                <font face="Arial, Helvetica, sans-serif" class="tex-totales-celda">
                    <%=Session("NomAlum")%></font>
      </td>
      </tr>
        <tr>
            <td height="25" background="Imagenes/fdo-cabecera-cel.jpg" bgcolor="#D7EAFD" class="text-cabecera-celda">
                <strong class="text-cabecera-celda">RUT : </strong>
            </td>
            <td height="25" bgcolor="#9FDBE8" class="tex-totales-celda">
                <%=Session("RutCliente")%>
            </td>
        </tr>
        <tr>
            <td background="Imagenes/fdo-cabecera-cel.jpg" bgcolor="#D7EAFD" class="text-cabecera-celda">
                <strong class="text-cabecera-celda">MONTO: </strong>
            </td>
            <td height="25" bgcolor="#9FDBE8" class="tex-totales-celda">
                <%=FormatNumber(Session("monto_a_pagar"), 0)%>
            </td>
        </tr>
        <tr>
            <td background="Imagenes/fdo-cabecera-cel.jpg" bgcolor="#D7EAFD" class="text-cabecera-celda">
                <strong class="text-cabecera-celda">ORDEN DE COMPRA: </strong>
            </td>
            <td height="25" bgcolor="#9FDBE8" class="tex-totales-celda">
                <%=OC%>
            </td>
        </tr>
    </table>
<center>
<br /> 
        <br />   
        
        <%
		monto = Session("monto_a_pagar")
		monto = monto * 100
		%>
        <INPUT TYPE="HIDDEN" NAME="TBK_MONTO" VALUE="<%=monto%>">
        <INPUT TYPE="HIDDEN" NAME="TBK_TIPO_TRANSACCION" VALUE="TR_NORMAL"> <BR>
		<INPUT TYPE="HIDDEN" NAME="TBK_ORDEN_COMPRA" VALUE="<%=OC%>"> <BR>   
        <INPUT TYPE="HIDDEN" NAME="TBK_ID_SESION" VALUE="<%=OC%>"> <BR>   
        <INPUT TYPE="HIDDEN" NAME="TBK_URL_EXITO" SIZE=”40” VALUE="<%=TBK_URL_EXITO%>" SIZE="50">
        <INPUT TYPE="HIDDEN" NAME="TBK_URL_FRACASO" SIZE=40 VALUE="<%=TBK_URL_FRACASO%>" SIZE=”50”>
        
        <% if xml<>"" then%>     
	        <input type="image" src="Imagenes/botones/continuar-of.gif"></center>
        <%end if %>     
        
    </form>
</body>
</html>
