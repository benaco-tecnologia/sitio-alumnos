<%
dim xml 
xml = Session("miXML")
ID = Session("IDPAGOWEB")
monto = Session("monto_a_pagar")

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
%>
<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css" />
<head>
    <title>Boton de Pago Servipag</title>
</head>
<body>
    <form method="post" action="http://164.77.87.90/BP2/pivote.aspx">
      <div align="center">
        <p>
          <% if xml="" then response.write("Su sesi&oacuten ha expirado.  Deber&aacute volver a loguearse antes de continuar.") end if %>
          <textarea style="width: 1; height: 1" name="xml"> <%=xml%>          </textarea>
            <textarea style="width: 1; height: 1" name="IDPAGOWEB"> <%=ID%>             </textarea>
            <textarea style="width: 1; height: 1" name="monto_a_pagar"> <%=monto%>             </textarea>
            <textarea style="width: 1; height: 1" name="NombreCliente"> <%=NombreCliente%>             </textarea>
            <textarea style="width: 1; height: 1" name="RutCliente"> <%=RutCliente%>             </textarea>
            <textarea style="width: 1; height: 1" name="EmailCliente"> <%=EmailCliente%>             </textarea>
            

            
          <br />
          <img src="images/T-confirmacion-pago.gif" width="257" height="38"><br />
        </p>
      </div>
      <table width="566" border="0" align="center" cellpadding="3" cellspacing="2">
<tr>
            <td width="88" height="25" background="Imagenes/fdo-cabecera-cel.jpg" class="text-cabecera-celda">
                <strong class="text-cabecera-celda">NOMBRE : </strong>            </td>
      <td width="460" height="25" bgcolor="#9FDBE8" class="tex-totales-celda">
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
    </table>
<center>
<br />
        <br />
        <input type="image" src="Imagenes/botones/continuar-of.gif"></center>
    </form>
</body>
</html>
