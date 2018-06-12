<%  response.buffer = false 
    Response.Expires = -1
%>
<!-- saved from url=(0022)http://internet.e-mail -->
<!--<OBJECT RUNAT=server PROGID=ADODB.Recordset id=malla name=malla></OBJECT> -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/NombreRamo.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->

<%
dim strsql,strsql2
codcli=session("codcli")

RutPago = Request("rutcli")
NumeroDoc = Request("nrodoc")
DocumPago = request("docum")

ano=session("ano")
Periodo=session("periodo")
CodSede=session("codsede")

strsql =  "Select dep.*, pag.nombre,ban.descripcion as banco from matricula.mt_ctadep dep, matricula.mt_docpag pag, matricula.mt_banco ban"
strsql =  strsql + " Where dep.CTADOC    = " + NumeroDoc
strsql =  strsql + " and dep.CTADOCNUM   = " +  DocumPago
strsql =  strsql + " and dep.CTAPAG = pag.tipodoc "
strsql =  strsql + " and dep.codban *= ban.codban "
strsql =  strsql + " and dep.codcli = '" + RutPago + "'"

'response.write strsql  
'response.end()

'strsql=strsql & "order by fecmod "
Set Rst = Conn.Execute(StrSql)


'                   Set Rst = Conn.Execute(StrSql)
'Do while not rst.eof
'   response.write(rst("codcli"))
'   rst.movenext
'loop
'response.end
		%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="665" cellspacing="0" cellpadding="0" height="213" border="1" bordercolor="#FFFFFF">
  <tr bgcolor="#FFFFFF"> 
    <td height="8" colspan="2"> 
      <div align="left">
        <p><font face="Verdana, Arial, Helvetica, sans-serif" color="#000000" size="1"><b>Detalle de Pagos</b></font></p>
        <table width="652" border="1">

          <tr>
            <th bgcolor="4A5DA1" scope="col"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><font color="#FFFFFF">Tipo Pago</font></font></b></th>
            <th bgcolor="4A5DA1" scope="col"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><font color="#FFFFFF">Descripci&oacute;n</font></font></b></th>
            <th bgcolor="4A5DA1" scope="col"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><font color="#FFFFFF">N&uacute;mero</font></font></b></th>
            <th bgcolor="4A5DA1" scope="col"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><font color="#FFFFFF">Banco</font></font></b></th>
            <th bgcolor="4A5DA1" scope="col"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><font color="#FFFFFF">Fecha</font></font></b></th>
            <th bgcolor="4A5DA1" scope="col"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><font color="#FFFFFF">Monto</font></font></b></th>
            <th bgcolor="4A5DA1" scope="col"><b><font color="#FFFFFF"><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Usuario</font></font></b></th>
            <th bgcolor="4A5DA1" scope="col"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><font color="#FFFFFF">Lugar de Pago</font></font></b></th>
            <th bgcolor="4A5DA1" scope="col"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><font color="#FFFFFF">Oficina de Pago</font></font></b></th>
          </tr>
         
          <%do while not rst.eof %>
          <% 
          if not rst.eof then
			  tipopago = rst("ctapag")
			  descripcion = rst("nombre")
			  numero = rst("ctapagnum")
			  banco = rst("banco")
			  fecha = rst("fecmod")
			  monto = rst("monto")
			  usuario = rst("usuario")
			  lugarpago = "" 'rst("usuario")
			  oficinapago = ""' rst("usuario")
		  end if
		   %>
          
          <tr>
            <th bgcolor="FFC172" <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=tipopago%></th>
            <th bgcolor="FFC172" <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=descripcion%></th>
            <th bgcolor="FFC172" <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=numero%></th>
            <th bgcolor="FFC172" <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=banco%></th>
            <th bgcolor="FFC172" <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=fecha%></th>
            <th bgcolor="FFC172" <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=monto%></th>
            <th bgcolor="FFC172" <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=usuario%></th>
            <th bgcolor="FFC172" <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=lugarpago%></th>
            <th bgcolor="FFC172" <div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=oficinapago%></th>
            <%  rst.movenext %>
          </tr>
          <% loop%>
        </table>
        <p>&nbsp;</p>
        <p>&nbsp;</p>
      </div>
    </td>
  </tr>

</table>
</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->