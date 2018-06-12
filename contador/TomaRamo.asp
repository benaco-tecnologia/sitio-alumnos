<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#INCLUDE FILE="conexion.inc" -->
<%
	Set Conn = Nothing
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.ConnectionTimeout = 100
	'ConnectStr = "PROVIDER=SQLOLEDB;Data Source="&cStr(Application("Servidor"))&";Initial Catalog="&cStr(Application("Base_new"))
	
	ConnectStr = "PROVIDER=SQLOLEDB;Data Source="&cStr(Application("Servidor"))&";Initial Catalog=umasnet; UID=sa; PWD=reflex"
	Conn.Open ConnectStr
%>
<% Dim InsWeb, InsTot, SolWeb, TotMat, NoMat %> 
<html>
<head>
<title>:: T.Ramo ::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<table border="0">
  <tr bgcolor="#000066"> 
    <td width="65"><div align="center"><a href="contador.asp"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF">hoy</a></font></div></td>
    <td width="65"><div align="center"><a href="contadorayer.asp"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"color="#FFFFFF">ayer</font></a></div></td>
    <td width="65"><div align="center"><a href="TomaRamo.asp"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"color="#FFFFFF">T.Ramo</font></a></div></td>
    <td width="65"><div align="center"><a href="retencion.asp"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"color="#FFFFFF">Retenci&oacute;n</font></a></div></td>
    <td width="65"><div align="center"><a href="contadorlargo.asp"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"color="#FFFFFF">Largo</font></a></div></td>
    <td width="65"><div align="center"></div></td>
  </tr>
</table>
<p><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> <br>
<b>Contador de Toma de Ramos Otono 2018</b> </font></p>
 
<table  border="0">
  <tr>
    <td width="260" ><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Alumnos Conectados</font></td>
    <td width="40"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" align="right">
      <% response.write (application("auxactivos"))%>
    </font></td>
  </tr>
  <tr>
    <td ><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Inscritos WEB </font></td>
    <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" >
      <% 
           set rst = createobject("adodb.recordset")
           rst.cursortype = 1
           rst.open "select count ( distinct codcli ) as InsWeb from sis_reg_inscripcion where fecha >= '20180123' ",Conn
           if rst.recordcount > 0 then
              response.write rst("InsWeb")
           else
              response.write "0"
           end if
           set rst=nothing
         %>
    </font></td>
  </tr>  
  <tr>
	<td><font size="2" face="Verdana, Arial">Solicitudes Web </font></td>
    <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" >
      <%
            set rst = createobject("adodb.recordset")
            rst.cursortype = 1
            rst.open "select count ( distinct codcli ) as SolWeb from sis_reg_solicitud where fecha >= '20180123' ",Conn
            if rst.recordcount > 0 then
                response.write rst("SolWeb")
            else
               response.write "0"
            end if
            set rst=nothing
        %>
    </font></td>
  </tr>
  <tr>
    <td ><font size="2" face="Verdana, Arial">Inscritos total:</font></td>
    <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" >
      <%
            set rst = createobject("adodb.recordset")
            rst.cursortype = 1
            rst.open "select count ( distinct codcli ) as InsTot from ra_nota where ano='2018' ",Conn
            if rst.recordcount > 0 then
                response.write rst("InsTot")
            else
               response.write "0"
            end if
            set rst=nothing
        %>
    </font></td>
  </tr>
  <tr>	
    <td><font size="2" face="Verdana, Arial">.</td>
    <td> </td>
  </tr>
  <tr>	
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Encuestas realizadas (matriculados)</font></td>
    <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" >
      <%
            set rst = createobject("adodb.recordset")
            rst.cursortype = 1
            rst.open "select count ( distinct codcli ) as valor from ra_encuestados where ano=2017 AND periodo=2 AND evaluado='SI' AND CodCli IN (sELECT CodCli FROM MT_ALUMNO WHERE  ANO_MAT=2018)",Conn
			if rst.recordcount > 0 then
               response.write rst("valor")
            else
               response.write "0"
            end if
            set rst=nothing
            %>
    </font></td>
  </tr>
  <tr>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> Alumnos Matriculados 2018-01</font></td>
    <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" >
      <%
            set rst = createobject("adodb.recordset")
            rst.cursortype = 1
            rst.open "select count(distinct codcli) as TotMat from mt_alumno WITH (NOLOCK) where ano_mat=2018 and estacad='VIGENTE' and periodo_mat = 1 and ano <2018 ",Conn
            if rst.recordcount > 0 then
               response.write rst("TotMat")
            else
               response.write "0"
            end if
            set rst=nothing
            %>
    </font></td>
  </tr>
  <tr>	
    <td><font size="2" face="Verdana, Arial">.</td>
    <td> </td>
  </tr>
  <tr>
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Alumnos no Matriculados 2017-01</font></td>
    <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" >
      <%
            set rst = createobject("adodb.recordset")
            rst.cursortype = 1
            rst.open "select count( distinct codcli) as NoMat from mt_alumno WITH (NOLOCK) where ano_mat=2017 and estacad='VIGENTE' and periodo_mat = 2",Conn
            if rst.recordcount > 0 then
               response.write rst("NoMat")
            else
               response.write "0"
            end if
            set rst=nothing
            %>
    </font></td>
  </tr>
    <tr>	
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Encuestas realizadas</font></td>
    <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif" >
      <%
            set rst = createobject("adodb.recordset")
            rst.cursortype = 1
            rst.open "select count ( distinct codcli ) as valor from ra_encuestados where ano=2017 AND periodo=2 AND evaluado='SI' ",Conn
			if rst.recordcount > 0 then
               response.write rst("valor")
            else
               response.write "0"
            end if
            set rst=nothing
            %>
    </font></td>
  </tr>

</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
</body>
</html>
<!--#INCLUDE file="desconexion.inc" -->

