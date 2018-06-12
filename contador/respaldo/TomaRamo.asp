<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#INCLUDE FILE="conexion.inc" -->
<%
	Set Conn = Nothing
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.ConnectionTimeout = 100
	ConnectStr = "PROVIDER=SQLOLEDB;Data Source="&cStr(Application("Servidor"))&";Initial Catalog="&cStr(Application("Base"))
	Conn.Open ConnectStr, cStr(Application("Usuario")), cStr(Application("password"))
%>
<% Dim var1, var2, var3, NueDiurnos_0, NueDiurnos_1, NueDiurnos_2, NueDiurnos_3, AntDiurnos_0, AntDiurnos_1, AntDiurnos_2, AntDiurnos_3, NueVesp_0, NueVesp_1, NueVesp_2, NueVesp_3, AntVesp_0 %> 
<% Dim AntVesp_1, AntVesp_2, AntVesp_3, NueTot_0, NueTot_1, NueTot_2, NueTot_3, AntTot_0, AntTot_1, AntTot_2, AntTot_3 %>
<html>
<head>
<title>::Contador::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body class="iphone">
<table border="0">
  <tr bgcolor="#000066"> 
    <td width="140"><div align="center"><a href="contador.asp"><font size="6" face="Verdana, Arial" color="#FFFFFF">hoy</font></a></div></td>
    <td width="140"><div align="center"><a href="contadorayer.asp"><font size="6" face="Verdana, Arial" color="#FFFFFF">ayer</a></font></div></td>
    <td width="140"><div align="center"><a href="TomaRamo.asp"><font size="6" face="Verdana, Arial" color="#FFFFFF">T.Ramo</font></a></font></div></td>
    <td width="140"><div align="center"><a href="retencion.asp"><font size="6" face="Verdana, Arial" color="#FFFFFF">Retenci</a></font></div></td>
    <td width="140"><div align="center"></div></td>
    <td width="140"><div align="center"></div></td>
    <td width="140"><div align="center"></div></td>
  </tr>
</table>
<br>
<font size="2" face="Verdana, Arial, Helvetica, sans-serif"> <br>
<b>Contador de Toma de Ramos Primavera 2011</b> <br>
</font>
<table>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Alumnos Conectados</font> 
    <td><div align="right"></div> 
    <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><% response.write (application("uactivos"))%></font></div></td>
  </tr>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Inscritos WEB </font> 
    <td><div align="right"></div> 
    <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% 
           set rst = createobject("adodb.recordset")
           rst.cursortype = 1
           rst.open "select count ( * ) as valor from  sis_reg_inscripcion where fecha >= '20170701' ",Conn
           if rst.recordcount > 0 then
              response.write rst("valor")
           else
              response.write "0"
           end if
           set rst=nothing
         %>
        </font></div></td>
  </tr>
  <tr> 
    <td><font size="2" face="Verdana, Arial">Solicitudes Web </font> 
    <td><div align="right"></div> 
    <td> <div align="right"><font size="2" face="Verdana, Arial"> 
        <%
            set rst = createobject("adodb.recordset")
            rst.cursortype = 1
            rst.open "select count ( * ) as valor from  sis_reg_solicitud where fecha >= '20170701' ",Conn
            if rst.recordcount > 0 then
                response.write rst("valor")
            else
               response.write "0"
            end if
            set rst=nothing
        %>
        </font></div></td>
  </tr>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Total Alumnos</font> 
    <td><div align="right"></div> 
    <td> <div align="right"><font size="2" face="Verdana, Arial"> 
        <%
            set rst = createobject("adodb.recordset")
            rst.cursortype = 1
            rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2017 and estacad='VIGENTE' and periodo_mat = 2",Conn
            if rst.recordcount > 0 then
               response.write rst("valor")
            else
               response.write "0"
            end if
            set rst=nothing
            %>
        </font></div></td>
  </tr>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Encuestas realizadas </font> 
    <td><div align="right"></div> 
    <td> <div align="right"><font size="2" face="Verdana, Arial"> 
        <%
            set rst = createobject("adodb.recordset")
            rst.cursortype = 1
            rst.open "select count ( * ) as valor from mt_alumno where ano_ed=2017 and periodo_ed=1 and encdoc='S' ",Conn
            if rst.recordcount > 0 then
               response.write rst("valor")
            else
               response.write "0"
            end if
            set rst=nothing
            %>
        </font></div></td>
  </tr>
</table>
<br>








</body>
</html>
<!--#INCLUDE file="desconexion.inc" -->
