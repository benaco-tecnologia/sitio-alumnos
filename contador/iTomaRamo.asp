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
    <td width="140"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif"><a href="icontador.asp"><font color="#FFFFFF">hoy</font></a></font></div></td>
    <td width="140"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif"><a href="icontadorayer.asp"><font color="#FFFFFF"> 
        ayer</font></a></font></div></td>
    <td width="140"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif"><a href="iTomaRamo.asp"><font color="#FFFFFF"> 
        T.Ramo</font></a></font></div></td>
    <td width="140"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif"><a href="iretencion.asp"><font color="#FFFFFF"> 
        Retenci </font></a></font></div></td>
    <td width="140"><div align="center"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif" size="6"></font></div></td>
    <td width="140"><div align="center"><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif" size="6"></font></div></td>
    <td width="140"><div align="center"><font color="#0033FF" size="6" face="Verdana, Arial, Helvetica, sans-serif"> 
        </font><font color="#0033FF" size="6" face="Verdana, Arial, Helvetica, sans-serif"> 
        </font><font color="#000000" size="6" face="Verdana, Arial, Helvetica, sans-serif"> 
        </font><font color="#FFFFFF" face="Verdana, Arial, Helvetica, sans-serif" size="6"></font></div></td>
  </tr>
</table>
<br>
<font size="2" face="Verdana, Arial, Helvetica, sans-serif"> <br>
<b>Contador de Toma de Ramos Primavera 2011</b> <br>
</font>
<table>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Alumnos Conectados 
      </font> 
    <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font> 
      </div> 
    <td> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% response.write (application("uactivos"))%>
        </font></div></td>
  </tr>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Inscritos 
      WEB </font> 
    <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font> 
      </div> 
    <td> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count ( * ) as valor from  sis_reg_inscripcion where fecha >= '20120101' ",Conn
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
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Solicitudes 
      Web </font> 
    <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font> 
      </div> 
    <td> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count ( * ) as valor from  sis_reg_solicitud where fecha >= '20120101' ",Conn
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
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Total Inscripción 
      </font> 
    <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font> 
      </div> 
    <td> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1",Conn
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
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Encuestas 
      realizadas </font> 
    <td><div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font> 
      </div> 
    <td> <div align="right"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count ( * ) as valor from mt_alumno where ano_ed=2011 and periodo_ed=2 and encdoc='S' ",Conn
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
<font size="2" face="Verdana, Arial, Helvetica, sans-serif">
<table border="0">
  <tr> 
    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Contador 
      Matrícula</font></strong></td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td width="150"><strong></strong></td>
    <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">2012</font></div></td>
    <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">2011</font></div></td>
    <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">2011</font></div></td>
    <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">2010</font></div></td>
  </tr>
  <tr> 
    <td width="150"><strong><font color="#0033FF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Nuevos</font></strong></td>
    <td width="50"><div align="center"><font color="#0033FF" size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012",Conn
if rst.recordcount > 0 then
   NueTot_0 = CInt(rst("valor"))
   response.write rst("valor")
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
        </font></div></td>
    <td width="50"><div align="center"><font color="#0033FF" size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2011",Conn
if rst.recordcount > 0 then
   NueTot_0 = CInt(rst("valor"))
   response.write rst("valor")
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
    </font></div></td>
    <td width="50"><div align="center"><font color="#0033FF" size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select NueDiurnos_tot+NueVesp_tot as valor from rc_contador WITH (NOLOCK) where fecha=CONVERT(VARCHAR(11), DATEADD(dd, - 364, GETDATE()))",Conn
if rst.recordcount > 0 then
   NueTot_1 = CInt(rst("valor"))
   response.write rst("valor")
else
   var2 = 0
   response.write "0"
end if

set rst=nothing

%>
    </font></div></td>
    <td width="50" align="center"><font color="#0033FF" size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select NueDiurnos_tot+NueVesp_tot as valor from rc_contador WITH (NOLOCK) where fecha=CONVERT(VARCHAR(11), DATEADD(dd, - 728, GETDATE()))",Conn
if rst.recordcount > 0 then
   NueTot_2 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%>
    </font></td>
  </tr>
  <tr> 
    <td width="150"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Diurnos</font></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012 and jornada='D'",Conn
if rst.recordcount > 0 then
   NueDiurnos_0 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%>
        </font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2011 and estacad='VIGENTE' and periodo_mat = 1 and ano =2011 and jornada='D'",Conn
if rst.recordcount > 0 then
   NueDiurnos_0 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%>
    </font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select NueDiurnos_tot as valor from rc_contador WITH (NOLOCK) where fecha=CONVERT(VARCHAR(11), DATEADD(dd, - 364, GETDATE()))",Conn
if rst.recordcount > 0 then
   NueDiurnos_1 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "s.i."
end if

set rst=nothing

%>
    </font></div></td>
    <td width="50" align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select NueDiurnos_tot as valor from rc_contador WITH (NOLOCK) where fecha= CONVERT(VARCHAR(11), DATEADD(dd, - 728, GETDATE()))",Conn
if rst.recordcount > 0 then
   NueDiurnos_2 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "JLAZ"
end if

set rst=nothing

%>
    </font></td>
  </tr>
  <tr> 
    <td width="150"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Vespertinos</font></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012 and jornada='V'",Conn
if rst.recordcount > 0 then
   NueVesp_0 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%>
        </font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2011 and estacad='VIGENTE' and periodo_mat = 1 and ano =2011 and jornada='V'",Conn
if rst.recordcount > 0 then
   NueVesp_0 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%>
    </font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select NueVesp_tot as valor from rc_contador WITH (NOLOCK) where fecha=CONVERT(VARCHAR(11), DATEADD(dd, - 364, GETDATE()))",Conn
if rst.recordcount > 0 then
   NueVesp_1 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%>
    </font></div></td>
    <td width="50" align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select NueVesp_tot as valor from rc_contador WITH (NOLOCK) where fecha=CONVERT(VARCHAR(11), DATEADD(dd, - 728, GETDATE()))",Conn
if rst.recordcount > 0 then
   NueVesp_2 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%>
    </font></td>
  </tr>
  <tr> 
    <td width="150"><font color="#0033FF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Antiguos</strong></font></td>
    <td width="50"><div align="center"><font color="#0033FF" size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano <2012",Conn
if rst.recordcount > 0 then
   AntTot_0 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%>
        </font></div></td>
    <td width="50"><div align="center"><font color="#0033FF" size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2011 and estacad='VIGENTE' and periodo_mat = 1 and ano <2011",Conn
if rst.recordcount > 0 then
   AntTot_0 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%>
    </font></div></td>
    <td width="50"><div align="center"><font color="#0033FF" size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select AntDiurnos_tot+AntVesp_tot as valor from rc_contador WITH (NOLOCK) where fecha=CONVERT(VARCHAR(11), DATEADD(dd, - 364, GETDATE()))",Conn
if rst.recordcount > 0 then
   AntTot_1 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%>
    </font></div></td>
    <td width="50" align="center"><font color="#0033FF" size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select AntDiurnos_tot+AntVesp_tot as valor from rc_contador WITH (NOLOCK) where fecha=CONVERT(VARCHAR(11), DATEADD(dd, - 728, GETDATE()))",Conn
if rst.recordcount > 0 then
   AntTot_2 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%>
    </font></td>
  </tr>
  <tr> 
    <td width="150"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Diurnos</font></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano <2012 and jornada='D'",Conn
if rst.recordcount > 0 then
   AntDiurnos_0 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%>
        </font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2011 and estacad='VIGENTE' and periodo_mat = 1 and ano <2011 and jornada='D'",Conn
if rst.recordcount > 0 then
   AntDiurnos_0 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%>
    </font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select AntDiurnos_tot as valor from rc_contador WITH (NOLOCK) where fecha=CONVERT(VARCHAR(11), DATEADD(dd, - 364, GETDATE()))",Conn
if rst.recordcount > 0 then
   AntDiurnos_1 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "no Enc"
end if

set rst=nothing

%>
    </font></div></td>
    <td width="50" align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select AntDiurnos_tot as valor from rc_contador WITH (NOLOCK) where fecha=CONVERT(VARCHAR(11), DATEADD(dd, - 728, GETDATE()))",Conn
if rst.recordcount > 0 then
   AntDiurnos_2 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "s.i."
end if

set rst=nothing

%>
    </font></td>
  </tr>
  <tr> 
    <td width="150"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Vespertinos</font></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
        <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano <2012 and jornada='V'",Conn
if rst.recordcount > 0 then
   AntVesp_0 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%>
        </font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2011 and estacad='VIGENTE' and periodo_mat = 1 and ano <2011 and jornada='V'",Conn
if rst.recordcount > 0 then
   AntVesp_0 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%>
    </font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select AntVesp_tot as valor from rc_contador WITH (NOLOCK) where fecha=CONVERT(VARCHAR(11), DATEADD(dd, - 364, GETDATE()))",Conn
if rst.recordcount > 0 then
   AntVesp_1 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%>
    </font></div></td>
    <td width="50" align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select AntVesp_tot as valor from rc_contador WITH (NOLOCK) where fecha=CONVERT(VARCHAR(11), DATEADD(dd, - 728, GETDATE()))",Conn
if rst.recordcount > 0 then
   AntVesp_2 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "0"
end if

set rst=nothing

%>
    </font></td>
  </tr>
  <tr> 
    <td width="150"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b style='mso-bidi-font-weight:normal'>Total 
      Matrícula :</b></font></td>
    <td width="50"><div align="center"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%

   response.write (AntVesp_0+AntDiurnos_0+NueDiurnos_0+NueVesp_0)

%>
    </font></strong></div></td>
    <td width="50"><div align="center"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%

   response.write (AntVesp_0+AntDiurnos_0+NueDiurnos_0+NueVesp_0)

%>
    </font></strong></div></td>
    <td width="50"><div align="center"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%

   response.write (AntVesp_1+AntDiurnos_1+NueDiurnos_1+NueVesp_1)

%>
    </font></strong></div></td>
    <td width="50" align="center"><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif">
      <%

   response.write (AntVesp_2+AntDiurnos_2+NueDiurnos_2+NueVesp_2)

%>
    </font></strong></td>
  </tr>
</table>
</font>
</body>
</html>
<!--#INCLUDE file="desconexion.inc" -->
