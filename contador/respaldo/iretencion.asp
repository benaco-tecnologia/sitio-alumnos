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
<p><br> 
  <font size="2" face="Verdana, Arial, Helvetica, sans-serif">
  </font></p>
<table border="0">
  <tr>
    <td colspan="7"><strong><font size="6" face="Verdana, Arial, Helvetica, sans-serif">Contador 
      Matr&iacute;cula</font></strong>
      <div align="center"></div>
      <div align="center"></div></td>
  </tr>
  <tr>
    <td width="254"><strong></strong></td>
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">2012</font></div></td>
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">2011</font></div></td>
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">2010</font></div></td>
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">2009</font></div></td>
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">%</font></div></td>
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">Prom</font></div></td>
  </tr>
  <tr>
    <td width="254" bgcolor="#FFFFCC"><strong><font color="#0033FF" size="6" face="Verdana, Arial, Helvetica, sans-serif">Nuevos</font></strong></td>
    <td width="100" bgcolor="#FFFFCC"><div align="center"><font color="#0033FF" size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012 ",Conn
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
    <td width="100" bgcolor="#FFFFCC"><div align="center"><font color="#0033FF" size="6" face="Verdana, Arial, Helvetica, sans-serif">
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
    <td width="100" bgcolor="#FFFFCC"><div align="center"><font color="#0033FF" size="6" face="Verdana, Arial, Helvetica, sans-serif">
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
    </font></div></td>
    <td width="100" bgcolor="#FFFFCC"><div align="center"><font color="#0033FF" size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select NueDiurnos_tot+NueVesp_tot as valor from rc_contador WITH (NOLOCK) where fecha=CONVERT(VARCHAR(11), DATEADD(dd, - 1092, GETDATE()))",Conn
if rst.recordcount > 0 then
   NueTot_3 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "s.i."
end if

set rst=nothing

%>
    </font></div></td>
    <td width="100" bgcolor="#FFFFCC"><div align="center"><font color="#0033FF" size="4" face="Verdana, Arial, Helvetica, sans-serif">
      <% 
		if(NueTot_1 = 0) then
			response.write "s.i."
		else
		   Response.write(FormatNumber(((NueTot_0 / NueTot_1) - 1)*100, 1))
		end if
	%>
      %</font></div></td>
    <td width="100" bgcolor="#FFFFCC"><div align="center"><font color="#0033FF" size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <% 
		if(NueTot_1 = 0) then
			response.write "s.i."
		else
		   Response.write(FormatNumber(((NueTot_1 + NueTot_2 + NueTot_3))/3, 0))
		end if
	%>
    </font></div></td>
  </tr>
  <tr>
    <td width="254"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">Diurnos</font></td>
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
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
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
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
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
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
    </font></div></td>
    <td width="100"><div align="center"><font color="#000000" size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select NueDiurnos_tot as valor from rc_contador WITH (NOLOCK) where fecha=CONVERT(VARCHAR(11), DATEADD(dd, - 1092, GETDATE()))",Conn
if rst.recordcount > 0 then
   NueDiurnos_3 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "s.i."
end if

set rst=nothing

%>
    </font></div></td>
    <td width="100"><div align="center"><font color="#333333" size="4" face="Verdana, Arial, Helvetica, sans-serif">
      <% 
		if(NueDiurnos_1 = 0) then
			response.write "s.i."
		else
		   Response.write(FormatNumber(((NueDiurnos_0 / NueDiurnos_1) - 1)*100, 1))
		end if
	%>
      %</font></div></td>
    <td width="100"><div align="center"><font color="#333333" size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <% 
		if(NueDiurnos_1 = 0) then
			response.write "s.i."
		else
		   Response.write(FormatNumber(((NueDiurnos_1 + NueDiurnos_2 + NueDiurnos_3))/3, 0))
		end if
	%>
    </font></div></td>
  </tr>
  <tr>
    <td width="254"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">Vespertinos</font></td>
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
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
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
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
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
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
    </font></div></td>
    <td width="100"><div align="center"><font color="#000000" size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select NueVesp_tot as valor from rc_contador WITH (NOLOCK) where fecha=CONVERT(VARCHAR(11), DATEADD(dd, - 1092, GETDATE()))",Conn
if rst.recordcount > 0 then
   NueVesp_3 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "s.i."
end if

set rst=nothing

%>
    </font></div></td>
    <td width="100"><div align="center"><font color="#333333" size="4" face="Verdana, Arial, Helvetica, sans-serif">
      <% 
		if(NueVesp_1 = 0) then
			response.write "s.i."
		else
		   Response.write(FormatNumber(((NueVesp_0 / NueVesp_1) - 1)*100, 1))
		end if
	%>
      %</font></div></td>
    <td width="100"><div align="center"><font color="#333333" size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <% 
		if(NueDiurnos_1 = 0) then
			response.write "s.i."
		else
		   Response.write(FormatNumber(((NueVesp_1 + NueVesp_2 + NueVesp_3))/3, 0))
		end if
	%>
    </font></div></td>
  </tr>
  <tr>
    <td width="254" bgcolor="#FFFFCC"><font color="#0033FF" size="6" face="Verdana, Arial, Helvetica, sans-serif"><strong>Antiguos</strong></font></td>
    <td width="100" bgcolor="#FFFFCC"><div align="center"><font color="#0033FF" size="6" face="Verdana, Arial, Helvetica, sans-serif">
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
    <td width="100" bgcolor="#FFFFCC"><div align="center"><font color="#0033FF" size="6" face="Verdana, Arial, Helvetica, sans-serif">
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
    <td width="100" bgcolor="#FFFFCC"><div align="center"><font color="#0033FF" size="6" face="Verdana, Arial, Helvetica, sans-serif">
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
    </font></div></td>
    <td width="100" bgcolor="#FFFFCC"><div align="center"><font color="#0033FF" size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select AntDiurnos_tot+AntVesp_tot as valor from rc_contador WITH (NOLOCK) where fecha=CONVERT(VARCHAR(11), DATEADD(dd, - 1092, GETDATE()))",Conn
if rst.recordcount > 0 then
   AntTot_3 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "s.i."
end if

set rst=nothing

%>
    </font></div></td>
    <td width="100" bgcolor="#FFFFCC"><div align="center"><font color="#0033FF" size="4" face="Verdana, Arial, Helvetica, sans-serif">
      <% 
		if(AntTot_1 = 0) then
			response.write "s.i."
		else
		   Response.write(FormatNumber(((AntTot_0 / AntTot_1) - 1)*100, 1))
		end if
	%>
      %</font></div></td>
    <td width="100" bgcolor="#FFFFCC"><div align="center"><font color="#0033FF" size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <% 
		if(AntTot_1 = 0) then
			response.write "s.i."
		else
		   Response.write(FormatNumber(((AntTot_1 + AntTot_2 + AntTot_3))/3, 0))
		end if
	%>
    </font></div></td>
  </tr>
  <tr>
    <td width="254"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">Diurnos</font></td>
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
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
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
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
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
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
    </font></div></td>
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select AntDiurnos_tot as valor from rc_contador WITH (NOLOCK) where fecha=CONVERT(VARCHAR(11), DATEADD(dd, - 1092, GETDATE()))",Conn
if rst.recordcount > 0 then
   AntDiurnos_3 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "s.i."
end if

set rst=nothing

%>
    </font></div></td>
    <td width="100"><div align="center"><font color="#333333" size="4" face="Verdana, Arial, Helvetica, sans-serif">
      <% 
		if(AntDiurnos_1 = 0) then
			response.write "s.i."
		else
		   Response.write(FormatNumber(((AntDiurnos_0 / AntDiurnos_1) - 1)*100, 1))
		end if
	%>
      %</font></div></td>
    <td width="100"><div align="center"><font color="#333333" size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <% 
		if(AntDiurnos_1 = 0) then
			response.write "s.i."
		else
		   Response.write(FormatNumber(((AntDiurnos_1 + AntDiurnos_2 + AntDiurnos_3))/3, 0))
		end if
	%>
    </font></div></td>
  </tr>
  <tr>
    <td width="254"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">Vespertinos</font></td>
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
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
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
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
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
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
    </font></div></td>
    <td width="100"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <%
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select AntVesp_tot as valor from rc_contador WITH (NOLOCK) where fecha=CONVERT(VARCHAR(11), DATEADD(dd, - 1092, GETDATE()))",Conn
if rst.recordcount > 0 then
   AntVesp_3 = CInt(rst("valor"))
   response.write rst("valor")
else
   response.write "s.i."
end if

set rst=nothing

%>
    </font></div></td>
    <td width="100"><div align="center"><font color="#333333" size="4" face="Verdana, Arial, Helvetica, sans-serif">
      <% 
		if(AntVesp_1 = 0) then
			response.write "s.i."
		else
		   Response.write(FormatNumber(((AntVesp_0 / AntVesp_1) - 1)*100, 1))
		end if
	%>
      %</font></div></td>
    <td width="100"><div align="center"><font color="#333333" face="Verdana, Arial, Helvetica, sans-serif">
      <% 
		if(AntVesp_1 = 0) then
			response.write "s.i."
		else
		   Response.write(FormatNumber(((AntVesp_1 + AntVesp_2 + AntVesp_3))/3, 0))
		end if
	%>
    </font></div></td>
  </tr>
  <tr>
    <td width="254" bgcolor="#FFFFCC"><font size="6" face="Verdana, Arial, Helvetica, sans-serif"><b style='mso-bidi-font-weight:normal'>Total 
      :</b></font></td>
    <td width="100" bgcolor="#FFFFCC"><div align="center"><strong><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <%

   response.write (AntVesp_0+AntDiurnos_0+NueDiurnos_0+NueVesp_0)

%>
    </font></strong></div></td>
    <td width="100" bgcolor="#FFFFCC"><div align="center"><strong><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <%

   response.write (AntVesp_1+AntDiurnos_1+NueDiurnos_1+NueVesp_1)

%>
    </font></strong></div></td>
    <td width="100" bgcolor="#FFFFCC"><div align="center"><strong><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <%

   response.write (AntVesp_2+AntDiurnos_2+NueDiurnos_2+NueVesp_2)

%>
    </font></strong></div></td>
    <td width="100" bgcolor="#FFFFCC"><div align="center"><strong><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <%

   response.write (AntVesp_3+AntDiurnos_3+NueDiurnos_3+NueVesp_3)

%>
    </font></strong></div></td>
    <td width="100" bgcolor="#FFFFCC"><div align="center"><font color="#333333" size="4" face="Verdana, Arial, Helvetica, sans-serif">
      <% dim var01, var02
		var01 = CInt((AntVesp_0+AntDiurnos_0+NueDiurnos_0+NueVesp_0))
		var02 = CInt((AntVesp_1+AntDiurnos_1+NueDiurnos_1+NueVesp_1))
		if(var02 = 0) then
			response.write "s.i."
		else
		   Response.write(FormatNumber(((var01 / var02) - 1)*100, 1))
		end if
	%>
      %</font></div></td>
    <td width="100" bgcolor="#FFFFCC"><div align="center"><font color="#333333" size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <% 
		if(AntTot_1 = 0) then
			response.write "s.i."
		else
		   Response.write(FormatNumber(((AntVesp_1+AntDiurnos_1+NueDiurnos_1+NueVesp_1)+(AntVesp_2+AntDiurnos_2+NueDiurnos_2+NueVesp_2)+(AntVesp_3+AntDiurnos_3+NueDiurnos_3+NueVesp_3))/3, 0))
		end if
	%>
    </font></div></td>
  </tr>
</table>
<p>&nbsp;</p>
<table border="0">
  <tr>
    <td colspan="7"><div align="center"></div>
    <div align="center"></div></td>
  </tr>
  <tr>
    <td width="228"><strong></strong></td>
    <td width="128"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">2012</font></div></td>
    <td width="80"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">2011</font></div></td>
    <td width="80"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">2010</font></div></td>
    <td width="80"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">2009</font></div></td>
    <td width="74"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">%</font></div></td>
    <td width="96"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">Prom</font></div></td>
  </tr>
  <tr>
    <td width="228" bgcolor="#FFFFCC"><strong><font color="#0033FF" size="6" face="Verdana, Arial, Helvetica, sans-serif">Nuevos</font></strong></td>
    <td width="128" bgcolor="#FFFFCC"><div align="center"><font color="#0033FF" size="6" face="Verdana, Arial, Helvetica, sans-serif">
              <% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012",Conn
if rst.recordcount > 0 then
   Response.write(FormatNumber(((nuetot_0 / 341))*100, 2))
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
    </font>%</div></td>
    <td width="80" bgcolor="#FFFFCC"><div align="center">341
      
    </div></td>
    <td width="80" bgcolor="#FFFFCC">&nbsp;</td>
    <td width="80" bgcolor="#FFFFCC">&nbsp;</td>
    <td width="74" bgcolor="#FFFFCC">&nbsp;</td>
    <td width="96" bgcolor="#FFFFCC">&nbsp;</td>
  </tr>
  <tr>
    <td width="228"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">Diurnos</font></td>
    <td width="128"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
        <% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012",Conn
if rst.recordcount > 0 then
   Response.write(FormatNumber(((nuediurnos_0 / 159))*100, 2))
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
    </font>%</div></td>
    <td width="80"><div align="center">159
      
    </div></td>
    <td width="80">&nbsp;</td>
    <td width="80">&nbsp;</td>
    <td width="74">&nbsp;</td>
    <td width="96">&nbsp;</td>
  </tr>
  <tr>
    <td width="228"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">Vespertinos</font></td>
    <td width="128"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
              <% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012",Conn
if rst.recordcount > 0 then
   Response.write(FormatNumber(((nuevesp_0 / 182))*100, 2))
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
    </font>%</div></td>
    <td width="80"><div align="center">182
      
    </div></td>
    <td width="80">&nbsp;</td>
    <td width="80">&nbsp;</td>
    <td width="74">&nbsp;</td>
    <td width="96">&nbsp;</td>
  </tr>
  <tr>
    <td width="228" bgcolor="#FFFFCC"><font color="#0033FF" size="6" face="Verdana, Arial, Helvetica, sans-serif"><strong>Antiguos</strong></font></td>
    <td width="128" bgcolor="#FFFFCC"><div align="center"><font color="#0033FF" size="6" face="Verdana, Arial, Helvetica, sans-serif">
        <% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012",Conn
if rst.recordcount > 0 then
   Response.write(FormatNumber(((anttot_0 / 951))*100, 2))
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
    </font>%</div></td>
    <td width="80" bgcolor="#FFFFCC"><div align="center">951
      
    </div></td>
    <td width="80" bgcolor="#FFFFCC">&nbsp;</td>
    <td width="80" bgcolor="#FFFFCC">&nbsp;</td>
    <td width="74" bgcolor="#FFFFCC">&nbsp;</td>
    <td width="96" bgcolor="#FFFFCC">&nbsp;</td>
  </tr>
  <tr>
    <td width="228"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">Diurnos</font></td>
    <td width="128"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012",Conn
if rst.recordcount > 0 then
   Response.write(FormatNumber(((antdiurnos_0 / 228))*100, 2))
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
    </font>%</div></td>
    <td width="80"><div align="center">228
      
    </div></td>
    <td width="80">&nbsp;</td>
    <td width="80">&nbsp;</td>
    <td width="74">&nbsp;</td>
    <td width="96">&nbsp;</td>
  </tr>
  <tr>
    <td width="228"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">Vespertinos</font></td>
    <td width="128"><div align="center"><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012",Conn
if rst.recordcount > 0 then
   Response.write(FormatNumber(((antvesp_0 / 723))*100, 2))
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
    </font>%</div></td>
    <td width="80"><div align="center">723
      
    </div></td>
    <td width="80">&nbsp;</td>
    <td width="80">&nbsp;</td>
    <td width="74">&nbsp;</td>
    <td width="96">&nbsp;</td>
  </tr>
  <tr>
    <td width="228" bgcolor="#FFFFCC"><font size="6" face="Verdana, Arial, Helvetica, sans-serif"><b style='mso-bidi-font-weight:normal'>Total 
      :</b></font></td>
    <td width="128" bgcolor="#FFFFCC"><div align="center"><strong><font size="6" face="Verdana, Arial, Helvetica, sans-serif">
      <% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012",Conn
if rst.recordcount > 0 then
   Response.write(FormatNumber((((anttot_0+nuetot_0) / 1292))*100, 2))
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
    </font></strong>%</div></td>
    <td width="80" bgcolor="#FFFFCC"><div align="center"><strong>1.292
      
    </strong></div></td>
    <td width="80" bgcolor="#FFFFCC">&nbsp;</td>
    <td width="80" bgcolor="#FFFFCC">&nbsp;</td>
    <td width="74" bgcolor="#FFFFCC">&nbsp;</td>
    <td width="96" bgcolor="#FFFFCC">&nbsp;</td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<table border="0">
  <tr> 
    <td><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> A&ntilde;o 
      Anterior</font></strong></td>
    <td width="60">&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td><strong></strong></td>
    <td width="60"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">2011</font></div></td>
    <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">2010</font></div></td>
    <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">2009</font></div></td>
    <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">2008<font size="1">*</font></font></div></td>
    <td width="60">&nbsp;</td>
    <td width="60"><div align="center"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">retenc</font></div></td>
    <td width="60">&nbsp;</td>
  </tr>
  <tr> 
    <td width="150"><font color="#0033FF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nuevos</strong></font></td>
    <td width="60"><div align="center"><font color="#0033FF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012",Conn
if rst.recordcount > 0 then
   Response.write(FormatNumber(((nuetot_0 / 341))*100, 2))
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
    % </font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">341</font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">411</font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">457</font></div></td>
    <td width="60">&nbsp;</td>
    <td width="60"><div align="center"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">ECAS 
        </font></div></td>
    <td width="60">&nbsp;</td>
  </tr>
  <tr> 
    <td width="150"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Diurnos</font></td>
    <td width="60"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012",Conn
if rst.recordcount > 0 then
   Response.write(FormatNumber(((nuediurnos_0 / 159))*100, 2))
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
        % </font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">159</font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">191</font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">202</font></div></td>
    <td width="60"><div align="center"><font color="#000000" size="2" face="Verdana, Arial, Helvetica, sans-serif">2011 
        </font></div></td>
    <td width="60"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">2010</font></div></td>
    <td width="60"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">2009</font></div></td>
  </tr>
  <tr> 
    <td width="150"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Vespertinos</font></td>
    <td width="60"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012",Conn
if rst.recordcount > 0 then
   Response.write(FormatNumber(((nuevesp_0 / 182))*100, 2))
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
        % </font><font color="#000000"></font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">182</font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">220</font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">255</font></div></td>
    <td width="60"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        </font><font color="#000000"></font></div></td>
    <td width="60">&nbsp;</td>
    <td width="60">&nbsp;</td>
  </tr>
  <tr> 
    <td width="150"><font color="#0033FF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Antiguos</strong></font></td>
    <td width="60"><div align="center"><font color="#0033FF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012",Conn
if rst.recordcount > 0 then
   Response.write(FormatNumber(((anttot_0 / 951))*100, 2))
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
        % </font><font color="#000000"></font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">951</font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">905</font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">872</font></div></td>
    <td width="60"><div align="center"><font color="#0033FF" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012",Conn
if rst.recordcount > 0 then
   Response.write(FormatNumber(((anttot_0 / (951+341-72)))*100, 2))
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
        % </font><font color="#000000"></font></div></td>
    <td width="60"><div align="center"><font color="#0033FF" size="1" face="Verdana, Arial, Helvetica, sans-serif">72.26%</font></div></td>
    <td width="60"><div align="center"><font color="#0033FF" size="1" face="Verdana, Arial, Helvetica, sans-serif">68.09%</font></div></td>
  </tr>
  <tr> 
    <td width="150"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Diurnos</font></td>
    <td width="60"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012",Conn
if rst.recordcount > 0 then
   Response.write(FormatNumber(((antdiurnos_0 / 228))*100, 2))
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
        % </font><font color="#000000"></font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">228</font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">226</font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">269</font></div></td>
    <td width="60"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012",Conn
if rst.recordcount > 0 then
   Response.write(FormatNumber(((antdiurnos_0 / (228+159)))*100, 2))
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
        % </font><font color="#000000"></font></div></td>
    <td width="60"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">54.67%</font></div></td>
    <td width="60"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">47.98%</font></div></td>
  </tr>
  <tr> 
    <td width="150"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">Vespertinos</font></td>
    <td width="60"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012",Conn
if rst.recordcount > 0 then
   Response.write(FormatNumber(((antvesp_0 / 723))*100, 2))
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
        % </font><font color="#000000"></font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">723</font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">679</font></div></td>
    <td width="50"><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">603</font></div></td>
    <td width="60"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012",Conn
if rst.recordcount > 0 then
   Response.write(FormatNumber(((antvesp_0 / (723+182-72)))*100, 2))
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
        % </font><font color="#000000"></font></div></td>
    <td width="60"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">80.42%</font></div></td>
    <td width="60"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">79.13%</font></div></td>
  </tr>
  <tr> 
    <td><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b style='mso-bidi-font-weight:normal'>Total 
      Matrícula :</b></font></td>
    <td width="60"><div align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
        <% 
set rst = createobject("adodb.recordset")
rst.cursortype = 1

rst.open "select count(codcli) as valor from mt_alumno WITH (NOLOCK) where ano_mat=2012 and estacad='VIGENTE' and periodo_mat = 1 and ano =2012",Conn
if rst.recordcount > 0 then
   Response.write(FormatNumber((((anttot_0+nuetot_0) / 1292))*100, 2))
else
   var1 = 0
   response.write "0"
end if

set rst=nothing

%>
        % </font><font color="#000000"></font><font color="#000000"></font></div></td>
    <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">1.292</font></div></td>
    <td><div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">1.316</font></div></td>
    <td> <div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif">1.329</font></div></td>
    <td width="60"> <div align="center"></div></td>
    <td width="60">&nbsp;</td>
    <td width="60">&nbsp;</td>
  </tr>
</table>
</body>
</html>
<!--#INCLUDE file="desconexion.inc" -->
