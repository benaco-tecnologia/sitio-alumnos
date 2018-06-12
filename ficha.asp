<%
    Response.Expires = -1
    Session.Timeout = 999
    Response.Buffer = false
    Server.ScriptTimeout = 99999
%>
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/liscarhis.inc" -->
<!--#INCLUDE file="analytics.asp" -->

<%
strParame="SELECT coalesce(dbo.Fn_ValorParame('USAANALYTICS'),'NO')Parame"
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
	USAANALYTICS=rstParame("Parame")
end if

if USAANALYTICS="SI" then
	analitycs() 
end if  
%>

<html>
<head>
<title>Ficha Curricular</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<%
   'ejecutar un proceso
   Codcli = Session("Codcli")
   CodPestud = Session("CodPestud")
   CodCarr = Session("CodCarr")
   
   'StrSql = "delete from liscarhis where  rut = '" & CodCli & "' "
   'Session("Conn").execute StrSql
   'Insertar Codcli, codPestud, CodCarr
   StrSql = "pa_Insertar '" & CodCli & "', '" & CodPestud & "','" & CodCarr & "' "
   
   'response.Write(strsql)
  ' response.End
   Session("Conn").execute StrSql
'  <input type="hidden" name="sf" value = "{liscarhis.rut}='<%=CodCli%>'">   
%>
  
  <form name="form1" method="GET" action="ra_0206a.rpt">
  <input type="hidden" name="sf" value = "{liscarhis.rut}='041SAARQD001'">
  <input type="hidden" name="user0" value = "matricula">
  <input type="hidden" name="password0" value = "dtb01s">
  
  <input type="hidden" name="user0@ra_resolhis" value = "matricula">
  <input type="hidden" name="password0@ra_resolhis" value = "dtb01s">

  <input type="hidden" name="user0@ra_hojavida" value = "matricula">
  <input type="hidden" name="password0@ra_hojavida" value = "dtb01s">

  <input type="hidden" name="user0@ra_solicitud" value = "matricula">
  <input type="hidden" name="password0@ra_solicitud" value = "dtb01s">

  <input type="hidden" name="init" value = "html_page">
  
</form>
</body>
<script>
  //alert('<%=CodPestud%>');
  document.form1.submit();

//var x;

//  x = "ra_0206.rpt?init=java&sf={liscarhis.rut}='<%=Session("CodCli")%>'";
//  x = x + "&user0=matricula&password0=dtb01s&user0@ra_resolhis=matricula";
//  x = x + "&password0@ra_resolhis=dtb01s";
//  alert(x);
 // window.location = x;
</script>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->

