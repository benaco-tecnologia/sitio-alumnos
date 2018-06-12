<%
    Response.Expires = -1
    Session.Timeout = 999
    Response.Buffer = false
    Server.ScriptTimeout = 99999
%>
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/liscarhis.inc" -->
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000">
<%
   'ejecutar un proceso
   Codcli = Session("Codcli")
   CodPestud = Session("CodPestud")
   CodCarr = Session("CodCarr")
   ano=Session("ano")
   
   Dim strsql, rst
   
   StrSql = "Select regimen from Mt_carrer where CodCarr = '" & CodCarr & "' "
   if BCL_ADO(strsql, rst) then
     Regimen = Rst("REGIMEN")
   else
     Regimen = ""
   end if
'    report.StoredProcParam(0) = Trim(txtcodpestud.Text)
'    report.StoredProcParam(1) = Trim(txtCodigo.Text)
'    report.StoredProcParam(2) = Trim(IIf(Val(CbRamos.ListIndex) = 0, "A", "N"))
'    report.StoredProcParam(3) = Trim(IIf(Val(CbDetalle.ListIndex) = 0, "S", "N"))
'    report.StoredProcParam(4) = Trim(TxtAno.Text)
'    strTMP = "select REGIMEN from MT_CARRER WHERE CODCARR ='" & Trim(txtCodcarr.Text) & "'"
'    If BuscaCodigoLectura(strTMP, RstTmp) Then
'        Regimen = Trim(ValNulo(RstTmp("regimen"), STR_))
'    Else
'        Regimen = "SEMESTRAL"
'    End If
'    If Regimen = "ANUAL" Then
'        report.StoredProcParam(5) = "1"
'        report.StoredProcParam(6) = "3"
'    Else
'        report.StoredProcParam(5) = Trim(TxtPeriodo.Text)
'        report.StoredProcParam(6) = Trim(txtPeriodoFin.Text)
'    End If
'    report.StoredProcParam(7) = Trim(txtCodcarr.Text)

   
%>
<form name="form1" method="GET" action="ra0415.rpt">
  <input type="hidden" name="init" value = "html_page">
  <input type="hidden" name="user0" value = "matricula">
  <input type="hidden" name="password0" value = "dtb01s">
  <input type="hidden" name="prompt0" value = "<%=CodPestud%>">
  <input type="hidden" name="prompt1" value = "<%=Codcli%>">
  <input type="hidden" name="prompt2" value = "N">
  <input type="hidden" name="prompt3" value = "S">
  <input type="hidden" name="prompt4" value = "<%=ano%>">
  <%if Regimen = "ANUAL" then %>
    <input type="hidden" name="prompt5" value = "1">
    <input type="hidden" name="prompt6" value = "3">     
  <%else %>
    <input type="hidden" name="prompt5" value = "<%=Periodo%>">
    <input type="hidden" name="prompt6" value = "<%=Periodo%>">     
  <%end if%> 
  <input type="hidden" name="prompt7" value = "<%=CodCarr%>">
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

