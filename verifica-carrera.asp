<!--<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">-->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--<OBJECT RUNAT=server PROGID=ADODB.Recordset id=Llamacarrera name=Llamacarrera></OBJECT>-->
<%
dim strcarrera
strcarrera = request.form("cmbCarrera")
Accionvar= request.form("Accion")
strCredito=request.form("chkCredito")

strBiblio=request.form("chkBiblio")
    if strCredito="" then
	   strCredito="N"
    else
	   strCredito="S"
    end if

	if strBiblio="" then
	   strBiblio="N"
    else
	   strBiblio="S"
    end if
	
		
set rs=Server.CreateObject ("ADODB.Recordset")
if AccionVar="Grabar" then
   if strcarrera="TODAS" then
	   	sql="Update MT_carrer set v_credito='" & strCredito & "',"
   		sql=sql & "v_bloqueo='" & strBiblio & "'"
   else
 	  	sql="Update MT_carrer set v_credito='" & strCredito & "',"
   		sql=sql & "v_bloqueo='" & strBiblio & "'"
   		sql=sql & " where codcarr='" & strcarrera & "'"	
   end if

   set rs=Session("Conn").Execute(sql)
   AccionVar=""
end if


'llenado de combos
'carga Carrera
'sql="select codcarr,nombre_C, nombre_l from MT_carrer order by sede, codcarr"
sql ="select codcarr,nombre_C, nombre_l from MT_carrer "
sql =sql & " Union "
sql =sql & " select convert(varchar(10),'TODAS') as codcarr, "
sql =sql & " convert(varchar(50),'OPCION TODAS LAS CARRERAS') AS nombre_C, "
sql =sql & " convert(varchar(50),'OPCION TODAS LAS CARRERAS') AS nombre_L "
sql =sql & " from mt_carrer where codcarr='GLOBAL' "




set rs=Session("Conn").Execute (sql)
if not rs.EOF then
	cmbCarrera=""
	pricarr=true	
	while not rs.EOF
		if strcarrera ="" and pricarr then
			'sel=" selected"
			pricarr=false
			strcarrera=rs("codcarr") 
		else
			if strcarrera=rs("codcarr") then
				sel=" selected"
			else
				sel=""
			end if
		end if
		cmbCarrera= cmbCarrera & "<option value=""" & rs("codcarr") & """ " & sel & ">"
		cmbCarrera=	cmbCarrera & rs("CodCarr") & " " & rs("nombre_l") & "</option>" & vbcrlf
		rs.MoveNext 
	wend
end if


sql="Select v_Credito,v_bloqueo from MT_Carrer where CodCarr='"& strcarrera & "'"
set rs=Session("Conn").Execute(sql)
if not rs.eof then
   if rs("v_credito")="S" then
      strCredito=" checked"
   else
      strCredito=""
   end if	  
   if rs("v_bloqueo")="S" then
      strBiblio=" checked"
   else
      strBiblio=""
   end if	     
else
  strCredito=""
  strBiblio=""   
end if

%>

<html>
<head>
<title>Documento sin t&iacute;tulo</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/tablas.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="TheForm" method="post" action="verifica-carrera.asp">
  <input type="hidden" name="Accion" value="<%=AccionVar%>">
  <table width="16%" border="0" bordercolor="#FF9900" height="20">
    <tr> 
        <td><input name="Guardar" type="button" value="Grabar" style="color: #00aa4f;FONT-SIZE: 10px; height:20; font-weight: bold;width:60" onClick="javascript:Grabar();"></td>
    </tr>
    </table>
  <table width="70%" height="80" border="0" cellpadding="1" cellspacing="1" bordercolor="#FFFFFF">
    <tr > 
        <td width="35%" height="8" background="Imagenes/fdo-cabecera-cel22.jpg"> 
          <div align="left" class="text-cabecera-celda"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><font color="#FFFFFF">Carrera:</font></font></b></div>
      </td>
        <td height="8" bgcolor="ffe5c3"><b> 
          <select name="cmbCarrera" onClick="javascript:Refrescar(this.value);" onChange="javascript:Actualizar();">
            <option value="">SELECCIONE </option>
			<%=cmbCarrera%> 
          </select>
          </b></td>
    </tr>
      <tr > 
        <td width="35%" height="25" background="Imagenes/fdo-cabecera-cel22.jpg"> 
          <div align="left" class="text-cabecera-celda"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><font color="#FFFFFF">Verificar 
            Bloqueo Financiero :</font></font></b></div>
        </td>
        <td height="25" bgcolor="ffe5c3" width="200"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
          <input type="checkbox" name="chkCredito" <%=strCredito%> value="checkbox">
          </font></b></td>
      </tr>
      <tr bgcolor="ffc172"> 
        <td width="35%" height="27" align="center" background="Imagenes/fdo-cabecera-cel22.jpg"> 
          <div align="left" class="text-cabecera-celda"><b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><font color="#FFFFFF">Verificar 
            Bloqueo Biblioteca:</font></font></b></div>
        </td>
        <td height="27" align="left" bgcolor="ffe5c3" width="200"><b><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
          <input type="checkbox" name="chkBiblio" <%=strBiblio%> value="checkbox">
          </font></b></td>
      </tr>
  </table>
</form>
</body>
<script>
function Refrescar(codcarr) {
 
 parent.bottomFrame.location.href="sis-inscripcion.asp?C=" + codcarr + "" 
}

function Actualizar() {
  document.TheForm.submit();
}

function Grabar() {
     if (document.TheForm.cmbCarrera.value==""){
			   alert("Debe Selecionar Carrera");
			   document.TheForm.elements['cmbCarrera'].focus();
			   return;
			 }		
     document.TheForm.Accion.value="Grabar";
     document.TheForm.submit();
}
</script>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
