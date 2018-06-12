<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<html>
<head>
<%
dim accionvar,strcarrera,strpromocion,strtexto,rs
AccionVar = request.form("Accion")
strcarrera=request.form("cmbcarrera")
strpromocion=request.form("cmbpromocion")
strtexto=request.form("txttexto")
set rs=Server.CreateObject ("ADODB.Recordset")

	if AccionVar ="Grabar" then
	   if strcarrera="" and strpromocion="" then
	        sql="update mt_parame set "
		    sql= sql & "texto_promocion ='" & strtexto & "'"		  
	   elseif strcarrera <> "" and strpromocion="" then 
		    sql="update mt_carrer set"
			sql=sql & " texto_promocion= '"& strtexto & "'"
			sql=sql & " where codcarr='" & strcarrera & "'"
	   else
			sql="select * from promocion_carrera"
			sql=sql & " where codcarr='" & strcarrera & "'"
			sql=sql & " and promocion ='" & strpromocion & "'"
		  	set rs=Session("Conn").execute(sql)
			if not rs.eof then
				sql="Update promocion_carrera set"
			  	sql=sql & " texto_promocion='"& strtexto & "'"
			 	sql=sql & " where codcarr='" & strcarrera & "'"
			    sql=sql & " and promocion ='" & strpromocion & "'"	
			 else
			 	sql="Insert Into promocion_carrera (codcarr,promocion,texto_promocion) values ("
			 	sql=sql & "'" & trim(strcarrera) & "',"
				sql=sql & "'" & trim(strpromocion) & "',"
				sql=sql & "'" & trim(strtexto) & "')"
			end if
	   end if
	    set rs=Session("Conn").execute(sql)
	end if
if strcarrera="" and strpromocion="" then
	      	sql="select texto_promocion from mt_parame  "
elseif strcarrera <> "" and strpromocion="" then 
		  	sql="select texto_promocion from mt_carrer where codcarr='" & strcarrera & "'"
else
			sql="select texto_promocion from promocion_carrera"
			sql=sql & " where codcarr='" & trim(strcarrera)& "' and  promocion= '" & trim(strpromocion)& "'"
end if
set rs=Session("Conn").execute(sql)
'response.Write(sql)
if not rs.eof then
    'response.write("Paso if")
	strtexto=rs("texto_promocion")
else
    'response.write("Paso else")
	strtexto=""
end if  
sql="select codcarr,nombre_C from MT_carrer order by Sede, CodCarr "
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
		cmbCarrera=	cmbCarrera & rs("CodCarr") & " " & rs("nombre_C") & "</option>" & vbcrlf
		rs.MoveNext 
	wend
end if

%>
<title>Documento sin t&iacute;tulo</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>
<form name="TheForm" method="post" action="">
<input type="hidden" name="Accion" value="<%=AccionVar%>">
  <p><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="3">Texto 
    Por Carrera / Promoci&oacute;n</font></b></font></p>
  <p><font size="4"></font> 
    <input name="Submit" style="color: #FFFFFF;FONT-SIZE: 10px; height:20; font-weight: bold;width:60" onClick="javascript:grabar();" type="submit" value="Grabar">
  </p>
  <table width="50%" border="0">
    <tr> 
      <td width="19%" bgcolor="4a5da1" height="20"> 
        <div align="left"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Promoci&oacute;n 
          : </font></div>
      </td>
      <td width="81%" bordercolor="#FFFFFF" bgcolor="ffe5c3" height="20"> 
	    <select name="cmbpromocion" id="cmbpromocion" style="width:100">
          <option value="">TODOS</option>
          <%for i= 1980 to 2020
		    if strpromocion=cstr(i) then
               sel=" Selected"
			else
			   sel=""   
			end if    %>	  
          <option value="<%=i%>" <%=sel%>><%=i%></option>
          <%next%>
        </select>
      </td>
    </tr>
    <tr> 
      <td bgcolor="4a5da1" width="19%"> 
        <div align="left"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Carrera 
          : </font></div>
      </td>
      <td bordercolor="#FFFFFF" bgcolor="ffe5c3" width="81%"> 
        <select name="cmbcarrera" onchange="javascript:refrescar();" id="cmbcarrera">
          <option value="">TODOS</option>
          <%=cmbCarrera%> 
        </select>
      </td>
    </tr>
    <tr> 
      <td height="11" bgcolor="4a5da1" width="19%"> 
        <div align="left"><font color="#FFFFFF" size="2" face="Arial, Helvetica, sans-serif">Texto 
          : </font></div>
      </td>
      <td height="63" bordercolor="#FFFFFF" bgcolor="ffe5c3" rowspan="2" width="81%"> 
        <textarea name="txttexto" id="txttexto" cols="50" rows="15"><%=strtexto%></textarea>
      </td>
    </tr>
    <tr> 
      <td bgcolor="4a5da1" width="19%">&nbsp;</td>
    </tr>
  </table>
  </form>
</body>
<script>
function grabar() {
	      		
		  	   if (document.TheForm.txttexto.value==""){
			   alert("Debe Ingresar Un Texto");
			   document.TheForm.elements['txttexto'].focus();
			   return;
			 }   
			    			 
	         document.TheForm.Accion.value ="Grabar"		
		  	 document.TheForm.submit();							 
			
	}
function refrescar() {
    document.TheForm.Accion.value =""		 
	document.TheForm.submit();
}

var splitIndex = 0;
var splitArray = new Array();

function splits(string,text) {
    var strLength = string.length, txtLength = text.length;
    if ((strLength == 0) || (txtLength == 0)) return;

    var i = string.indexOf(text);
    if ((!i) && (text != string.substring(0,txtLength))) return;
    if (i == -1) {
        splitArray[splitIndex++] = string;
        return;
    }

    splitArray[splitIndex++] = string.substring(0,i);
     
    if (i+txtLength < strLength)
        splits(string.substring(i+txtLength,strLength),text);

    return;
}

function split(string,text) {
    splitIndex = 0;
    splits(string,text);
}

function validate() {
	var keycode = event.keyCode
	if (keycode !=37 && keycode != 38 && keycode!= 39 && keycode != 40 && keycode !=16 && keycode !=17 && keycode!=18 && keycode !=46 && keycode !=8 && keycode !=35 && keycode != 36) {
		split(document.TheForm.txttexto.value,'\n');
		if (splitIndex > 15) {
			alert('Se permiten solo 15 filas')
			return false;
		}
		for (var i=0;i<splitIndex;i++) {
			if (splitArray[i].length > 220) {
				alert('Se permiten solo 250 caracteres como máximo.');
				return false;
			}
		}
	}
    return true;
}
document.TheForm.txttexto.onkeydown=validate
	
	
</script>
</html>

<!--#INCLUDE file="include/desconexion.inc" -->
