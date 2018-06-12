 <htm>
 <script language="javascript">
 function EnviaDatos(){
 if (window.document.cursos.curso[0].checked || window.document.cursos.curso[1].checked)
     {
	  if (window.document.cursos.curso[0].checked){
	     window.document.cursos.CodCalc.value = "Ca1";
		 window.document.cursos.SecCalc.value = "1";
		 window.document.cursos.HorCalc.value = "Lu - Ju";
		 window.document.cursos.action = "pasoCalc.asp";
		 window.document.cursos.submit();
	  }
	  if (window.document.cursos.curso[1].checked){
	     window.document.cursos.CodCalc.value = "Ca2";
		 window.document.cursos.SecCalc.value = "2";
		 window.document.cursos.HorCalc.value = "Ma - Mi";
		 window.document.cursos.action = "pasoCalc.asp";
		 window.document.cursos.submit();
	  }
	 }
 else{
     alert("Debe seleccionar un curso.");
	 return false;
 }	 
 }
 </script>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="cursos">
        
  <table width="600" cellspacing="1" cellpadding="0" height="100" border="0" bordercolor="#FFFFFF">
    <tr bgcolor="4a5da1"> 
            <td width="48" height="16" background="Imagenes/fdo-cabecera-cel22.jpg" > 
              <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">C&oacute;digo</font></b></font></div>
        </td>
            <td width="119" height="16" background="Imagenes/fdo-cabecera-cel22.jpg" bgcolor="#FFFFFF"> 
              <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Asignatura</font></b></font></div>
        </td>
            <td width="66" height="16" align="center" background="Imagenes/fdo-cabecera-cel22.jpg" bgcolor="#FFFFFF"> 
              <div align="center"><b><font size="1" face="Verdana" color="#FFFFFF">secci&oacute;n</font></b></div>
        </td>
            <td width="60" height="16" align="center" background="Imagenes/fdo-cabecera-cel22.jpg" bgcolor="#FFFFFF"> 
              <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Horario</font></b></font></div>
        </td>
            <td width="56" height="16" valign="top" background="Imagenes/fdo-cabecera-cel22.jpg" bgcolor="#FFFFFF"> 
              <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Inscribir</font></b></font></div>
        </td>
      </tr>
          <tr bgcolor="ffc172"> 
            
    <td width="48" height="16" bgcolor="#DBECF2"><font face="Verdana" size="1">&nbsp;Ca1</font></td>
            <td width="119" height="16" bgcolor="#DBECF2"><font face="Verdana" size="1">Calculo&nbsp;I</font></td>
            
    <td width="66" height="16" align="center" bgcolor="#DBECF2"><font face="Verdana" size="1">1</font></td>
            
    <td width="60" height="16" align="center" bgcolor="#DBECF2"><font face="Verdana" size="1">Lu 
      - Ju</font></td>
            <td width="56" height="16" valign="top" bgcolor="#DBECF2"> 
              <div align="center"> 
                <font face="Verdana" size="1"> 
                <input type="radio" name="curso" value="1">
                </font>
              </div>
            </td>
          </tr>
          <tr bgcolor="ffc172"> 
            <td width="48" height="16" bgcolor="#DBECF2"><font face="Verdana"><font size="1">&nbsp;Ca2</font></font></td>
            <td width="119" height="16" bgcolor="#DBECF2"><font face="Verdana" size="1">Calculo I</font></td>
            <td width="66" height="16" align="center" bgcolor="#DBECF2">
              <p align="center"><font face="Verdana" size="1">2</font></p>
            </td>
            <td width="60" height="16" align="center" bgcolor="#DBECF2"><font face="Verdana" size="1">Ma
              - Mi</font></td>
            <td width="56" height="16" valign="top" bgcolor="#DBECF2"> 
              <div align="center"> 
                <font face="Verdana" size="1"> 
                <input type="radio" name="curso" value="2">
                </font>
              </div>
            </td>
          </tr>
    </table>
  <p align="left"> 
    <input name="button" type="button" onClick="javascript:EnviaDatos();" value="Inscribir" named="inscribir">
    <input type="hidden" name="CodCalc" value="">
    <input type="hidden" name="SecCalc" value="">
    <input type="hidden" name="HorCalc" value="">
  </p>
  </form>		
</htm>