<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenuexterno.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%

strParame="SELECT coalesce(dbo.Fn_ValorParame('LargoMinClave'),0)LargoMinClave,coalesce(dbo.Fn_ValorParame('CantidadPreviaClave'),0)CantidadPreviaClave"
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
	LargoMinClave=rstParame("LargoMinClave")
	CantidadPreviaClave=rstParame("CantidadPreviaClave")
else
	LargoMinClave="0" 
	CantidadPreviaClave="0"
end if

strID="select id_usuario,id_persona from ca_usuarios WHERE us_consuser='" & Session("usrNW") & "'" 
set rstID= Session("Conn").Execute(strID)
if not rstID.eof then
	session("cc_id_usuario") = rstID("id_usuario")
	session("cc_id_persona") = rstID("id_persona")
else
	session("cc_id_usuario") = "0" 
	session("cc_id_persona") = "0"
end if 

if LargoMinClave = "1" then
	mensajePass="La nueva password debe tener un minimo de "& LargoMinClave &" caracter."
else
	mensajePass="La nueva password debe tener un minimo de "& LargoMinClave &" caracteres."
end if

ClavesConcatenadas =""
contador = 0

strClaves = "pa_ca_ordenaclaves2 '"& session("cc_id_usuario") &"' "
set rstClaves= Session("Conn").Execute(strClaves)		

while not rstClaves.eof 
	if cint(contador) < cint(CantidadPreviaClave) then
		ClavesConcatenadas = ClavesConcatenadas & rstClaves("us_passwordnew") & "-"
	end if 

	contador = contador + 1
	rstClaves.movenext	
wend
%>
<html>
<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript"  SRC="include/validarut.js"></SCRIPT>
<script language="JavaScript">

function Validar(campo1, campo2, campo3){
	 
	 
if (campo1 != "" && campo2 != "" && campo3 != "")
{
	if(	campo2.length >= <%=LargoMinClave%> )
	{		
	  if (ValidaCampo(campo1) && ValidaCampo(campo2))
		  { 
		   if (ComparaClaves(campo1,campo2))
		   {
			   if((tiene_letras(campo1)== 1 && tiene_numeros(campo1)== 1 && '<%=RequerirMixClave%>' == "SI") || ('<%=RequerirMixClave%>' != "SI")) 
			   {
				   	//validacion claves utilizadas 
					var ClavesConcatenadas = "<%=ClavesConcatenadas%>"
					ClavesConcatenadas = ClavesConcatenadas
					var result = "" 
					var claveUtilizadaAntes = "NO" 
					var x = 0
					var contador = 0
					
					for (i=0;i<ClavesConcatenadas.length;i++) {  
						if (ClavesConcatenadas.charAt(i) =="-")
						{	 
							result = ClavesConcatenadas.substr(x,i-x)
							if(result == campo1)
							{
								claveUtilizadaAntes="SI"
							}
							x = i + 1
							contador = contador + 1 
						}						
					} 		
					if (claveUtilizadaAntes=='NO')		
					{
						window.document.claves.action = "CambioClaveSql.asp";
				  	   	window.document.claves.submit(); 
					}
					else
					{				
						if (contador > 1)
						{
							alert("No se puede realizar el cambio, la nueva clave ya ha sido utilizada en una de las \u00faltimas "+ contador +" modificaciones.");
						}
						else
						{
							alert('No se puede realizar el cambio, la nueva clave ya ha sido utilizada en la última modificaci\u00f3n.');
						}
					}
			   }
			   else
			   {
				   alert('La nueva password debe contener letras y n\u00fameros.');
			   }
		   }
		   else
			  return false;
		 }
		 else	 
			 return false;  		
	}
	else
	{ 
		alert("<%=mensajePass%>");  
		return false;
	}
} 
else {
   alert("Por favor ingrese todos los datos solicitados.");
   return false;
  }
}

function tiene_letras(texto){
   var letras="abcdefghyjklmnñopqrstuvwxyz";	
   texto = texto.toLowerCase();
   for(i=0; i<texto.length; i++){
      if (letras.indexOf(texto.charAt(i),0)!=-1){
         return 1;
      }
   }
   return 0;
}
function tiene_numeros(texto){
	
   var numeros="0123456789";
   for(i=0; i<texto.length; i++){
      if (numeros.indexOf(texto.charAt(i),0)!=-1){
         return 1;
      }
   }
   return 0;
}
</script>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title>Portal de Alumnos en L&iacute;nea</title>
<body >
<table border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right" ><% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td valign="top">
			<% CargarTop1()%><% SubMenu()%>
			<table width="800" border="0" cellspacing="0" cellpadding="15" height="550" bgcolor="#FFFFFF">
			  <tr> 
                <td valign="top" align="left"><img src="imagenes/titulos/T-cambio-password-2.gif"> <br><br>
                Su clave ha expirado. Deber&aacute; Cambiarla
                <br><br>
                    <form name="claves" >
                      <table width="335" height="125" border="0" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
                        <tr align="center" valign="middle" bgcolor="#FFFFFF">
                          <td width="208" height="30" class="text-normal-celdas"><div align="right">Ingrese su password actual :</div></td>
                          <td width="163" height="30" valign="middle" background="imagenes/fondo-password.gif"><div align="center">
                              <input type="password" name="Clave" size="12" maxlength="20" class="casillas-form">
                          </div></td>
                        </tr>
                        <tr align="center" valign="middle"> 
                          <td height="30" class="text-normal-celdas"><div align="right">Ingrese su nueva password :</div></td>
                          <td height="30" background="imagenes/fondo-password.gif"><div align="center">
                              <input type="password" name="NuevaClave" size="12" maxlength="20" class="casillas-form">
                          </div></td>
                        </tr>
                        <tr align="center" valign="middle" bgcolor="#FFFFFF">
                          <td height="30" class="text-normal-celdas"><div align="right">Confirme su nueva password :</div></td>
                          <td height="30" background="imagenes/fondo-password.gif"><input type="password" name="RepClave" size="12" maxlength="20" class="casillas-form"></td>
                        </tr>
                        <tr align="center" valign="middle" bgcolor="#FFFFFF">
                          <td height="30" class="text-normal-celdas"><div align="right"></div></td>
              <td><div align="center"><a href="#" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('_r11_c21','','Imagenes/botones/A-_r12_c2_f2alt.jpg',1)" >
                              <input type="hidden" name="CambioClave" value="<%=CambioClave%>">
                              <img src="Imagenes/botones/A-_r12_c2alt.jpg" name="_r11_c21" width="130" height="21" border="0" align="left" id="_r11_c21" onClick="javascript: Validar(window.document.claves.NuevaClave.value, window.document.claves.RepClave.value, window.document.claves.Clave.value)"></a></div></td>
                        </tr>
                      </table>
                      <div align="left"></div>
                      <div align="left"></div>
                </form></td>
			  </tr>
			</table>
			</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>
</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->