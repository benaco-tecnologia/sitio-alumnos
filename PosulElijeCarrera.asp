<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1"> 
<head profile="http://gmpg.org/xfn/11">
	<title>Sistema de Certificados en l&iacute;nea</title>	
	<link rel="shortcut icon" href="image/favicon.ico" />
	<link rel="stylesheet" href="style.css" type="text/css" media="screen" />
	<meta http-equiv="content-type" content="text/html; charset=utf-8" />
	<meta http-equiv="content-language" content="en-gb" />
	<meta http-equiv="imagetoolbar" content="false" />
	<meta name="author" content="Christopher Robinson" />
	<meta name="copyright" content="Copyright (c) Christopher Robinson 2005 - 2007" />
	<meta name="description" content="" />
	<meta name="keywords" content="" />	
	<meta name="last-modified" content="Sat, 01 Jan 2007 00:00:00 GMT" />
	<meta name="mssmarttagspreventparsing" content="true" />	
	<meta name="robots" content="index, follow, noarchive" />
	<meta name="revisit-after" content="7 days" />
</head>
<%
rut = request("rut")

if session("tipo")="ALUMNO" then
	strA = "select carr.CODCARR, c.nombre +' '+c.paterno NOMBRE,carr.codcarr + ' - ' +carr.nombre_l CARRERA from mt_alumno a,mt_client c,mt_Carrer carr  "
	strA = strA & " WHERE a.rut=c.codcli AND a.rut='" & rut & "' AND carr.CODCARR=a.codcarpr "
else
	strA = "select carr.CODCARR, c.nombre +' '+c.paterno NOMBRE,carr.codcarr + ' - ' +carr.nombre_l CARRERA from MT_POSCAR p,mt_client c,mt_Carrer carr  "
	strA = strA & " WHERE p.CODPOSTUL=c.codcli AND p.CODPOSTUL='" & rut & "' and p.ESTADO='A' AND carr.CODCARR=p.codcarr "
end if 

set RstA = Session("Conn").Execute (strA)

%>
<body>
	<div id="container">
		<div id="navigation">
			<ul>
				
			</ul>
		</div>
		<div id="content">       
        <table width="650" border="0" cellpadding="0" cellspacing="0" >
					<tr valign="top" >
					  <td width="425" height="100">
					    <h1>Seleccione una Carrera.</h1>
            <br />
            <table width="341" border="1"> 
            <form action="VerificaPostulantes.asp" method="post" name="TheForm" id="TheForm">                
                  <tr>
                    <!--<td>Matr&iacute;cula</td>-->
                    <td width="331">
                    
                    <select name="codcarr" style="width:331px;border: 1px black outset" >                    
						<%While Not RstA.Eof%>
    						<option value="<%=RstA("CODCARR")%>"><%=RstA("CARRERA")%></option>
                        <%RstA.movenext 
                        wend%>
                    </select>
                    <input type="hidden" name="rut" value="<%=rut%>">
                    </td>                    
                  </tr>   
                  
                 <tr>
                   <td width="331"></td>  
                 </tr> 
                 
                 <tr>
                   <td width="331"></td>  
                 </tr> 
                  
                  <tr>
                   <td width="331">                   
                   <input  style="background-color: #eeeeee; border: 1px black outset" type="submit" value=" Ingresar ">
                   </td>  
                  </tr>   
                  </form>
                </table>   
            </td>
					  <td width="66">&nbsp;</td>
				      <td width="159"><img src="imagenes/2_r1_c2.jpg" height="130" width="150"></a></td>
					</tr>
                    </table>			
            <br />
			<br />	
        	<input  style="background-color: #eeeeee; border: 1px black outset" type="button" value=" Volver " onclick="location.href = 'documentospostulantes.asp'">
            
		</div>
		<div id="footer"></div>
	</div>
</body>

</html>