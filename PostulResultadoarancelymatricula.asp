<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="include/conexion.inc" -->

<%
strDocUarm =  "SELECT dbo.Fn_ValorParame('CUPONESMATRICULAUARM')Parame"
set rstDocUarm = Session("Conn").execute(strDocUarm)
CUPONESMATRICULAUARM = rstDocUarm("Parame")
%>

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
					    <h1>Comprobante de Matr&iacute;cula y Arancel.</h1>
                        <h1 style="font-size:20px"></h1>
            <br />
            <table border="1">
                  <tr align="center">
                    <th scope="col">Producto &nbsp;&nbsp;</th>
                    <th scope="col">Tipo de Pago</th>
                    <th scope="col">Cup&oacute;n</th>
                    <th scope="col">Total</th>
                  </tr>
                 
                  <tr>
                    <!--<td>Matr&iacute;cula</td>-->
                    <td>Arancel B&aacute;sico</td>
                    <td><%=session("Matricula")%></td>
                    <td> 
					<%if session("MatriculaBanco")<>"" then%>
                    	<a  target="_blank"href="<%=session("MatriculaBanco")%>">Descargar</a>
                    <%elseif session("MatriculaPagare")<>"" then%>
                    	<a  target="_blank"href="<%=session("MatriculaPagare")%>">Descargar</a>
                    <%else%>
                    	N/A
					<%end if %>
                    </td>
                    <td>
                    <%if session("Matricula")="SERVIPAG" then%>
                    	Monto
                    <%else%>
                    	N/A
					<%end if %>
                    </td>
                  </tr>
                  <%if CUPONESMATRICULAUARM <> "SI" then%>
                  <tr>    
                 	<!--<td>Arancel</td>-->
                    <td>Arancel de Matr&iacute;cula</td>
                    <td><%=session("Arancel")%></td>
                    <td> 
                    <%if session("ArancelBanco")<>"" then%> 
                    	<a  target="_blank"href="<%=session("ArancelBanco")%>">Descargar</a>
                    <%elseif session("ArancelPagare")<>"" then%>
                    	<a  target="_blank"href="<%=session("ArancelPagare")%>">Descargar</a>
                    <%else%>
                    	N/A
					<%end if %>
                    </td>
                    <td>
                     <%if session("Arancel")="SERVIPAG" then%>
                    	Monto
                    <%else%>
                    	N/A
					<%end if %>
                    </td>
                  </tr>   
                  <%end if%>              
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