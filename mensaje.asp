<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top12.asp" -->
<!--#INCLUDE file="f-izq-bloqueado.asp" -->
<!--#INCLUDE file="SubMenuExternoPrincipal.asp" -->


<%
Session.Abandon()
%>
<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title>Portal de Alumnos en L&iacute;nea</title>
<body >
<table border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right"><% CargarTop()%><% CargarMenu()%></td>
			<td width="516" valign="top" >
			<% CargarTop1()%>
                <table width="100%" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                  <tr>
                    <td align="center" valign="middle"> 
                      <table width="300" height="200" border="0" cellpadding="0" cellspacing="0">
                        <tr> 
                          <td background="/Imagenes/fondo_banner.jpg"><br> <br> <table width="500" border="0" align="left" cellpadding="0" cellspacing="0">
                              <tr> 
                                <td width="400"><div align="justify"> 
                                    <p><font color="#003366" size="4"><strong><font size="4" face="Arial, Helvetica, sans-serif"><br>
                                    Estimado  Alumno,<br>
                                      <br>
                                      En este momento no podemos atenderle.
                                       <br>
                                      Por favor, int&eacute;ntelo mas tarde.
                                       <br>
                                      Muchas Gracias.</font></strong></font><font color="#003366" size="4" face="Arial, Helvetica, sans-serif"><br>
                                      </font><font color="#003366" size="4"></font></p>
                                  </div></td>
                              </tr>
                          </table></td>
                        </tr>
                      </table>
                    </td>
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
