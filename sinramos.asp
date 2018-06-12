<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->


<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="css/tablas.css" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000099">
<table border="0" cellpadding="0" cellspacing="0"  align="left" >
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td valign="top" >
				<% CargarTop1()%><% SubMenu()%>
				<table border="0" cellpadding="0" cellspacing="15">
					<p>&nbsp;</p>
					<td><font size="3" class="Tit-celdas"><b>Ud. No posee ramos propuestos por 
					el sistema ...<br>
					<br>
					<br>
					</b></font>
                    <%
					strParame="SELECT coalesce(dbo.Fn_ValorParame('USAMENSAJESINRAMOS_ESUCO'),'NO')Parame"
					set rstParame= Session("Conn").Execute(strParame)		
					if not rstParame.eof then
						USAMENSAJESINRAMOS_ESUCO=rstParame("Parame")
					end if
					
					if USAMENSAJESINRAMOS_ESUCO <> "SI" then%>
                        <b>
                        <font size="2" class="Tit-celdas">Deber&aacute; realizar 
                        su inscripci&oacute;n de toda su carga acad&eacute;mica<br>
                        a trav&eacute;s de una Solicitud Especial de Inscripci&oacute;n de <br>
                        Asignaturas</font>
                        </b>
                    <%else%>
                        <b>
                        <font size="2" class="Tit-celdas">De acuerdo a su avance acad&eacute;mico, Usted no tiene asignaturas para inscribir.<br>
Deber&aacute; realizar su Proceso de Inscripci&oacute;n de asignaturas de manera asistida con su Jefatura de Carrera.</font>
                        </b>
                    <%end if%>
                    </TD>
			</table>
			</td>
		 </tr>
	  </table>
	</td>
  </tr>
</table>	
</body>
</html>
