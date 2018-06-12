<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
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
        <%
		rut=session("RutPostul")
		
		ano =AnoAcad()
		periodo=PeriodoAcad()
		
		str="INSERT INTO ra_Documentos_web( codcli ,documento ,ANO ,PERIODO ,fecha ,estado,tipo)VALUES('"& session("RutPostul") &"','CUPONARANCEL',"& ano &","& periodo &",GETDATE(),'EN PROCESO','POSTULANTE') "
		Session("Conn").execute(str)
		
		strident= "SELECT @@IDENTITY codigo " 
		set rst = Session("Conn").execute(strident)
		 
		strexec="sp_genera_documentos_web "& rst("codigo") &""
		Session("Conn").execute(strexec)
		
		%>
        
        <table width="650" border="0" cellpadding="0" cellspacing="0" >
            <tr valign="top" >
                <td width="425" height="100">
                    <h1>Cup&oacute;n de Arancel</h1>                    
                    <br /><a href="Documentos/<%=session("RutPostul")& rst("codigo")%>.pdf" style="font-size:18px">Descargar Cup&oacute;n Arancel<a></td>                
                <td width="159"><img src="imagenes/2_r1_c2.jpg" height="130" width="150"></td>
            </tr>           
        </table>			
        <br />
        <br />
        <input  style="background-color: #eeeeee; border: 1px black outset" type="button" value=" Volver " onclick="location.href = 'verificapostulantes.asp?rut=<%=rut%>'">  
		</div>
		<div id="footer"></div>
	</div>
</body>

</html>