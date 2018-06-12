<!--#INCLUDE FILE="include/funciones.inc" -->
<html>
    <head>
        <title> Redirect Moodle</title>
    </head> 
    <body onLoad="document.formulario.submit()"> 
    <body>  
        <form action="http://virtual.ucevalpo.cl/local/ssostandar/index.php" name="formulario"> 
            <input type="hidden" name="key" value="<%=ObtienekeyMoodleUCV%>"/> 
        </form>
    </body> 
</html>