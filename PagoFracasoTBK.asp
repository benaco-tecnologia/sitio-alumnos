<!--#INCLUDE FILE="include/conexion.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <title>Pago de cuotas</title>
    <link href="estilos.css" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" href="css/tablas.css" type="text/css" />

    <script type="text/JavaScript">
        function cerrar() {
            window.open('', '_parent', '');
            window.close();
        }
    </script>

    <style type="text/css">
<!--
.style1 {font-size: 24px}
.style2 {
	font-size: 12px
}
-->
    </style>
</head>

<% 
	' Recupera número de Orden de Compra
	if Session("OC") <> "" then
		OC = Session("OC")
		Session("OC") = ""
	else
		OC = Request.Form("TBK_ORDEN_COMPRA")
	end if 
	
	str = "UPDATE MT_WEBPAY SET ESTADO ='FRACASO' WHERE TBK_ORDEN_COMPRA ='" & OC & "' AND ESTADO = 'CREADO'"
	Session("Conn").execute(str)
	
%>
<body onload="" link="#717274" vlink="#717274" alink="#717274">
    
    <center>
      <p><br />
      <img src="imagenes/warning.gif"  /><br /> 
        <br />
        <br />
        <b class="Tit-celdas style2"><font face="Verdana, Arial, Helvetica, sans-serif">Transacción Rechazada<br /> 
        OC Nº: <%=OC%> <br />  
        </font> </b>
       
        <TABLE  WIDTH=679>
        <TR>
        <TD WIDTH=100><b class="Tit-celdas style2"><font face="Verdana, Arial, Helvetica, sans-serif">Las posibles causas de este rechazo son:<br />					        </font> </b></TD> 
        </TR>
        <TR>
        <TD WIDTH=100></TD> 
        </TR>
        <TR>
        <TD WIDTH=100> <b class="Tit-celdas style2"><font face="Verdana, Arial, Helvetica, sans-serif" >
        •	Error en el ingreso de los datos de su tarjeta de crédito o débito (fecha y/o código de seguridad). <br />          
        </font> </b></TD> 
        </TR>
        <TR>
        <TD WIDTH=100><b class="Tit-celdas style2"><font face="Verdana, Arial, Helvetica, sans-serif" >
        •	Su tarjeta de crédito o débito no cuenta con el cupo necesario para cancelar la compra.  <br />
        </font> </b></TD>   
        </TR>
        <TR>
        <TD WIDTH=100> <b class="Tit-celdas style2"><font face="Verdana, Arial, Helvetica, sans-serif" >
        •	Tarjeta aún no habilitada en el sistema financiero. <br />          
        </font> </b></TD>   
        </TR>
        </TABLE>
        
        <br />
        <br />
        <br />
      <a href="javascript: res=cerrar();"><img src="Imagenes/botones/continuar-of.gif" alt="Volver al Portal de Alumnos" name="boton_volver" width="149"
                                    height="49" border="0" id="boton_volver" onclick="return cerrar()" /></a></p>
</center>
    
</body>
</html>
