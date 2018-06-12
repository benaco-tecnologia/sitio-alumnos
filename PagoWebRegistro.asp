<%  response.buffer = false 
    Response.Expires = -1 
%>
<!-- saved from url=(0022)http://internet.e-mail -->
<!--<OBJECT RUNAT=server PROGID=ADODB.Recordset id=malla name=malla></OBJECT> -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<SCRIPT LANGUAGE="JavaScript" >

var total=0;
function inicio() {
}

function suma(cuanto) {
  total=total+cuanto;
}

function resta(cuanto) {
  total=total-cuanto;
}

function totalcero() {
  total=0;
}

function determina(cual, cuanto) {
         if(document.getElementById(cual).checked == true) {
                   suma(cuanto);
         } else {
                    resta(cuanto)
          }
        document.getElementById('pago').value = total; 
		document.getElementById('cuota1').value = total; 

}

function sumaresta(cual, cuanto, regtope) {
         if(document.getElementById(cual).checked == true) {
                   suma(cuanto);
         } else {
                    resta(cuanto);
					if (total < 0) {
					for(var ipos=0; ipos<regtope;ipos++)
                       { document.getElementById('id_apagar'+ipos).checked = false  
					   }
					totalcero();
					};
          }
        document.getElementById('pago').value = total; 
		document.getElementById('cuota1').value = total; 
}

function vatransbank( ) {
         if(document.getElementById('cuota1').value > 0) {
                   formBD.submit();
         } else {
                    alert('Debe seleccionar cuota a pagar');
          }
		
}
</script>
<!--<OBJECT RUNAT=server PROGID=ADODB.Recordset id=Detalle name=Detalle > </OBJECT>-->
<% 
RESPUESTA = Request.Form("TBK_RESPUESTA")
response.write "/LARESPUESTAES:" 
response.write RESPUESTA
response.write "/" 
ID_SESION = Request.Form("TBK_ID_SESION")
ORDEN_COMPRA = Request.Form("TBK_ORDEN_COMPRA")

if (RESPUESTA="0") then
 ' ******  COMMIT  ID_SESION
response.write("ACEPTADO")
GrabaPagosCommit()
elseif (RESPUESTA>"0") then

 ' ******  ROLLBACK  ID_SESION
 GrabaPagosRollback()
response.write("RECHAZADO")  
else 
response.write("CASO_ELSE")  
' GENERA ORDEN DE COMPRA A PARTIR DE FECHA 


FECHAACTUAL = NOW()
ANO = YEAR(FECHAACTUAL)
MES = MONTH(FECHAACTUAL)
DIA = DAY(FECHAACTUAL)
MINUTO = MINUTE(FECHAACTUAL)
SEGUNDO = SECOND(FECHAACTUAL)
OC = "OC_"&ANO&MES&DIA&MINUTO&SEGUNDO

codcli=session("codcli")
ID_TRANSACCION=codcli&OC 
response.write "***"
response.write "ID_TRANSACCION:" & ID_TRANSACCION
response.write "***"

dim strsql,strsql2  
dim montototal    
dim montoacancelar 
dim montoacancelarcuota  
dim maxfilas
dim feccorte

feccorte = Date()  

maxfilas=0  
montototal = 0
montoacancelar = 0 

dim i
dim t 
' RECOGE LA CANTIDAD DE FILAS A REVISAR y MONTOTOTAL
maxfilas=Request.Form("FILAS")
montoacancelar=Request.Form("cuota1") 
	  
response.write "/"
response.write "maxfilas:" & maxfilas
response.write "/"
response.write "montoacancelar:" & montoacancelar
response.write "/" 

'    *************************   ACÁ CORRESPONDE EL BEGIN TRANSACTION  "ID_TRANSACCION"   -- identificador unico

GrabaPagos_Inicia_Transaccion(ID_TRANSACCION)


' REVISA FILA POR FILA A VER SI LE CORRESPONDE PAGO
for i=0 to maxfilas-1

    if Request.Form("filaene"&cstr(i))=1 then	   
       response.write Request.Form("montofila"&cstr(i))  
       response.write Request.Form("montointmora"&cstr(i))   
       response.write Request.Form("ctadoc"&cstr(i))  
       response.write Request.Form("ctadocnum"&cstr(i))  
       response.write Request.Form("codban"&cstr(i))  
       response.write Request.Form("codapod"&cstr(i))  
	   response.write Request.Form("codcarr"&cstr(i))
	   response.write Request.Form("codcli"&cstr(i))

	   Regctadoc = Request.Form("ctadoc"&cstr(i))  
	   Regctadocnum = Request.Form("ctadocnum"&cstr(i))   
	   Regcodcli = Request.Form("codcli"&cstr(i))
	   Regcodapod = Request.Form("codapod"&cstr(i))
	   Regcodban = Request.Form("codban"&cstr(i))
	   Regcodcarr = Request.Form("codcarr"&cstr(i))
	   Regmontofila = Request.Form("montofila"&cstr(i))
	   
	   response.write "/" 
	   '********************************  INSERTAR EN LA TABLA DE LA CUOTA i   ****************
	   'ACA SE DEBE SEGUIR METIENDO MANO
	   'GrabaPagosLineaDetalle(Reg_ctadoc, Reg_ctadocnum, Reg_codcli, Reg_codapod, Reg_codban, Reg_codcarr, Reg_montofila)
	   'GPLineaDetalle2(Reg_ctadoc, Reg_ctadocnum, Reg_codcli, Reg_codapod, Reg_codban, Reg_codcarr, Reg_montofila)
	   GrabaPagosLineaDetalle Regctadoc, Regctadocnum, Regcodcli, Regcodapod, Regcodban, Regcodcarr, Regmontofila
	   'GrabaPagosLineaDetalle()
	   'GrabaPagosLineaDetalle(1,2,3,4,5,6)
	   
	   
	end if
	
next
'response.write "-FIN LISTADO-"
 
 'response.write "FIN FROMJGG"


%>
<html>
<head>

</head>
<body >
<% '********************************   VA  A   TRANSBANK     HTTP://webpay.nn.cl/PagoTabla.asp    ***********PagoTablaFalso.asp METHOD="POST" ACTION="/WEBPAY/TBK_BP_PAGO.CGI"
 %>
<FORM name="formBD" ACTION="/WEBPAY/TBK_BP_PAGO.CGI" METHOD="POST">  
		  <INPUT TYPE="HIDDEN" NAME="FILAS" VALUE="1"> <!--  < % =totalfilas%>  -->
		  <INPUT TYPE="HIDDEN" NAME="TBK_ORDEN_COMPRA" VALUE="<%=ID_TRANSACCION%>">
          <INPUT TYPE="HIDDEN" NAME="TBK_ID_SESION" VALUE="<%=ID_TRANSACCION%>">
		  <INPUT TYPE="HIDDEN" NAME="TBK_URL_EXITO" SIZE=40  VALUE="HTTP://alumnos.vallecentral.cl/PagoWebRegistro.ASP?TBK_RESPUESTA=0" SIZE="50">  
          <INPUT TYPE="HIDDEN" NAME="TBK_URL_FRACASO" SIZE=40 VALUE="HTTP://alumnos.vallecentral.cl/PagoWebRegistro.ASP?TBK_RESPUESTA=99" SIZE=50>  
		  <INPUT TYPE="HIDDEN" NAME="pagar1" VALUE="PAGAR">
		  <INPUT TYPE="HIDDEN" NAME="cuota1" VALUE="<%=montoacancelar%>" id="cuota1">
          <INPUT TYPE="button" NAME="pagar" id="bpagar" value="CONFIRME PAGAR" onClick="Javascript:vatransbank()">
</FORM>
</body>
</html>
<% 
end if
%>
<!--#INCLUDE file="include/desconexion.inc" -->
