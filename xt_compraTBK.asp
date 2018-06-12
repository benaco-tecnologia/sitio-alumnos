<!--#INCLUDE FILE="include/conexion.inc" -->  
<HTML http-equiv="Content-Type" content="text/html;charset=ISO-8859-8">
<% 
RESPUESTA = Request.Form("TBK_RESPUESTA")
session("OC") = Request.Form("TBK_ORDEN_COMPRA")
session("TBK_ORDEN_COMPRA") = Request.Form("TBK_ORDEN_COMPRA") 
session("TBK_TIPO_TRANSACCION") = Request.Form("TBK_TIPO_TRANSACCION") 
session("TBK_RESPUESTA") = Request.Form("TBK_RESPUESTA") 
session("TBK_MONTO") = Request.Form("TBK_MONTO")      
session("TBK_CODIGO_AUTORIZACION") = Request.Form("TBK_CODIGO_AUTORIZACION")
session("TBK_FINAL_NUMERO_TARJETA") = Request.Form("TBK_FINAL_NUMERO_TARJETA")
session("TBK_FECHA_CONTABLE") = Request.Form("TBK_FECHA_CONTABLE")
session("TBK_FECHA_TRANSACCION") = Request.Form("TBK_FECHA_TRANSACCION")
session("TBK_HORA_TRANSACCION") = Request.Form("TBK_HORA_TRANSACCION")
session("TBK_ID_SESION") = Request.Form("TBK_ID_SESION")
session("TBK_ID_TRANSACCION") = Request.Form("TBK_ID_TRANSACCION")
session("TBK_TIPO_PAGO") = Request.Form("TBK_TIPO_PAGO")
session("TBK_NUMERO_CUOTAS") = Request.Form("TBK_NUMERO_CUOTAS")
session("TBK_TASA_INTERES_MAX") = Request.Form("TBK_TASA_INTERES_MAX")
session("TBK_VCI") = Request.Form("TBK_VCI") 
session("TBK_MAC") = Request.Form("TBK_MAC") 

if (RESPUESTA="0") then
	if (check_mac = "CORRECTO") then	' Sòlo si MAC ok, oc ok, monto ok se debe responder "ACEPTADO"
		if (valida_oc = "CORRECTO")then ' Aqui debe validar estado de OC
			if (valida_monto = "CORRECTO")then 'Aqui debe validar Monto
				response.write("ACEPTADO")
			else
				response.write("RECHAZADO")
			end if  		
		else 
			response.write("RECHAZADO")
		end if
	else
		response.write("RECHAZADO")
	end if
else
' Acepta no autorización de la transacción 
response.write("ACEPTADO")
end if

function valida_monto() 

	sqlValidaM="SP_VALIDA_MONTO_TBK '"& session("TBK_ORDEN_COMPRA") &"','"& session("TBK_MONTO") &"'" 
	if BCL_ADO(sqlValidaM, RstVM) then
	
		if RstVM("RESPUESTA") = "CORRECTO" then	
			valida_monto="CORRECTO"
		else
		    valida_monto="INCORRECTO"
		end if 
		
    end if
	
end function

function valida_oc() 

	sqlValida="SP_VALIDA_OC_TBK '"& session("TBK_ORDEN_COMPRA") &"'"
	if BCL_ADO(sqlValida, RstV) then
	
		if RstV("RESPUESTA") = "CORRECTO" then
	        strA ="SP_ACTUALIZA_WEBPAY '"& session("TBK_ORDEN_COMPRA") &"','"& Session("CodCli") &"','"& session("TBK_TIPO_TRANSACCION") &"','"& session("TBK_RESPUESTA") &"','"& session("TBK_MONTO") &"','"& session("TBK_CODIGO_AUTORIZACION") &"','"& session("TBK_FINAL_NUMERO_TARJETA") &"','"& session("TBK_FECHA_CONTABLE") &"','"& session("TBK_FECHA_TRANSACCION") &"','"& session("TBK_HORA_TRANSACCION") &"','"& session("TBK_ID_SESION") &"','"& session("TBK_ID_TRANSACCION") &"','"& session("TBK_TIPO_PAGO") &"','"& session("TBK_NUMERO_CUOTAS") &"','"& session("TBK_VCI")&"'" 
			Session("Conn").execute(strA)
	
			valida_oc = "CORRECTO" 
		else
		    valida_oc = "INCORRECTO"
		end if 
		
    end if
	
end function

function check_mac ()

strParame="SELECT dbo.Fn_ValorParame('tbk_carpetaLogs')tbk_carpetaLogs,dbo.Fn_ValorParame('tbk_archivoBat')tbk_archivoBat,dbo.Fn_ValorParame('tbk_ejecutable_CheckMac')tbk_ejecutable_CheckMac"

if BCL_ADO(strParame, rstParame) then
	tbk_carpetaLogs = trim(valnulo(rstParame("tbk_carpetaLogs"),STR_))
	tbk_archivoBat = trim(valnulo(rstParame("tbk_archivoBat"),STR_))
	tbk_ejecutable_CheckMac = trim(valnulo(rstParame("tbk_ejecutable_CheckMac"),STR_))
end if

carpetaLogs = tbk_carpetaLogs
archivoBat = tbk_archivoBat
ejecutable_CheckMac = tbk_ejecutable_CheckMac

archivo_temporal = carpetaLogs & "DatosParaCheckMac_" & session("OC") & ".txt"
archivo_resultado = carpetaLogs & "ResultadoCheckMac_" & session("OC") & ".txt"

set filesys = CreateObject("Scripting.FileSystemObject")
set file = filesys.CreateTextFile(archivo_temporal)
' Recupera parametros y guarda en archivo
file.write("TBK_ORDEN_COMPRA=" & session("TBK_ORDEN_COMPRA") & "&")
file.write("TBK_TIPO_TRANSACCION=" & session("TBK_TIPO_TRANSACCION") & "&")
file.write("TBK_RESPUESTA=" & session("TBK_RESPUESTA") & "&")
file.write("TBK_MONTO=" & session("TBK_MONTO") & "&")
file.write("TBK_CODIGO_AUTORIZACION=" & session("TBK_CODIGO_AUTORIZACION") & "&")
file.write("TBK_FINAL_NUMERO_TARJETA=" & session("TBK_FINAL_NUMERO_TARJETA") & "&")
file.write("TBK_FECHA_CONTABLE=" & session("TBK_FECHA_CONTABLE") & "&")
file.write("TBK_FECHA_TRANSACCION=" & session("TBK_FECHA_TRANSACCION") & "&")
file.write("TBK_HORA_TRANSACCION=" & session("TBK_HORA_TRANSACCION") & "&")
file.write("TBK_ID_SESION=" & session("TBK_ID_SESION") & "&")
file.write("TBK_ID_TRANSACCION=" & session("TBK_ID_TRANSACCION") & "&")
file.write("TBK_TIPO_PAGO=" & session("TBK_TIPO_PAGO") & "&") 
file.write("TBK_NUMERO_CUOTAS=" & session("TBK_NUMERO_CUOTAS") & "&")
'file.write("TBK_TASA_INTERES_MAX=" & TBK_TASA_INTERES_MAX & "&")
file.write("TBK_VCI=" & session("TBK_VCI") & "&") 
file.write("TBK_MAC=" & session("TBK_MAC"))  
file.Close

cmd = archivoBat & " " & ejecutable_CheckMac & " " & archivo_temporal & " " & archivo_resultado
Set WshShell = CreateObject ("WScript.Shell")
iReturn = WshShell.Run(cmd,1,true)
set WshShell = nothing
' Lee resultado de validación de MAC
set ArchResultado = CreateObject("Scripting.FileSystemObject")
set tf = ArchResultado.Opentextfile(archivo_resultado)
check_mac = tf.readLine

end function
%>
</HTML>