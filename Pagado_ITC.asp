
<%
 'Borrado Cesar Araya M.
	'Fecha 2009-10-19
	'<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
	Response.AddHeader "cache-control","private"
	Response.AddHeader " pragma","no-cache"
	
	' Página redirige a EPS c/envío de parámetros
	Dim TxId
	Dim UrlEPS
	Dim Conn
	Dim Recset
	Dim XmlRespuesta
	
	Dim Firma
	Dim Token
	
	Dim etapa
	Dim CodigoBanco
	Dim boton_Log
	
' Objeto para Generar Firma
	Dim oGeneraFirma
	Dim GeneraFirma
	Dim lbRetGeneraFirma
	Dim lsDataFirma
	Dim lsfirmaRetorno
	Dim lsRetorno
	Dim oFirma
	Dim lsErrorData
	Dim DataFirma
	Dim keyprivada	
	Dim XML_out
	Dim XML_in
	
	etapa = 888
	boton_Log=""
	
	botonEstandar = "botonESTANDAR"

	sXML = Request.Form("XML")
	XML_in = sXML
	 
	 If (XML_in = "") Then
		'LogTransac "Error al Recepcionar XML desde Banco", "XML Vacio",BancoEstandar
		SendError -1, "Error al Recepcionar XML desde Banco [XML Vacio]"
	Else
		Set oXML = Server.CreateObject("MSXML2.DOMDocument")
		oXML.validateOnParse = True 
		If Not(oXML.loadXML(XML_in)) Then
			'LogTransac "XML recepcionado desde banco no válido", oXML.parseError.reason,BancoEstandar
			SendError -2, "XML recepcionado desde banco no válido [" & oXML.parseError.reason & "]"
		Else
			IdTrxServipag = BuscaParametros("Servipag/IdTrxServipag")
			IdTxCliente = BuscaParametros("Servipag/IdTxCliente")
			'EstadoPago = BuscaParametros("Servipag/EstadoPago")
			'Mensaje = BuscaParametros("Servipag/Mensaje")
			
			Dim miConn
		    Set miConn = Server.CreateObject("ADODB.Connection")
		    miConn.Open "Driver={SQL Server};Server=leo;Database=ITC;Uid=matricula;Pwd=dtb01s"
    				
		    Dim StrSql 
    		
		    StrSql = "Update Mt_Pago_Web Set Estado = 'OK WEB', ID_SERVIPAG = '" & IdTrxServiPag & "' Where Id = " & IdTxCliente
		    miConn.Execute StrSql
    		
    		StrSql = "Update Mt_Pago_Web_corr Set Estado = 'OK WEB', ID_SERVIPAG = '" & IdTrxServiPag & "' Where Id = " & IdTxCliente
		    miConn.Execute StrSql
            
            StrSql = " Exec Pr_Pagoweb_Ejecuta_Pago '" & IdTxCliente & "' "
            miConn.Execute StrSql
                		    		
		    miConn.Close
		    Set miConn = Nothing
			
		end if	
	end if
	 
	 'response.Write(">>***xml 2***<<")
	 'response.Write(sXML)
	 'response.Write(">>***xml 2***<<")
	 'response.End()
	 
	LogTransac "Xml Entrada", sXML, botonEstandar
					
	HeaderXML = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "ISO-8859-1" & Chr(34) & " ?>"
						
'-----------------------------------------------------------------
'  Conformar XML E.P.S
'-----------------------------------------------------------------

	XML_out = HeaderXML
	XML_out = XML_out & "<Servipag>"
	XML_out = XML_out & "<CodigoRetorno>0</CodigoRetorno>"
	XML_out = XML_out & "<MensajeRetorno>OK</MensajeRetorno>"
	XML_out = XML_out & "</Servipag>"

	'LogTransac "Xml Salida", XML_out,botonEstandar				

	XML_out = Replace(XML_out,vbCrlf," ")
	
	Response.Write XML_out
	
'-----------------------------------------------------------------
'  Fin Conformar XML E.P.S
'-----------------------------------------------------------------

'Marcar transacción pagada en sistema:

    
    


	
	'LogTransac "Despues de Response", XML_out,botonEstandar
	
'--------------------------------------------------------------------------------
' Log de Transacciones.
'--------------------------------------------------------------------------------
	Sub LogTransac(psContexto,psParametro,botonEstandar)
	Dim dia
	Dim mes
	Dim ano
           
	Dim archivo
	Dim fs
	Dim fname   
        
		dia = Day(Date)
		mes = Month(Date)
		ano = Year(Date)
      
		If (dia < 10) Then
			dia = "0" & dia 
		End If
		If (mes < 10) Then
			mes = "0" & mes 
		End If
		
		Dim miConn
		Set miConn = Server.CreateObject("ADODB.Connection")
		miConn.Open "Driver={SQL Server};Server=leo;Database=ITC;Uid=matricula;Pwd=dtb01s"
				
		Dim StrSql 
		
		StrSql = "Insert Into log(miLog) Values ('xml2:" + psContexto + "') 	"
		miConn.Execute StrSql
		
		miConn.Close
		Set miConn = Nothing
		

		'archivo = "C:\Simulador_bpe\Log\BPE_Recepciona_H2H" & botonEstandar &"." & ano & mes & dia & ".log"
		'archivo = "BPE_Recepciona_H2H" & botonEstandar &"." & ano & mes & dia & ".log"
		
		'document.write(archivo)
		
		'Set fs = Server.CreateObject("Scripting.FileSystemObject")		
		'Set fname = fs.OpenTextFile(archivo, 8, True)
		'fname.Write(Now & VbTab & boton_Log & VbTab & Mid(psContexto & Space(40),1,40) & " : " & VbTab & psParametro & VbCrLf)
		'fname.Close
		'Set fname = Nothing
		'Set fs = Nothing		
	End Sub	
	
	
	Function BuscaParametros(StrStruct)
	Dim NodeList
	Dim Node

		Set NodeList = oXML.selectNodes(StrStruct)
		For Each Node In NodeList
			BuscaParametros = Node.Text
		Next
	End Function	

%>

