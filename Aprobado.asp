<% 
'-------------------------------------------------------------------------------------------------------------
' Pagina			: BPE_Comprobante.asp
' Propósito			: Pagina que emite el comprobante de pago de una E.P.S.
' Fecha Creación	: 12/03/2008
' Autor				: Luis Reyes Z.
' Empresa			: Kibernum S.A.
'-------------------------------------------------------------------------------------------------------------

Dim XML_out

	XML_out = Request.form("xml")
	  
	Dim oHTTP
	Dim oRexXML
	Dim TimeOut
	Dim lsRecXML
	Dim lsTXRec
	Dim lsDataRec
	Dim lsURLRec
	Dim lsDataConfirmaenPortal
	Dim lsURLConfirmaenPortal
	Dim CodRetREPago
	Dim Status
	Dim lsServerLocal
	Dim direc	

	Dim oXML
	Dim Conn
	Dim RecSet
	Dim SQL

	
	
	LogTransac "Recepcion XML",XML_out,BancoEstandar	
	
	'response.write("*** xml 4 ***")
	'reponse.write(
	'response.Write(replace(replace(XML_out, "<", "["),">","]")
	'response.write("*** xml 4 ***")
	
	'response.End()
	
	If (XML_out = "") Then
		'LogTransac "Error al Recepcionar XML desde Banco", "XML Vacio",BancoEstandar
		SendError -1, "Error al Recepcionar XML desde Banco [XML Vacio]"
	Else
		Set oXML = Server.CreateObject("MSXML2.DOMDocument")
		oXML.validateOnParse = True 
		If Not(oXML.loadXML(XML_out)) Then
			'LogTransac "XML recepcionado desde banco no válido", oXML.parseError.reason,BancoEstandar
			SendError -2, "XML recepcionado desde banco no válido [" & oXML.parseError.reason & "]"
		Else
			IdTrxServipag = BuscaParametros("Servipag/IdTrxServipag")
			IdTxCliente = BuscaParametros("Servipag/IdTxCliente")
			EstadoPago = BuscaParametros("Servipag/EstadoPago")
			Mensaje = BuscaParametros("Servipag/Mensaje")
			
		end if	
	end if

'----------------------------------------------------------------------------------------------------
' Log de Transacciones.
'----------------------------------------------------------------------------------------------------
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

		archivo = "xx_BPE_Comprobante_" & botonEstandar &"." & ano & mes & dia & ".log"
		response.Write(archivo)
		response.Write(Server.MapPath(archivo))
		
		Dim miConn
		Set miConn = Server.CreateObject("ADODB.Connection")
		miConn.Open "Driver={SQL Server};Server=leo;Database=ipll;Uid=matricula;Pwd=dtb01s"
		
		'oConn.Open "Driver={SQL Server};" & _ 
        '   "Server=MyServerName;" & _
        '   "Database=myDatabaseName;" & _
        '   "Uid=;" & _
        '   "Pwd="
		
		Dim StrSql 
		
		StrSql = "Insert Into log(miLog) Values ('xml4:" + psContexto + "') 	"
		miConn.Execute StrSql
		
		miConn.Close
		Set miConn = Nothing
		
		
		
		'Set fs = Server.CreateObject("Scripting.FileSystemObject")		
		'Set fname = fs.OpenTextFile(Server.MapPath(archivo), 8, True)
		'fname.Write(Now & VbTab & boton_Log & VbTab & Mid(psContexto & Space(40),1,40) & " : " & VbTab & psParametro & VbCrLf)
		'fname.Close
		'Set fname = Nothing
		'Set fs = Nothing		
	End Sub	

'--------------------------------------------------------------------------------
' Busca Parametros Con Estructura Fija Dentro De Un XML.
'--------------------------------------------------------------------------------
	Function BuscaParametros(StrStruct)
	Dim NodeList
	Dim Node

		Set NodeList = oXML.selectNodes(StrStruct)
		For Each Node In NodeList
			BuscaParametros = Node.Text
		Next
	End Function	

'--------------------------------------------------------------------------------
' Busca Servidor de instalacion
'--------------------------------------------------------------------------------
Dim Direc_Ini
Dim var_servidor


'Direc_Ini = "C:/simulador_bpe/simuladorBPE.ini"

var_servidor = "200.54.71.237" 'LeeIni(Direc_Ini,"servidor")

'------------------------------------------------------------------------------------------------------------------
' Lee ini
'------------------------------------------------------------------------------------------------------------------
	Function LeeIni(archivo,linea)
		Const fsoLectura = 1
  		Dim objFSO
		Dim objTextStream
		Dim Key


		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")

		If objFSO.FileExists(archivo) Then
		   Set objTextStream = objFSO.OpenTextFile(archivo, fsoLectura,true)
			Do While not objTextStream.AtEndOfStream
     			Key = Split(objTextStream.ReadLine, "=")
				If Key(0) = linea Then
				 	LeeIni = Key(1)
					Exit Do
				End If
   			loop
   				objTextStream.Close
   			Set objTextStream = Nothing
		Else
	    	LogTransac "Archivo Servidor", "Archivo de Token de direccion, no existe"
			LeeIni = "No Existe Archivo"
		End if
	End Function	
%> 



<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Pago de cuotas</title>
<link href="estilos.css" rel="stylesheet" type="text/css" />
<script type="text/JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function cerrar() {
    window.opener = null;
    window.close();
    return false;
}


//-->
</script>
</head>

<body onload="<script language='javascript'>compt=setTimeout('self.close();',5000);</script>" link="#717274" vlink="#717274" alink="#717274">
<table width="772" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td background="img/top-right.gif"><img src="img/spacerhoriz.gif" width="15" height="10" /></td>
    <td background="img/top.gif"><img src="img/spacerhoriz.gif" width="16" height="10" /></td>
    <td background="img/top-left.gif"><img src="img/spacerhoriz.gif" width="15" height="15" /></td>
  </tr>
  <tr>
    <td width="5" background="img/left.gif">&nbsp;</td>
    <td bordercolor="0"><table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFCC00">
      <tr>
        <td height="24" colspan="4" background="img/top-amarillo.gif">&nbsp;</td>
        </tr>
      
      <tr>
        <td width="4" height="100" rowspan="7" bgcolor="#717274"><img src="img/spacerhoriz.gif" width="4" height="10" /></td>
          <td height="27" colspan="2" background="img/titulo.gif" bgcolor="#E1E6EA">
              <span class="titulo_blanco_bold">Pago de cuotas Servipag</span></td>
        <td width="4" rowspan="7" bgcolor="#717274"><img src="img/spacerhoriz.gif" width="4" height="10" /></td>
      </tr>
      <tr>
        <td colspan="2" bgcolor="#FFFFFF"><img src="img/spacerhoriz.gif" width="15" height="10" /></td>
      </tr>
      <tr>
        <td height="46" colspan="2" background="img/top-gris.gif" class="titulo_blanco_bold">Operaci&oacute;n Finalizada, Antencedentes del Pago: </td>
      </tr>
      <tr>
        <td height="20" colspan="2" bgcolor="#E1E6EA">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="2" align="left" bgcolor="#E1E6EA"><table width="500" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" id="datos">
          <tr>
            <td width="100%" height="20" bgcolor="#FACC21" class="volver"><div align="left">Identificador de Transacción : <%=IdTrxServipag%></div></td>
           </tr>
		   <tr>
		    <td width="100%" height="20" bgcolor="#FACC21" class="volver"><div align="left">Identificador de Cliente : <%=IdTxCliente%></div></td>
            </tr>
		   <tr>
			<td width="100%" height="20" bgcolor="#FACC21" class="volver"><div align="left">Estado del Pago : <%=EstadoPago%></div></td>
			</tr>
		   <tr>
			<td width="100%" height="20" bgcolor="#FACC21" class="volver"><div align="left">Mensaje de Transacción : <%=Mensaje%></div></td>
			</tr>
			<tr>
			<td width="100%" height="20" bgcolor="#FACC21" class="volver"><div align="left">XML de confirmación de Pago : <%=XML_out%></div></td>
			</tr>
          
        </table>
       </tr>
      <tr>
	      <td height="20" colspan="2" bgcolor="#E1E6EA" align="center"><a href="javascript: res=cerrar();"><img src="img/Volver2-a.gif" alt="Volver al Simulador" name="boton_volver" width="149" height="49" border="0" id="boton_volver" /></a></td> 
		 
      </tr>
      <tr>
        <td colspan="4" background="img/bottom-amarillo.gif"><img src="img/spacerhoriz.gif" width="15" height="10" /></td>
      </tr>
    </table></td>
    <td width="5" background="img/right.gif">&nbsp;</td>
  </tr>
  <tr>
    <td bordercolor="0" background="img/bottom-left.gif"><img src="img/spacerhoriz.gif" width="15" height="15" /></td>
    <td background="img/bottom.gif"><img src="img/spacerhoriz.gif" width="16" height="10" /></td>
    <td background="img/bottom-right.gif"><img src="img/spacerhoriz.gif" width="15" height="10" /></td>
  </tr>
</table>
</body>
</html>




  	