<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%

if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if 

Rut = Session("Rutcli")
Matricula=request.form("rbMatricula")
Arancel=request.form("rbArancel")
ano =AnoAcad()
periodo=PeriodoAcad()
AnoAd=AnoAdmision()

MatriculaBanco=""
ArancelBanco=""
ArancelPagare=""
IDPAGOWEB=""
 
session("MatriculaBanco")=""
session("ArancelBanco")=""
session("ArancelPagare")=""
session("Matricula")=Matricula
session("Arancel")=Arancel

if Matricula ="SERVIPAG" and Arancel="SERVIPAG" then 
	
	sqlCreaPago = "Exec Pr_PagoWeb_Crea_Intencion @Rut = " + Rut
	Set RstCreaPago = Session("Conn").Execute(sqlCreaPago)
	
	SqlID = "Select Convert(Varchar, Coalesce(@@Identity, 0)) as IDPAGOWEB "
	Set RstID = Session("Conn").Execute(SqlID)	 
	IDPAGOWEB = RstID("IDPAGOWEB")
	
	Sql = "INSERT INTO mt_pago_web (id, rut)  SELECT id, rut FROM mt_pago_web_corr WHERE id = " & IDPAGOWEB
	Session("Conn").Execute Sql
	
	strM="INSERT INTO ra_Documentos_web( codcli ,documento ,ANO ,PERIODO ,fecha ,estado,tipo,idpagoweb)VALUES('"& session("codcli") &"','SERVIPAGMATRICULA',"& ano &","& periodo &",GETDATE(),'EN PROCESO','ALUMNO'," & IDPAGOWEB & ") "
	Session("Conn").execute(strM)
	
	strA="INSERT INTO ra_Documentos_web( codcli ,documento ,ANO ,PERIODO ,fecha ,estado,tipo,idpagoweb)VALUES('"& session("codcli") &"','SERVIPAGARANCEL',"& ano &","& periodo &",GETDATE(),'EN PROCESO','ALUMNO'," & IDPAGOWEB & ") "
	Session("Conn").execute(strA)  

	SqlMonto = "select TOP 1 COALESCE(sum(monto),0) as MONTOAPAGAR from MT_LISTAITEM WHERE CODCARR='" &  session("codcarr") & "' and CODITEM IN (1,2) AND GETDATE() BETWEEN FECINIVIG AND FECTERVIG and ano=" & AnoAd &""	
	Set RstMonto = Session("Conn").Execute(SqlMonto) 
	
	MONTOAPAGAR = RstMonto("MONTOAPAGAR") 
	
	'validacion cuando MONTOAPAGAR sea 0 
	
	Session("IDPAGOWEB")= IDPAGOWEB
	Session("monto_a_pagar") =MONTOAPAGAR
	Session("miXML") ="1"
	
%><script language="JavaScript">
	var props = "height=425,width=675,top=100,left=100,resizable=no,menubar=no,location=no,toolbar=no,status=no,scrollbars=no,directories=no";        
	window.open('BP.asp', 'Confirmación', props);
</script><%

else
	if Matricula="BANCO" then
		
		str="INSERT INTO ra_Documentos_web( codcli ,documento ,ANO ,PERIODO ,fecha ,estado,tipo)VALUES('"& session("codcli") &"','CUPONMATRICULA',"& ano &","& periodo &",GETDATE(),'EN PROCESO','ALUMNO') "
		Session("Conn").execute(str)
		
		strident= "SELECT @@IDENTITY codigo " 
		set rst = Session("Conn").execute(strident)
		MatriculaBanco = rst("codigo")
		
		strexec="sp_genera_documentos_web "& MatriculaBanco &""
		Session("Conn").execute(strexec)
		
		session("MatriculaBanco")="Documentos/" & session("codcli") & MatriculaBanco & ".pdf"
		
	elseif Matricula ="SERVIPAG" then
	
		sqlCreaPago = "Exec Pr_PagoWeb_Crea_Intencion @Rut = " + Rut
		Set RstCreaPago = Session("Conn").Execute(sqlCreaPago)
		
		SqlID = "Select Convert(Varchar, Coalesce(@@Identity, 0)) as IDPAGOWEB "
		Set RstID = Session("Conn").Execute(SqlID)	 
		IDPAGOWEB = RstID("IDPAGOWEB")
		
		Sql = "INSERT INTO mt_pago_web (id, rut)  SELECT id, rut FROM mt_pago_web_corr WHERE id = " & IDPAGOWEB
		Session("Conn").Execute Sql
		
		str="INSERT INTO ra_Documentos_web( codcli ,documento ,ANO ,PERIODO ,fecha ,estado,tipo,idpagoweb)VALUES('"& session("codcli") &"','SERVIPAGMATRICULA',"& ano &","& periodo &",GETDATE(),'EN PROCESO','ALUMNO'," & IDPAGOWEB & ") "
		Session("Conn").execute(str)  
	
		SqlMonto = "select TOP 1 COALESCE(sum(monto),0) as MONTOAPAGAR from MT_LISTAITEM WHERE CODCARR='" &  session("codcarr") & "' and CODITEM =1 AND GETDATE() BETWEEN FECINIVIG AND FECTERVIG and ano=" & AnoAd &""
		Set RstMonto = Session("Conn").Execute(SqlMonto) 
		MONTOAPAGAR = RstMonto("MONTOAPAGAR") 
		
		'validacion cuando MONTOAPAGAR sea 0 
		
		Session("IDPAGOWEB")= IDPAGOWEB
		Session("monto_a_pagar") =MONTOAPAGAR
		Session("miXML") ="1"
		
	%><script language="JavaScript">
		var props = "height=425,width=675,top=100,left=100,resizable=no,menubar=no,location=no,toolbar=no,status=no,scrollbars=no,directories=no";        
		window.open('BP.asp', 'Confirmación', props);
	</script><%
	
	end if 
	
	if Arancel="BANCO" then
	
		str="INSERT INTO ra_Documentos_web( codcli ,documento ,ANO ,PERIODO ,fecha ,estado,tipo)VALUES('"& session("codcli") &"','CUPONARANCEL',"& ano &","& periodo &",GETDATE(),'EN PROCESO','ALUMNO') "
		Session("Conn").execute(str)
		
		strident= "SELECT @@IDENTITY codigo " 
		set rst = Session("Conn").execute(strident)
		ArancelBanco = rst("codigo")
		
		strexec="sp_genera_documentos_web "& ArancelBanco &""
		Session("Conn").execute(strexec)
		
		session("ArancelBanco")="Documentos/" & session("codcli") & ArancelBanco & ".pdf"
		
	elseif Arancel ="SERVIPAG" then 
	
		sqlCreaPago = "Exec Pr_PagoWeb_Crea_Intencion @Rut = " + Rut
		Set RstCreaPago = Session("Conn").Execute(sqlCreaPago)
		
		SqlID = "Select Convert(Varchar, Coalesce(@@Identity, 0)) as IDPAGOWEB "
		Set RstID = Session("Conn").Execute(SqlID)	 
		IDPAGOWEB = RstID("IDPAGOWEB")
		
		Sql = "INSERT INTO mt_pago_web (id, rut)  SELECT id, rut FROM mt_pago_web_corr WHERE id = " & IDPAGOWEB
		Session("Conn").Execute Sql
		
		str="INSERT INTO ra_Documentos_web( codcli ,documento ,ANO ,PERIODO ,fecha ,estado,tipo,idpagoweb)VALUES('"& session("codcli") &"','SERVIPAGARANCEL',"& ano &","& periodo &",GETDATE(),'EN PROCESO','ALUMNO'," & IDPAGOWEB & ") "
		Session("Conn").execute(str)  
	
		SqlMonto = "select TOP 1 COALESCE(sum(monto),0) as MONTOAPAGAR from MT_LISTAITEM WHERE CODCARR='" &  session("codcarr") & "' and CODITEM =2 AND GETDATE() BETWEEN FECINIVIG AND FECTERVIG and ano=" & AnoAd &""
		Set RstMonto = Session("Conn").Execute(SqlMonto)
		MONTOAPAGAR = RstMonto("MONTOAPAGAR")
		 
		'validacion cuando MONTOAPAGAR sea 0 
		
		Session("IDPAGOWEB")= IDPAGOWEB
		Session("monto_a_pagar") =MONTOAPAGAR
		Session("miXML") ="1"
		
	%><script language="JavaScript">
		var props = "height=425,width=675,top=100,left=100,resizable=no,menubar=no,location=no,toolbar=no,status=no,scrollbars=no,directories=no";        
		window.open('BP.asp', 'Confirmación', props);
	</script><%
	
	elseif Arancel ="PAGARE" then
	
		str="INSERT INTO ra_Documentos_web( codcli ,documento ,ANO ,PERIODO ,fecha ,estado,tipo)VALUES('"& session("codcli") &"','PAGARE',"& ano &","& periodo &",GETDATE(),'EN PROCESO','ALUMNO') "
		Session("Conn").execute(str) 
		
		strident= "SELECT @@IDENTITY codigo " 
		set rst = Session("Conn").execute(strident)
		ArancelPagare = rst("codigo")
		 
		strexec="sp_genera_documentos_web "& ArancelPagare &""
		Session("Conn").execute(strexec)
		
		session("ArancelPagare")="Documentos/" & session("codcli") & ArancelPagare & ".pdf"
	
	end if  
end if



if Matricula <> "SERVIPAG" and Arancel <> "SERVIPAG" then 
	response.Redirect("Resultadoarancelymatricula.asp") 
end if 
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="css/css/parrafo.css" rel="stylesheet" type="text/css">
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<body >
<table border="0" cellpadding="0" cellspacing="0"  align="left">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right"><% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td valign="top" >
			<% CargarTop1()%><% SubMenu()%>
				  <table width="100%" border="0" cellspacing="0" cellpadding="15" height="550" bgcolor="#FFFFFF">
			  <tr> 
				<td valign="top">
				  <table width="800" border="0" cellpadding="0" cellspacing="0">
                  <tr valign="top" bgcolor="#FFFFFF">                   
                    </tr>
                    
                    <tr valign="top" bgcolor="#FFFFFF">
					  <td height="24" >
					    
					  </td> 
					  <td>&nbsp;</td>
				      <td>&nbsp;</td>
                      <td width="221">&nbsp;</td>
					</tr>
                    <tr valign="top" bgcolor="#FFFFFF">
					  <td height="24">
					    <p class="text-menu">&nbsp;</p>
					  </td>
					  <td>&nbsp;</td>
				      <td></td>
					</tr>
                    <tr valign="top" bgcolor="#FFFFFF">
					  <td height="24">
					    <p style="font-size:25px" class="text-menu">Validando Pago...</p>
					  </td>
					  <td>&nbsp;</td>
				      <td>&nbsp;</td>                      
					</tr>
                    <tr valign="top" bgcolor="#FFFFFF">
					  <td height="24">
					  </td>
					  <td></td>
				      <td></td>                      
					</tr>
                    <tr valign="top" bgcolor="#FFFFFF">
					  <td height="24">
					  </td>
					  <td></td>
				      <td></td>                      
					</tr>
                    <tr valign="top" bgcolor="#FFFFFF">
					  <td height="24">
                      </td>
					  <td></td>
                      <td></td> 
				      <td>&nbsp;</td>                      
					</tr>
			    </table>      
			      <table width="738">
                    <tr>
                    </tr>
                    <tr>
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
	</td>
  </tr>
</table>
</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
<!--javascript:parent.mainframe.history.go(0)-->