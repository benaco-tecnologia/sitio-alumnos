<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Strict//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-strict.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<meta http-equiv="Content-Type" content="text/html; charset=ISO-8859-1"> 
<head profile="http://gmpg.org/xfn/11">
	<title>Documentos Postulantes</title>	
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
			if Session("RutPostul") = "" then
			  Response.Redirect("documentospostulantes.asp")
			end if
			  
			Rut = session("RutPostul")
			Codcarr=session("codcarrPostul")
			
			Matricula=request.form("rbMatricula")
			Arancel=request.form("rbArancel")
			
			ano =AnoAdmision()			
			periodo=PeriodoAdmision()
			AnoAd=AnoAdmision()
			
			MatriculaBanco=""
			MatriculaPagare=""
			ArancelBanco=""
			ArancelPagare=""
			IDPAGOWEB=""
			 
			session("MatriculaBanco")=""
			session("MatriculaPagare")=""
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
				
				strM="INSERT INTO ra_Documentos_web( rut,codcarr ,documento ,ANO ,PERIODO ,fecha ,estado,tipo,idpagoweb)VALUES('"& Rut &"','"& Codcarr &"','SERVIPAGMATRICULA',"& ano &","& periodo &",GETDATE(),'EN PROCESO','" & session("tipo") & "'," & IDPAGOWEB & ") "
				Session("Conn").execute(strM)
				
				strA="INSERT INTO ra_Documentos_web( rut,codcarr ,documento ,ANO ,PERIODO ,fecha ,estado,tipo,idpagoweb)VALUES('"& Rut &"','"& Codcarr &"','SERVIPAGARANCEL',"& ano &","& periodo &",GETDATE(),'EN PROCESO','" & session("tipo") & "'," & IDPAGOWEB & ") "
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
					
					str="INSERT INTO ra_Documentos_web(rut,codcarr,documento ,ANO ,PERIODO ,fecha ,estado,tipo)VALUES('"& Rut &"','"& Codcarr &"','CUPONMATRICULA',"& ano &","& periodo &",GETDATE(),'EN PROCESO','" & session("tipo") & "') "
					Session("Conn").execute(str)

					strident= "select TOP 1 id codigo from ra_Documentos_web ORDER BY id DESC"  
					set rst = Session("Conn").execute(strident)
					MatriculaBanco = rst("codigo")  
					
					strexec="sp_genera_documentos_web "& MatriculaBanco &""
					Session("Conn").execute(strexec)
					
					session("MatriculaBanco")="Documentos/" & Rut & MatriculaBanco & ".pdf"
					
				elseif Matricula ="SERVIPAG" then
				
					sqlCreaPago = "Exec Pr_PagoWeb_Crea_Intencion @Rut = " + Rut
					Set RstCreaPago = Session("Conn").Execute(sqlCreaPago)
					
					SqlID = "Select Convert(Varchar, Coalesce(@@Identity, 0)) as IDPAGOWEB "
					Set RstID = Session("Conn").Execute(SqlID)	 
					IDPAGOWEB = RstID("IDPAGOWEB")
					
					Sql = "INSERT INTO mt_pago_web (id, rut)  SELECT id, rut FROM mt_pago_web_corr WHERE id = " & IDPAGOWEB
					Session("Conn").Execute Sql
					
					str="INSERT INTO ra_Documentos_web( rut,codcarr ,documento ,ANO ,PERIODO ,fecha ,estado,tipo,idpagoweb)VALUES('"& Rut &"','"& Codcarr &"','SERVIPAGMATRICULA',"& ano &","& periodo &",GETDATE(),'EN PROCESO','" & session("tipo") & "'," & IDPAGOWEB & ") "
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
				elseif Matricula ="PAGARE" then
				
					str="INSERT INTO ra_Documentos_web( rut,codcarr ,documento ,ANO ,PERIODO ,fecha ,estado,tipo)VALUES('"& Rut &"','"& Codcarr &"','PAGAREMATRICULA',"& ano &","& periodo &",GETDATE(),'EN PROCESO','" & session("tipo") & "') "
					Session("Conn").execute(str) 
					
					strident=  "select TOP 1 id codigo from ra_Documentos_web ORDER BY id DESC"
					set rst = Session("Conn").execute(strident)
					MatriculaPagare = rst("codigo") 
					 
					strexec="sp_genera_documentos_web "& MatriculaPagare &""
					Session("Conn").execute(strexec) 
					
					session("MatriculaPagare")="Documentos/" & Rut & MatriculaPagare & ".pdf"
				
				end if 
				
				if Arancel="BANCO" then
				
					str="INSERT INTO ra_Documentos_web( rut,codcarr ,documento ,ANO ,PERIODO ,fecha ,estado,tipo)VALUES('"& Rut &"','"& Codcarr &"','CUPONARANCEL',"& ano &","& periodo &",GETDATE(),'EN PROCESO','" & session("tipo") & "') "
					Session("Conn").execute(str)
					
					strident= "select TOP 1 id codigo from ra_Documentos_web ORDER BY id DESC"  
					set rst = Session("Conn").execute(strident)
					ArancelBanco = rst("codigo")
					
					strexec="sp_genera_documentos_web "& ArancelBanco &""
					Session("Conn").execute(strexec)
					
					session("ArancelBanco")="Documentos/" & Rut & ArancelBanco & ".pdf"
					
				elseif Arancel ="SERVIPAG" then 
				
					sqlCreaPago = "Exec Pr_PagoWeb_Crea_Intencion @Rut = " + Rut
					Set RstCreaPago = Session("Conn").Execute(sqlCreaPago)
					
					SqlID = "Select Convert(Varchar, Coalesce(@@Identity, 0)) as IDPAGOWEB "
					Set RstID = Session("Conn").Execute(SqlID)	 
					IDPAGOWEB = RstID("IDPAGOWEB")
					
					Sql = "INSERT INTO mt_pago_web (id, rut)  SELECT id, rut FROM mt_pago_web_corr WHERE id = " & IDPAGOWEB
					Session("Conn").Execute Sql
					
					str="INSERT INTO ra_Documentos_web( rut,codcarr ,documento ,ANO ,PERIODO ,fecha ,estado,tipo,idpagoweb)VALUES('"& Rut &"','"& Codcarr &"','SERVIPAGARANCEL',"& ano &","& periodo &",GETDATE(),'EN PROCESO','" & session("tipo") & "'," & IDPAGOWEB & ") "
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
				
					str="INSERT INTO ra_Documentos_web( rut,codcarr ,documento ,ANO ,PERIODO ,fecha ,estado,tipo)VALUES('"& Rut &"','"& Codcarr &"','PAGARE',"& ano &","& periodo &",GETDATE(),'EN PROCESO','" & session("tipo") & "') "
					Session("Conn").execute(str) 
					
					strident=  "select TOP 1 id codigo from ra_Documentos_web ORDER BY id DESC"
					set rst = Session("Conn").execute(strident)
					ArancelPagare = rst("codigo") 
					 
					strexec="sp_genera_documentos_web "& ArancelPagare &""
					Session("Conn").execute(strexec) 
					
					session("ArancelPagare")="Documentos/" & Rut & ArancelPagare & ".pdf"
				
				end if  
			end if 
			 
			if Matricula <> "SERVIPAG" and Arancel <> "SERVIPAG" then 
				response.Redirect("PostulResultadoarancelymatricula.asp") 
			end if 
			%>
        <table width="650" border="0" cellpadding="0" cellspacing="0" >
            <tr valign="top" >
                <td width="425" height="100">
                    <h1>Validando Pago...</h1>
                    <h1 style="font-size:20px"></h1>
                </td>
                <td width="66">&nbsp;</td>
                <td width="159"><img src="imagenes/2_r1_c2.jpg" height="130" width="150"></a></td>
            </tr>
        </table>	
           
		</div>
		<div id="footer"></div>
	</div>
</body>

</html>