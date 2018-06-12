<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="include/conexion.inc" -->

<%
strDocUarm =  "SELECT coalesce(dbo.Fn_ValorParame('CUPONESMATRICULAUARM'),'NO')Parame"
set rstDocUarm = Session("Conn").execute(strDocUarm)
CUPONESMATRICULAUARM = rstDocUarm("Parame")
%>
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
<script type="text/javascript">
function submitform()
{
	RbMatriculaBanco = document.getElementById('RbMatriculaBanco').checked;
	
	<%if CUPONESMATRICULAUARM <> "SI" then%>
		RbMatriculaSevipag = document.getElementById('RbMatriculaSevipag').checked;
		RbMatriculaPagare = document.getElementById('RbMatriculaPagare').checked;
		
		RbArancelBanco = document.getElementById('RbArancelBanco').checked;
		RbArancelSevipag = document.getElementById('RbArancelSevipag').checked;
		RbArancelPagare = document.getElementById('RbArancelPagare').checked;
	<%else%>
		RbMatriculaSevipag = false;
		RbMatriculaPagare = false;
		
		RbArancelBanco = true;
		RbArancelSevipag = false;
		RbArancelPagare = false;
	<%end if%>
	
	if (RbMatriculaBanco == false && RbMatriculaSevipag == false && RbMatriculaPagare == false)
	{
		alert("Debe Seleccionar una opci\u00f3n de pago")
	}
	else if (RbArancelBanco == false && RbArancelSevipag == false && RbArancelPagare == false)
	{
		alert("Debe Seleccionar una opci\u00f3n de pago")
	}
	else
	{
		document.myform.submit();
	}	
}
</script>
<body>
	<div id="container">
		<div id="navigation">
			<ul>
				
			</ul>
		</div>
		<div id="content">
        <%
		ano =AnoAdmision()
		periodo=PeriodoAdmision()
		
		VistaDocumentos=0
		Mensaje =""
		NOMBRE=""
		CARRERA=""
		rut = request("rut")
		codcarr = request("codcarr")
		
		if rut="" then
			rut=0 
		end if  
		
		session("RutPostul")=rut 
		
		strEsPoA = "select top 1 1 from mt_alumno where rut ='"& rut &"' and estacad <> 'ELIMINADO' "'and ano="& ano &" and periodo="& periodo &"  "
		set RstEsPoA = Session("Conn").Execute (strEsPoA) 						
		if RstEsPoA.eof then ' Es Postulante
			TIPO="Postulante"
			session("tipo")="POSTULANTE"
			if codcarr = "" then
			
				str="select count(*)Cantidad  from MT_POSCAR p,mt_client c WHERE p.CODPOSTUL=c.codcli AND p.CODPOSTUL='" & rut & "' "			
				set Rst = Session("Conn").Execute (str) 		
				
				if not Rst.eof then
				 
					strB="select top 1 1 from MT_MOROSOS where rut='" & rut & "'"
					set RstB = Session("Conn").Execute (strB)
					 
					if RstB.eof then 
					
						strC="select count(*)Cantidad from MT_POSCAR p,mt_client c WHERE p.CODPOSTUL=c.codcli AND p.CODPOSTUL='" & rut & "' and ESTADO='A' "
						set RstC = Session("Conn").Execute (strC)
						if not RstC.eof then
						
							if RstC("Cantidad") <> 0 then											
								if RstC("Cantidad")>1 then 
									response.Redirect("PosulElijeCarrera.asp?rut=" & rut & "")
								else
									strA = "select carr.CODCARR,c.nombre +' '+c.paterno NOMBRE,carr.nombre_l CARRERA from MT_POSCAR p,mt_client c,mt_Carrer carr "
									strA = strA & "WHERE p.CODPOSTUL=c.codcli AND p.CODPOSTUL='" & rut & "' and p.ESTADO='A' AND carr.CODCARR=p.codcarr "
									
									set RstA = Session("Conn").Execute (strA)
									if not RstA.eof then 
										NOMBRE = RstA("NOMBRE")
										CARRERA = RstA("CARRERA")
										session("codcarrPostul")=RstA("CODCARR")
									end if							 
									VistaDocumentos=1
								end if 
							else 
								Mensaje ="No tiene carreras posibles de matricular. Por favor verifique en m&oacute;dulos de admisi&oacute;n."
								VistaDocumentos=0
							end if
						else
							Mensaje ="No tiene carreras posibles de matricular. Por favor verifique en m&oacute;dulos de admisi&oacute;n."
							VistaDocumentos=0
						end if 
					else
						Mensaje ="Usted tiene compromisos financieros o de biblioteca pendientes. Por favor diríjase a la unidad de finanzas a revisar su situación."
						VistaDocumentos=0
					end if 
				else				
					Mensaje ="El Rut ingresado no corresponde a un postulante v&aacute;lido."
					VistaDocumentos=0
				end if
			else 
				strA = "select carr.CODCARR, c.nombre +' '+c.paterno NOMBRE,carr.nombre_l CARRERA from MT_POSCAR p,mt_client c,mt_Carrer carr WHERE "
				strA = strA & "p.CODPOSTUL=c.codcli AND p.CODPOSTUL='" & rut & "' AND p.codcarr='" & codcarr & "'and p.ESTADO='A' AND carr.CODCARR=p.codcarr"
				
				set RstA = Session("Conn").Execute (strA)
				if not RstA.eof then
					NOMBRE = RstA("NOMBRE")
					CARRERA = RstA("CARRERA")
					session("codcarrPostul")=RstA("CODCARR")
				end if 
				VistaDocumentos=1
			end if  
		else ' Es Alumno
			TIPO="Alumno"
			session("tipo")="ALUMNO"
			if codcarr = "" then
				 
				Sstr=" exec sp_alumno_deuda '"& (rut) &"'"
				Session("Conn").Execute Sstr 
	
				Sstr=" Select Deuda from mt_client Where Codcli = '"& (rut) &"'"
				Set Rrst = Session("Conn").execute(Sstr)
				
				DEUDA = valnulo(Rrst("Deuda"),NUM_)
								 
				if DEUDA = 0 then 
					
					strC = "Select COUNT(codcarpr)CARRERAS from mt_alumno where rut ='"& (rut) &"'"
					strC = strC & "and (coalesce(ano_mat,0) <> "& ano &""
					strC = strC & "or coalesce(periodo_mat,0) <> " & periodo & ") and estacad in ('VIGENTE', 'EGRESADO') "
					
					set RstC = Session("Conn").Execute (strC)
					if not RstC.eof then
															
						if RstC("CARRERAS")>1 then 
							response.Redirect("PosulElijeCarrera.asp?rut=" & rut & "")
						else
							strA = "select carr.CODCARR,c.nombre +' '+c.paterno NOMBRE,carr.nombre_l CARRERA from mt_alumno a,mt_client c,mt_Carrer carr "
							strA = strA & "WHERE a.rut=c.codcli AND a.rut='" & rut & "' AND carr.CODCARR=a.codcarpr  and estacad in ('VIGENTE', 'EGRESADO') "

							set RstA = Session("Conn").Execute (strA)
							if not RstA.eof then 
								NOMBRE = RstA("NOMBRE")
								CARRERA = RstA("CARRERA")
								session("codcarrPostul")=RstA("CODCARR")
							end if							 
							VistaDocumentos=1
						end if 
					end if 
				else
					Mensaje ="Usted tiene compromisos financieros o de biblioteca pendientes. Por favor diríjase a la unidad de finanzas a revisar su situación."
					VistaDocumentos=0
				end if 
			else 
				strA = "select carr.CODCARR,c.nombre +' '+c.paterno NOMBRE,carr.nombre_l CARRERA from mt_alumno a,mt_client c,mt_Carrer carr "
				strA = strA & "WHERE a.rut=c.codcli AND a.rut='" & rut & "' AND a.codcarpr='" & codcarr & "' AND carr.CODCARR=a.codcarpr "
				
				set RstA = Session("Conn").Execute (strA)
				if not RstA.eof then 
					NOMBRE = RstA("NOMBRE")
					CARRERA = RstA("CARRERA")
					session("codcarrPostul")=RstA("CODCARR")
				end if							 
				VistaDocumentos=1
			end if  
		
		end if 
		
		if VistaDocumentos = 1 then%>
        <table width="650" border="0" cellpadding="0" cellspacing="0" >
					<tr valign="top" >
					  <td width="425" height="100">
					    <h1 style="font-size:22px"><%=TIPO%> v&aacute;lido para la Carrera <%=CARRERA%>.</h1>
                        <h1 style="font-size:20px"><%=NOMBRE%></h1>
            <br />
			<p align="justify" style="font-size:16px">Seleccione un medio de pago.</p>
            <br />
            <!--<table width="400" border="1" cellpadding="1" cellspacing="1" >
                <tr>
                    <td width="67"><a style="font-size:18px" href="PostulantePagare.asp">Pagar&eacute;</a></td>
                </tr>
                <tr>
                    <td><a style="font-size:18px" href="PostulanteCuponMatricula.asp">Cup&oacute;n Matr&iacute;cula</a></td>
                </tr>
                <tr>
                    <td><a style="font-size:18px" href="PostulanteCuponArancel.asp">Cup&oacute;n Arancel</a></td>
                </tr>
            </table>-->
            <table width="500" border="0" cellpadding="0" cellspacing="0">
                  <form name="myform" id="myform" action="PostulProcesaarancelymatricula.asp" method="POST">                   
					
                    <tr valign="top" >
					  <td width="95" height="24">
					    <!--<p class="text-menu">Matr&iacute;cula</p>-->
                        <p class="text-menu">Arancel B&aacute;sico</p>
					  </td>
				<td width="165"><input id ="RbMatriculaBanco" type="radio" name="rbMatricula" value="BANCO" checked><b class="Tit-celdas">&nbsp; Banco (Contado).</b></td>
				<td width="98"></td>                      
					</tr>  
                    <tr valign="top" >
					  <td width="95" height="24">					   
					  </td>
                <%if CUPONESMATRICULAUARM <> "SI" then%>     
				<td width="165"><input id ="RbMatriculaSevipag" disabled type="radio" name="rbMatricula" value="SERVIPAG"><b class="Tit-celdas">&nbsp; Servipag (Contado).</b></td>
				<td width="98"></td>                      
					</tr>    
                     <tr valign="top" >
					  <td width="95" height="24"> 					   
					  </td>
                <td width="165"><input id ="RbMatriculaPagare" type="radio" name="rbMatricula" value="PAGARE"> <b class="Tit-celdas">&nbsp; Pagar&eacute; (Cuotas).</b></td>
                <%end if%>
				<td width="98"></td>                      
					</tr>
                    <tr valign="top" >
					  <td width="95" height="24">					   
					  </td>
				<td width="165"></td>
				<td width="98"></td>                      
					</tr>
                    <%if CUPONESMATRICULAUARM <> "SI" then%>                  
                    <tr valign="top" >
					  <td height="24" >
					    <!--<p class="text-menu">Arancel</p>--> 
                        <p class="text-menu">Arancel de Matr&iacute;cula</p>
					  </td> 
					  <td><input id ="RbArancelBanco" type="radio" name="rbArancel" value="BANCO" checked><b class="Tit-celdas">&nbsp; Banco (Contado).</b> </td>
				      <td></td>
                      <td width="142"></td>
					</tr>   
                    <tr valign="top" >
					  <td height="24" >
					  </td> 
					  <td><input id ="RbArancelSevipag" type="radio" disabled name="rbArancel" value="SERVIPAG"><b class="Tit-celdas">&nbsp; Servipag (Contado).</b> </td>
				      <td></td>
                      <td width="142"></td>
					</tr> 
                    <tr valign="top" >
					  <td height="24" >
					  </td> 
					  <td><input id ="RbArancelPagare" type="radio" name="rbArancel" value="PAGARE"> <b class="Tit-celdas">&nbsp; Pagar&eacute; (Cuotas).</b></td>
				      <td></td>
                      <td width="142"></td>
					</tr>  
                     <%end if%>                
                    </form>
                    <tr valign="top" >
					  <td height="24"><input  style="background-color: #eeeeee; border: 1px black outset" type="button" value=" Volver " onclick="location.href = 'documentospostulantes.asp'">
                      </td>
					  <td></td>
                      <td>&nbsp;<input  style="background-color: #eeeeee; border: 1px black outset" type="button" value=" Continuar " onclick="javascript: submitform()"></td> 
				      <td></td>                      
					</tr>
			    </table>
            </td>
					  <td width="66">&nbsp;</td>
				      <td width="159"><img src="imagenes/2_r1_c2.jpg" height="130" width="150"></a></td>
					</tr>
                    </table>			
            <br />
			<br />	
        <%else%>
        <table width="650" border="0" cellpadding="0" cellspacing="0" >
            <tr valign="top" >
                <td width="425" height="100">
                    <h1  style="text-align:justify"><%=Mensaje%></h1>                    
                    <br /></td>
                <td width="66">&nbsp;</td>
                <td width="159"><img src="imagenes/2_r1_c2.jpg" height="130" width="150"></a></td>
            </tr>
        </table>			
        <br />
        <br />
        <input  style="background-color: #eeeeee; border: 1px black outset" type="button" value=" Volver " onclick="location.href = 'documentospostulantes.asp'">  			
        <%end if %>		
		</div>
		<div id="footer"></div>
	</div>
</body>

</html>