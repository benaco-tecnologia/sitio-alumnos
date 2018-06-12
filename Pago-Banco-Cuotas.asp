<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" --> 
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->

<%  response.buffer = false 
    Response.Expires = -1 

    if Session("CodCli") = "" then
      Response.Redirect("saltoinicio.htm")
    end if       
	
	if EstaHabilitadaNW (485)="S" then 
		if GetPermisoNW(485) ="N" then
			response.Redirect("MensajesBloqueos.asp")	
		end if
	else
		response.Redirect("MensajeBloqueoHabilita.asp")
	end if
 
    Audita 485,"Ingresa a Consulta de Cuentas"
%>



<% 
' GENERA ORDEN DE COMPRA A PARTIR DE FECHA
FECHAACTUAL = NOW()
ANO = YEAR(FECHAACTUAL)
MES = MONTH(FECHAACTUAL)
DIA = DAY(FECHAACTUAL)
MINUTO = MINUTE(FECHAACTUAL)
SEGUNDO = SECOND(FECHAACTUAL)
OC = "OC_"&ANO&MES&DIA&MINUTO&SEGUNDO

codcli=Session("CodCli")


dim strsql,strsql2  
dim montototal    
dim montoacancelar
dim montoacancelarcuota  
dim maxfilas
dim feccorte
dim cuotas
dim total_cuotas

feccorte = Date()  

maxfilas=0  
montototal = 0
montoacancelar = 0 

' TRAE RUT PARA EL CODIGO DEL ALUMNO DE LA CARRERA
strsql= "select a.rut, a.codcarpr, c.nombre_l, a.ano from mt_alumno a, mt_carrer c where a.codcli = '"& codcli & "' and a.codcarpr = c.codcarr"
Set Rst = Session("Conn").Execute(StrSql)
if not rst.eof then
	rutcli =rst("rut")
	codcarr =rst("codcarpr")
	nombrecarrera =rst("nombre_l")	
	anoing =rst("ano")		
end if

' CONSULTA DATOS DE ALUMNO PARA ENCABEZADO
strsql = "select codcli,dig,nombre,paterno,materno,diractual,comuna,ciudadact,fonoact,codaval,fecnac,codsissalud, Coalesce(codaval, '') as codaval,Coalesce(codapod, '') as codapod from mt_client  where codcli = '" & rutcli &"'"
	Set Rst=Session("Conn").Execute(StrSql)
	if not rst.eof then
		alumrut =rst("codcli")
		RUT =rst("codcli")
		alumdv =rst("dig")
		alumnombre =rst("nombre")
		alumpaterno =rst("paterno")
		alummaterno =rst("materno")
		apoderado = (rst("codapod"))
	end if

	' TRAE DATOS DEL APODERADO
	if apoderado = "" then 
	    apodrut = "0"
		apoddv = ""
		apodpaterno = ""
		apodmaterno = ""
		apodnombres = ""
		apoddireccion = ""
		apodcomuna = ""
		apodciudad = ""
		apodtelefono = ""
		apodcelular = ""
		apodmail = ""
	else
	
	    strsql= "select * from mt_apoder where codapod = '"+ apoderado + "'"
	    Set Rst = Session("Conn").Execute(StrSql)
	    if not rst.eof then
		    apodrut =rst("codapod")
		    apoddv =rst("dig")
		    apodpaterno =rst("paterno")
		    apodmaterno =rst("materno")
		    apodnombres =rst("nombre")
		    apoddireccion =rst("dir_part")
		    apodcomuna =rst("comu_part")
		    apodciudad =rst("ciu_part")
		    apodtelefono =rst("tel_part")
		    apodcelular =rst("celular")
		    apodmail =rst("mail")
	    end if
    end if

' CONSULTA DATOS DE ALUMNO PARA CTA.CTE.
strsql= "select  c.ctadoc, c.ctadocnum, b.nombre, c.codban, c.codsuc, c.fecven, gastosop, isnull(c.saldo,0) saldo, c.estado, i.descripcion as itemdesc, por1, por2, por3, "
strsql= strsql + " por4, c.estado, est.descripcion as estadodesc, c.feccancel, c.codapod, c.IntDeuda, c.monto, i.coditem, c.num_operacion, c.TipoPagCond, i.activa,"
strsql= strsql + " c.item, (c.monto - c.saldo) pago, (c.monto + c.intdeuda) montoacancelar, d.nombre as ubicacion, c.gastospro, c.intrepac, datediff(d, c.fecven,'" 
strsql= strsql + cstr(feccorte) + "') fechadiff, c.ID_CUOTA "
strsql = strsql + " ,CASE WHEN dbo.Fn_ValorParame('Interesdiario') = 'N' THEN COALESCE(dbo.InteresMora(c.Saldo, c.FecVen, GETDATE(), -1),0) WHEN dbo.Fn_ValorParame('Interesdiario') = 'S' THEN COALESCE(dbo.Interespordia(c.Saldo, c.FecVen, CONVERT(VARCHAR(8), GETDATE(), 112), -1 ),0)	 ELSE 0 	END AS mis_interes "
strsql = strsql + ", dbo.fn_Obtiene_gastos_cob(c.ID_CUOTA) as Gastos_Cob,c.VCTOORI "
strsql= strsql + " from mt_ctadoc c ,mt_item i, mt_estdoc est, mt_docum b, mt_detubi d"
strsql= strsql + " where codcli = '" + rutcli + "' and saldo > 0 "
strsql= strsql + " and c.codcarr = '" + codcarr + "'"
strsql= strsql + " and c.item*=i.coditem"
strsql= strsql + " and c.estado=est.estado"
strsql= strsql + " and c.ctadoc  = b.tipodoc"
StrSql= StrSql + " and c.ubicacion = d.codubifis    "
StrSql = StrSql + " and b.PagoWeb = 'SI' " 'Sólo aplica para la cancelación de cuotas a través del portal
strsql= strsql + " order by fecreg desc ,CTADOC,fecven, ctadocnum "



'response.Write(strsql)
'response.End
dim HayRegistros

Set Rst = Session("Conn").Execute(StrSql)  

if not rst.eof  then 
  CuentasPay = rst.getrows  
  maxfilas = Ubound(CuentasPay, 2) 
  HayRegistros = 1
  'response.Write("max:" + cstr(maxfilas))
else
    HayRegistros = 0
end if 
%>
<html>
<script language="JavaScript" type="text/javascript">

function CargaCupon(ID)
{

	var opciones="toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=no, resizable=no, width=800, height=540, top=85, left=250";
	window.open("ProcesaCuponCuota.asp?ID=" + ID ,"",opciones);
}

</script>
<head>
    <title>.:: EMITE CUOTAS POR INTERNET ::.</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
    <link href="css/tex-normales.css" rel="stylesheet" type="text/css" />
    <link rel="stylesheet" href="css/tablas.css" type="text/css" />
    <link href="estilos_mas.css" rel="stylesheet" type="text/css">
</head>


<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <form name="datos" id="datos" method="post" action="browse.asp">
    <table border="0" cellpadding="0" cellspacing="0" align="left">
        <tr>
            <table width="200" border="0" cellpadding="0" cellspacing="0" align="left">
                <tr>
                    <td colspan="2" valign="top" align="right">
                        <%CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
                        %>
                    </td>
                    <td valign="top">
                        <% CargarTop1()%><% SubMenu()%>
                        <!-- codigo sin formato U++ -->
                        <table>
                            <tr>
                                <td>
                                    <table width="834" border="0">
                                        <tr>
                                            <td>
                                                <p>
                                                    <img src="Imagenes/titulos/T-Consulta-Cuentas.gif" width="271" height="38" /></p>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td width="100%">
                                                <tr>
                                                    <td colspan="2" height="30">
                                                        <p>
                                                            <span style="font-size: 13px; font-weight: bold; font-family: Verdana, Arial, Helvetica, sans-serif;">
                                                                <b class="Tit-celdas"><font face="Verdana, Arial, Helvetica, sans-serif">Datos del alumno</font></b></span>
                                                        </p>
                                                        <table width="834" border="1">
                                                            <tr>
                                                                <th background="imagenes/fdo-cabecera-cel22.jpg" width="70" bgcolor="4A5DA1" scope="col">
                                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                        Rut</font></strong>
                                                                </th>
                                                                
                                                                <th background="imagenes/fdo-cabecera-cel22.jpg" width="130" bgcolor="4A5DA1" scope="col">
                                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                        Nombres</font></strong>
                                                                </th>
                                                                <th background="imagenes/fdo-cabecera-cel22.jpg" width="130" bgcolor="4A5DA1" scope="col">
                                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                        Apellido Paterno</font></strong>
                                                                </th>
                                                                <th background="imagenes/fdo-cabecera-cel22.jpg" width="130" bgcolor="4A5DA1" scope="col">
                                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                        Apellido Materno</font></strong>
                                                                </th>
                                                                <th background="imagenes/fdo-cabecera-cel22.jpg" width="90" bgcolor="4A5DA1" scope="col">
                                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                        C&oacute;digo Carrera</font></strong>
                                                                </th>
                                                                <th background="imagenes/fdo-cabecera-cel22.jpg" width="130" bgcolor="4A5DA1" scope="col">
                                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                        Nombre Carrera</font></strong>
                                                                </th>
                                                                <th background="imagenes/fdo-cabecera-cel22.jpg" width="82" bgcolor="4A5DA1" scope="col">
                                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                        A&ntilde;o Ingreso</font></strong>
                                                                </th>
                                                            </tr>
                                                            <tr bgcolor="#DBECF2">
                                                                <td width="70" height="32">
                                                                    <div align="center">
                                                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                                                            <%=FormatNumber(alumrut, 0) + "-" + UCase(alumdv)%></font></div>
                                                                </td>
                                                                
                                                                <td width="130" height="32">
                                                                    <div align="center">
                                                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                                                            <%=alumnombre%></font></div>
                                                                </td>
                                                                <td width="130" height="32">
                                                                    <div align="center">
                                                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                                                            <%=alumpaterno%></font></div>
                                                                </td>
                                                                <td width="130" height="32">
                                                                    <div align="center">
                                                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                                                            <%=alummaterno%></font></div>
                                                                </td>
                                                                <td width="90" height="32">
                                                                    <div align="center">
                                                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                                                            <%=codcarr%></font></div>
                                                                </td>
                                                                <td width="130" height="32">
                                                                    <div align="center">
                                                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                                                            <%=nombrecarrera%></font></div>
                                                                </td>
                                                                <td width="82" height="32">
                                                                    <div align="center">
                                                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                                                            <%=anoing%></font></div>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                        <p>
                                                            <span style="font-size: 13px; font-weight: bold; font-family: Verdana, Arial, Helvetica, sans-serif;">
                                                                <b class="Tit-celdas"><font face="Verdana, Arial, Helvetica, sans-serif">Antecedentes
                                                                    Responsable Financiero</font></b></span>
                                                        </p>
                                                        <table width="610" border="1">
                                                            <tr>
                                                                <th background="imagenes/fdo-cabecera-cel22.jpg" width="70" bgcolor="4A5DA1" scope="col">
                                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                        Rut</font></strong>
                                                                </th>
                                                               
                                                                <th background="imagenes/fdo-cabecera-cel22.jpg" width="130" bgcolor="4A5DA1" scope="col">
                                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                        Nombres</font></strong>
                                                                </th>
                                                                <th background="imagenes/fdo-cabecera-cel22.jpg" width="130" bgcolor="4A5DA1" scope="col">
                                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                        Apellido Paterno</font></strong>
                                                                </th>
                                                                <th background="imagenes/fdo-cabecera-cel22.jpg" width="130" bgcolor="4A5DA1" scope="col">
                                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                        Apellido Materno</font></strong>
                                                                </th>
                                                                <th background="imagenes/fdo-cabecera-cel22.jpg" width="90" bgcolor="4A5DA1" scope="col">
                                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                        Tel&eacute;fono</font></strong>
                                                                </th>
                                                            </tr>
                                                            <tr bgcolor="#DBECF2">
                                                                <td width="70" height="32">
                                                                    <div align="center">
                                                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                                                            <%=FormatNumber(apodrut, 0) + "-" + UCase(apoddv)%></font></div>
                                                                </td>
                                                                
                                                                <td width="130" height="32">
                                                                    <div align="center">
                                                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                                                            <%=apodnombres%></font></div>
                                                                </td>
                                                                <td width="130" height="32">
                                                                    <div align="center">
                                                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                                                            <%=apodpaterno%></font></div>
                                                                </td>
                                                                <td width="130" height="32">
                                                                    <div align="center">
                                                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                                                            <%=apodmaterno%></font></div>
                                                                </td>
                                                                <td width="90" height="32">
                                                                    <div align="center">
                                                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                                                            <%=apodtelefono%></font></div>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                        <br />
                                                        <tr>
                                                            <td>
                                                                <p>
                                                                    
                                                                    <span style="font-size: 13px; font-weight: bold; font-family: Verdana, Arial, Helvetica, sans-serif;">
                                                                        <b class="Tit-celdas">
                                                                        <font face="Verdana, Arial, Helvetica, sans-serif">Detalle de cuenta corriente</font>
                                                                            <br>
                                                                            <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#666666">(Este documento
                                                                                NO constituye certificado)</font> <font face="Verdana, Arial, Helvetica, sans-serif"
                                                                                    color="#FFFFFF"><b><font size="1">
                                                                                        <br />
                                                                </p>
                                                                <!--<p>
                                                                    <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#666666">Seleccione
                                                                        las cuotas a pagar</font>
                                                                </p>-->
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>
                                                                <table border="0" cellpadding="0" cellspacing="0" height="246" align="left" width="899">
                                                                    <tr valign="top">
                                                                        <td width="899" height="70" colspan="1">
                                                                            <table width="900" cellspacing="0" cellpadding="0" height="72" border="1" bordercolor="#FFFFFF">
                                                                                <tr background="imagenes/fdo-cabecera-cel22.jpg">
                                                                                    <!--<td background="imagenes/fdo-cabecera-cel22.jpg" width="30" height="30">
                                                                                        <div align="center">
                                                                                            <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                                                Pagar</font></strong></div>
                                                                                    </td>-->
                                                                                    <td background="imagenes/fdo-cabecera-cel22.jpg" width="90" height="30">
                                                                                        <div align="center">
                                                                                            <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                                                Vencimiento</font></strong></div>
                                                                                    </td>
                                                                                    <td background="imagenes/fdo-cabecera-cel22.jpg" width="60" height="30">
                                                                                        <div align="center">
                                                                                            <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                                                Documento</font></strong></div>
                                                                                    </td>
                                                                                    <td background="imagenes/fdo-cabecera-cel22.jpg" width="80" height="30">
                                                                                        <div align="center">
                                                                                            <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                                                Monto a Cancelar</font></strong>
                                                                                        </div>
                                                                                    </td>
                                                                                    <td background="imagenes/fdo-cabecera-cel22.jpg" width="60" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">
                                                                                                Correlativo</font></b></font>
                                                                                        </div>
                                                                                    </td>
                                                                                    <td background="imagenes/fdo-cabecera-cel22.jpg" width="60" height="30">
                                                                                        <div align="center">
                                                                                            <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                                                Monto Capital</font></strong></div>
                                                                                    </td>
                                                                                    <td background="imagenes/fdo-cabecera-cel22.jpg" width="60" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">
                                                                                                Int.Mora</font></b></font></div>
                                                                                    </td>
                                                                                    <td background="imagenes/fdo-cabecera-cel22.jpg" width="60" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">
                                                                                                Gastos Cobranza</font></b></font></div>
                                                                                    </td>
                                                                                    <td background="imagenes/fdo-cabecera-cel22.jpg" width="80" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">
                                                                                                Pagos</font></b></font></div>
                                                                                    </td>
                                                                                    <td background="imagenes/fdo-cabecera-cel22.jpg" width="90" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">
                                                                                                Estado</font></b></font></div>
                                                                                    </td>
                                                                                    <td background="imagenes/fdo-cabecera-cel22.jpg" height="30" width="90">
                                                                                        <div align="center">
                                                                                            <font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">
                                                                                                Ubicaci&oacute;n</font></b></font></div>
                                                                                    </td>
                                                                                    <td background="imagenes/fdo-cabecera-cel22.jpg" height="30" width="90">
                                                                                        <div align="center">
                                                                                            <font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">
                                                                                                Cup&oacute;n</font></b></font></div>
                                                                                    </td>
                                                                                </tr>
                                                                                <form name="form1" method="POST" action="https://200.71.198.46/Alumnos/pago_servipag.asp?codcli=<%=RUT%>">
                                                                                <%
			pos = 0 
			if maxfilas >= 0 And HayRegistros = 1 then 
				do while pos <= maxfilas
				montoacancelarcuota = 0
				
				'mejora antes era solo montoacancelar = montoacancelar  + ccur(CuentasPay(19, pos))
				montoacancelar = montoacancelar  + (ccur(CuentasPay(19, pos)) - ccur(CuentasPay(25, pos))) 
				montoacancelarcuota = montoacancelarcuota  + ccur(CuentasPay(19, pos)) + ccur(CuentasPay(33, pos)) + ccur(CuentasPay(32, pos))
				
				'revisar el tema de los intereses y cuando aplicarlos
				' REVISA CARGO DE INTERESES				
				if ccur(CuentasPay(7, pos)) = 0 then
					intrec  = 0
					recargos = ccur(CuentasPay(28, pos))
				else
			  		'intrec = 0 'ccur(CuentasPay(29, pos)) + GastosCobranza(CuentasPay(0, pos), CuentasPay(1, pos), CuentasPay(19, pos), CuentasPay(5, pos))
			  		interc = 0' vamos a asumir que no hay gastos de cobranzas
			  		
					if CuentasPay(5, pos) = "" then
						fecpago = CuentasPay(16, pos)
					else
						fecpago = CuentasPay(5, pos)
					end if
	
					diasdif = CuentasPay(30, pos)
					recargos = ccur(CuentasPay(32, pos))' + ccur(CuentasPay(33, pos)) 'intereses + gastos de cobranza

					'intrec = intrec + recargos
					intrec = cdbl(recargos)
				end if				
				
			  	intrectotal = intrectotal + intrec 
            	monto = monto + ccur(CuentasPay(19, pos)) 
				montototal = montototal + monto + ccur(CuentasPay(33, pos))
            	saldo = saldo + ccur(CuentasPay(7, pos))
				saldototal = saldototal + saldo
				
				montoacancelarcuota = montoacancelarcuota  
				montoacancelar = montoacancelar  + intrec + Ccur(CuentasPay(33, pos))
				' REVISA CARGO DE INTERESES
				
				
				strFecha="SELECT GETDATE()fecha"
				Set RstFecha = Session("Conn").Execute(strFecha)
				fechahoy=RstFecha("fecha")
				
				strdia="SELECT dbo.Fn_ValorParame('diasatr')dia "
				Set Rstdia = Session("Conn").Execute(strdia)
				diaAtraso=valnulo(Rstdia("dia"),NUM_)
				
				'Logica para morosos y vencidos				
				If DateDiff("d",CuentasPay(5, pos) , fechahoy) < 0 Then				
					If IsDate(CuentasPay(34, pos)) Then
						If CuentasPay(5, pos) > CuentasPay(34, pos) Then
							CuentasPay(15, pos)="Prorrogado-VIGENTE"
						Else
							CuentasPay(15, pos)="VIGENTE"
						End If
					Else
						CuentasPay(15, pos)="VIGENTE"
					End If
				ElseIf DateDiff("d",CuentasPay(5, pos) , fechahoy)>= 0 Then
					If DateDiff("d",CuentasPay(5, pos) , fechahoy) > diaAtraso Then
						If CuentasPay(5, pos) > CuentasPay(34, pos) Then
							CuentasPay(15, pos)="Prorrogado-MOROSO"
						Else
							CuentasPay(15, pos)="MOROSO"
						End If
					Else
						if CuentasPay(5, pos) > CuentasPay(34, pos) Then
							CuentasPay(15, pos)="Prorrogado-VENCIDO"
						Else
							CuentasPay(15, pos)="VENCIDO"
						End If
					End If
				End IF 

                                                                                %>
                                                                                <tr bgcolor="#DBECF2">
                                                                                    
                                                                                    <td width="63" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Arial, Helvetica, sans-serif" size="1">
                                                                                                <%=CuentasPay(5, pos)%></font></div>
                                                                                    </td>
                                                                                    <td width="90" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Arial, Helvetica, sans-serif" size="1">
                                                                                                <%=CuentasPay(2, pos)%></font></div>
                                                                                    </td>
                                                                                    <td width="60" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Arial, Helvetica, sans-serif" size="1">
                                                                                                <%=FormatNumber(montoacancelarcuota, 0)%></font></div>
                                                                                    </td>
                                                                                    <td width="80" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Arial, Helvetica, sans-serif" size="1">
                                                                                                <%=CuentasPay(1, pos)%></font></div>
                                                                                    </td>
                                                                                    <td width="60" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Arial, Helvetica, sans-serif" size="1">
                                                                                                <%=FormatNumber(CuentasPay(7, pos), 0)%></font></div>
                                                                                    </td>
                                                                                    <td width="60" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Arial, Helvetica, sans-serif" size="1">
                                                                                                <%=FormatNumber(intrec, 0)%></font></div>
                                                                                    </td>
                                                                                    <td width="60" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Arial, Helvetica, sans-serif" size="1">
                                                                                                <%=FormatNumber(CuentasPay(33, pos), 0)%></font></div>
                                                                                    </td>
                                                                                    <td width="60" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Arial, Helvetica, sans-serif" size="1">
                                                                                                <%=FormatNumber(CuentasPay(25, pos), 0)%></font></div>
                                                                                    </td>
                                                                                    <td width="80" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Arial, Helvetica, sans-serif" size="1">
                                                                                                <%=CuentasPay(15, pos)%></font></div>
                                                                                    </td>
                                                                                    <td width="90" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Arial, Helvetica, sans-serif" size="1">
                                                                                                <%=CuentasPay(27, pos)%>
                                                                                            </font>
                                                                                        </div>
                                                                                    </td>
                                                                                    <td width="90" height="30">
                                                                                        <div align="center"> 
                                                                                        <a href="" onClick="javascript:CargaCupon(<%=CuentasPay(31, pos)%>);">
                                                                                        <img src="Imagenes/pdf.png" height="30" width="30"></a>
                                                                                        </div>
                                                                                    </td>
                                                                                </tr>
                                                                                <% 
		  	pos = pos + 1
		    	loop  
		  end if 
                                                                                %>
                                                                            </table>
                                                                        </td>
                                                                    </tr>
                                                                    <tr valign="top">
                                                                        <td colspan="1" height="2">
                                                                            <table  width="100%" border="1">
                                                                                <tr>
                                                                                    <th background="imagenes/fdo-cabecera-cel22.jpg" width="88%" bgcolor="4a5da1" scope="col" align=right>
                                                                                        <div align="center" style="width: 152px">
                                                                                            <span class="style2"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"
                                                                                                size="1">Monto Total   </font> </span>
                                                                                        </div>
                                                                                    </th>
                                                                                    <td width="15%" height="32" bgcolor="#DBECF2">
                                                                                        <div align="center">
                                                                                            <font face="Arial, Helvetica, sans-serif" size="2">
                                                                                                <%=FormatNumber(montoacancelar, 0)   %></font></div>
                                                                                    </td>
                                                                                    

                                                                                </tr>
                                                                            </table>
                                                                        </td>
                                                                    </tr>
                                                                    <!--
                                                                    <tr>
                                                                        <td align="right">
                                                                            <a onclick="cuenta_cuotas()" href="#">                                                                                
                                                                                <img src="img/boton-sp.gif" alt="Pagar en servipag" name="boton" width="200" height="70"
                                                                                    border="0" id="Img2" />
                                                                            </a>
                                                                        </td>
                                                                    </tr>
                                                                    -->
                                                                    
                                                                </table>
                                                            </td>
                                                        </tr>
                                                </tr>
                                            </td>
                                        </tr>
                                    </table>
                                </td>
                            </tr>
                    </td>
                </tr>
            </table>    
        <b class="Tit-celdas">
        <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#666666">La información contenida en este documento no tiene valor legal, es meramente referencial.</font> 
        </br>
        </br>
    </table>
    </br>
    </br>
    <!-- fin codigo sin formato U++ -->
    </form>
</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
