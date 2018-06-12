<%  if Session("CodCli") = "" then
      Response.Redirect("saltoinicio.htm")
    end if        
	
	if EstaHabilitadaNW (486)="S" then 
		if GetPermisoNW(486) ="N" then
		  response.Redirect("MensajesBloqueos.asp")	
		end if
	else
		response.Redirect("MensajeBloqueoHabilita.asp")
	end if
	
    Audita 486,"Ingresa a Pago de cuotas"
    
    var_servidor = "200.54.71.237" 
    
    Dim IdPagoWeb 
    
	strPD="SELECT DBO.FN_MONEDA_USADECIMAL(1) DECIMAL"
	if BCL_ADO(strPD, rstPD) then
		USADECIMAL = rstPD("DECIMAL")
	end if
	
	if USADECIMAL = "SI" then
		CantDecimales = 2
	else
		CantDecimales = 0
	end if 

    'IdPagoWeb = Request("IDPAGOWEB")
	IdPagoWeb = session("IDPAGOWEBBP") 
    if IdPagoWeb = "" then 
        IdPagoWeb = 0
    end if
    
    if IdPagoWeb = 0 then 
        response.Write("Se sesión ha expirado") 
        response.End
    end if
    
    'response.Write(IDPAGOWEB)
    'response.End
    
    Dim CuentaCuotas, FechaVen
    Dim miSql, miRst 
    
    
    miSql = "Select Convert(Varchar(8), Coalesce(min(c.FecVen), getdate()), 112) as FechaVen from mt_ctadoc c, mt_pago_web_cuotas k where c.id_cuota = k.id_cuota and k.idPagoWeb = " & IDPAGOWEB
	'response.Write(miSql)

    Set miRst = Session("Conn").Execute(miSql)
    if not miRst.Eof then
        FechaVen = miRst(0)
    end if
    miRst.Close
    Set miRst = Nothing
        
    miSql = "Select Coalesce(count(*), 0) from mt_pago_web_cuotas Where IdPagoWeb = " & IdPagoWeb
	'response.Write(miSql)
    Set miRst = Session("Conn").Execute(miSql)
    if not miRst.Eof then 
        CuentaCuotas = miRst(0)
    End if
    miRst.Close
    Set miRst = Nothing
    
	if CantDecimales = 2 then
	    miSql = "SELECT Convert(varchar, Convert(decimal(12,2), Coalesce(dbo.fn_PagoWeb_Obtiene_monto_a_Pagar(" & IdPagoWeb & "), 0) ) )"
	else
		miSql = "SELECT Convert(varchar, Convert(decimal(10,0), Coalesce(dbo.fn_PagoWeb_Obtiene_monto_a_Pagar(" & IdPagoWeb & "), 0) ) )"
	end if 
    Set miRst = Session("Conn").Execute(miSql)
    'response.Write(miSql)
    if not miRst.Eof then 
        monto_a_pagar = miRst(0)
        Session("monto_a_pagar") = monto_a_pagar
    end if
    miRst.Close
    Set miRst = Nothing
    
    Dim FechaPago, Firma
    FechaPago = ""
    Firma = ""
    
    miSql = "Select Convert(varchar(8), Getdate(), 112) as FechaPago, Resultado From mt_pago_web_corr Where Id = " & IDPAGOWEB
    Set miRst = Session("Conn").Execute(miSql)
    if not miRst.Eof then 
        FechaPago = miRst(0)
        Firma = miRst(1)
    end if
    
    'response.Write(IDPAGOWEB + "*" + FechaVen + "*" + FechaPago + "*" + Firma)
    'response.End
    
    dim Sql 
    Sql = ""
    Sql = "<?xml version='1.0' encoding='ISO8859-1'?><Servipag><Header><FirmaEPS>" & Firma & "</FirmaEPS><CodigoCanalPago>454" 
    Sql = Sql +        "</CodigoCanalPago><IdTxCliente>" & IDPAGOWEB & "</IdTxCliente><FechaPago>" & FechaPago & "</FechaPago><MontoTotalDeuda>" & monto_a_pagar 
    Sql = Sql +    "</MontoTotalDeuda><NumeroBoletas>1</NumeroBoletas></Header><Documentos><IdSubTrx>1</IdSubTrx><CodigoIdentificador>" & IDPAGOWEB 
    Sql = Sql +        "</CodigoIdentificador><Boleta>1</Boleta><Monto>" & monto_a_pagar 
    Sql = Sql +        "</Monto><FechaVencimiento>" & FechaVen & "</FechaVencimiento></Documentos></Servipag>"

    Session("miXML") = Sql
    Session("monto_a_pagar") = trim(monto_a_pagar)
    Session("IDPAGOWEB") = trim(IDPAGOWEB)
    
   Function PagoIncluido(idpagoweb, id_cuota)
        Dim Sql, Rst, Aux
        
        Sql = "Select 1 From mt_pago_web_cuotas Where IdPagoWeb = " & idpagoweb & " and id_cuota = " & id_cuota
        Set Rst = Session("Conn").Execute(Sql)
        if not Rst.Eof then 
            Aux = "checked"
        Else
            Aux = "false"   
        end if
        
        PagoIncluido = Aux
        
    End Function
	
	

	
%>
<!-- saved from url=(0022)http://internet.e-mail -->
<!--<OBJECT RUNAT=server PROGID=ADODB.Recordset id=malla name=malla></OBJECT> -->
<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<script language="JavaScript">

    var total = 0;
    var total_cuotas = " ";

    function inicio() {
    }

    function suma(cuanto) {
        total = total + cuanto;
    }

    function resta(cuanto) {
        total = total - cuanto;
    }

    function totalcero() {
        total = 0;
        total_cuotas = " ";
    }

    function determina(cual, cuanto, cuotas) {
        if (document.getElementById(cual).checked == true) {
            suma(cuanto);
        } else {
            resta(cuanto);
        }
        document.getElementById('TBK_MONTO').value = total;
        document.getElementById('cuota1').value = total;
    }

    function cantidad_cuotas(cuotas) {
        total_cuotas = total_cuotas + "/" + cuotas;
    }

    function resta_cantidad_cuotas(cuotas) {
        total_cuotas = " ";
    }

    function sumaresta(cual, cuanto, regtope, cuotas, id_cuota, Idpagoweb) {
		document.getElementById("divProceso").style.visibility = "visible";
		
        if (document.getElementById(cual).checked == true) {
            suma(cuanto);

        } else {
            resta(cuanto);
            if (total < 0) {
                for (var ipos = 0; ipos < regtope; ipos++) {
                    document.getElementById('id_apagar' + ipos).checked = false
                    resta_cantidad_cuotas(cuotas);
                }
                totalcero();
            }
        }
		
        cantidad_cuotas(cuotas);
        document.getElementById('TBK_MONTO').value = total;
        document.getElementById('cuota1').value = total;
        document.getElementById('cuo').value = total_cuotas;
        //cantidad_cuotas(cuotas);
        //alert(total_cuotas);
        document.location = "PagoWeb_MarcaCuota.asp?IDPAGOWEB=" + Idpagoweb + "&ID_CUOTA=" + id_cuota
		//document.getElementById("Img2").setAttribute('disabled','true');

    }

    function cuenta_cuotas(cual, cuanto, regtope, cuotas) {
        if (document.getElementById('cuota1').value > 0) {
            for (var ipos = 0; ipos < regtope; ipos++) {
                document.getElementById('id_apagar' + ipos).checked = true
                cantidad_cuotas(cuotas);
            }
            document.getElementById('cuo').value = total_cuotas;
            //alert(total_cuotas);
            //form1.submit();
            enviar();
        } else {
            alert('Debe seleccionar cuota a pagar');
        }
    }

	function cuenta_cuotasTBK(cual, cuanto, regtope, cuotas) {
        if (document.getElementById('cuota1').value > 0) {
            for (var ipos = 0; ipos < regtope; ipos++) {
                document.getElementById('id_apagar' + ipos).checked = true
                cantidad_cuotas(cuotas);
            }
            document.getElementById('cuo').value = total_cuotas;
            //alert(total_cuotas);
            //form1.submit();
            enviarTBK();
        } else {
            alert('Debe seleccionar cuota a pagar');
        }
    }
	
    function vapago() {
        if (document.getElementById('cuota1').value > 0) {
            form1.submit();
        } else {
            alert('Debe seleccionar cuota a pagar');
        }

    }

    function enviar() {

        var elementos = document.getElementsByName("apagar");
 
		  for (x=0;x<elementos.length;x++)
		    elementos[x].disabled= true;
			
        var props = "fullscreen,resizable=no,menubar=no,location=no,toolbar=no,status=no,scrollbars=no,directories=no";
        

        window.open('BP.asp', 'Confirmación', props);

    }
	
	function enviarTBK() {  

        var elementos = document.getElementsByName("apagar");
 
		  for (x=0;x<elementos.length;x++)
		    elementos[x].disabled= true;
			  
        var props = "fullscreen,resizable=no,menubar=no,location=no,toolbar=no,status=no,scrollbars=no,directories=no";         

        window.open('BPTBK.asp', 'Confirmación', props);

    }
    
</script>
<!--<OBJECT RUNAT=server PROGID=ADODB.Recordset id=Detalle name=Detalle > </OBJECT>-->
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
strsql= "select a.rut, a.codcarpr, c.nombre_l, a.ano from mt_alumno a, mt_carrer c where a.codcli = '" & codcli & "' and a.codcarpr = c.codcarr"
Set Rst = Session("Conn").Execute(StrSql)
if not rst.eof then
	rutcli =rst("rut")
	codcarr =rst("codcarpr")
	nombrecarrera =rst("nombre_l")	
	anoing =rst("ano")		
end if

' CONSULTA DATOS DE ALUMNO PARA ENCABEZADO
strsql = "select codcli,dig,nombre,paterno,materno,diractual,comuna,ciudadact,fonoact,codaval,fecnac,codsissalud, Coalesce(codaval, '') as codaval, coalesce(codapod, '') as codapod from mt_client  where codcli = '" & rutcli & "'"
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


    if trim(apoderado) = "" then 
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
    	
	    ' TRAE DATOS DEL APODERADO
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
strsql="SP_CONSULTA_CCORRIENTE_PA '" & rutcli & "','" & codcarr & "',null,'SI'" 
dim HayRegistros 


Set Rst = Session("Conn").Execute(StrSql)  


if not rst.eof  then 
  CuentasPay = rst.getrows  
  maxfilas = Ubound(CuentasPay, 2) 
  HayRegistros = 1 
else
    HayRegistros = 0
end if 

%>
<html>
<head>
<title><%=Session("NombrePestana")%></title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html;">
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
                                                    <img src="Imagenes/titulos/T-Pago-Cuotas.gif" width="271" height="38" /></p>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td width="100%">
                                                <tr>
                                                    <td colspan="2" height="30">
                                                        <p>
                                                            <!-- <span class="" width="834"><font face="Arial, Helvetica, sans-serif">Datos del Alumno</font></span> -->
                                                            <span style="font-size: 13px; font-weight: bold; font-family: Verdana, Arial, Helvetica, sans-serif;">
                                                                <b class="Tit-celdas"><font face="Verdana, Arial, Helvetica, sans-serif" id="lblDatosAlumno">Datos del alumno</font></b></span>
                                                        </p>
                                                        <table width="834" border="1">
                                                            <tr>
                                                                <th background="imagenes/fdo-cabecera-cel22.jpg" width="70" bgcolor="4A5DA1" scope="col">
                                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif" id="lblRut" name="lblRut">
                                                                        Rut</font></strong>
                                                                </th>
                                                                <!--<th background="imagenes/fdo-cabecera-cel22.jpg" width="20" bgcolor="4A5DA1" scope="col">
                                                                            <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                                DV</font></strong>
                                                                        </th>-->
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
                                                                            <%if session("LOCALIZACION") ="PERU" then
																				response.Write(alumrut)
																			else
																				response.Write(FormatNumber(alumrut, 0) + "-" + UCase(alumdv))
																			end if%>
                                                                            </font></div>
                                                                </td>
                                                                <!--<td width="20" height="32" >
                                                                        <div align="center">
                                                                            <font face="Arial, Helvetica, sans-serif" size="1"> <%=alumdv%></font></div>
                                                                    </td>-->
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
                                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif" id="lblRut" name="lblRut">
                                                                        Rut</font></strong>
                                                                </th>
                                                                <!--<th background="imagenes/fdo-cabecera-cel22.jpg" width="20" bgcolor="4A5DA1" scope="col">
                                                                            <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                                DV</font></strong>
                                                                        </th>-->
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
                                                                            <%if session("LOCALIZACION") ="PERU" then
																				response.Write(apodrut)
																			else
																				response.Write(FormatNumber(apodrut, 0) + "-" + UCase(apoddv))
																			end if%>
                                                                            </font></div>
                                                                </td>
                                                                <!--<td width="20" height="32" >
                                                                        <div align="center">
                                                                            <font face="Arial, Helvetica, sans-serif" size="1"><%=apoddv%></font></div>
                                                                    </td>-->
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
                                                                    <!-- <font face="Arial, Helvetica, sans-serif" class="Tx_Gral3_azul">Estas son las cuotas
                                                                            que puedes pagar por Internet.</font>-->
                                                                    <span style="font-size: 13px; font-weight: bold; font-family: Verdana, Arial, Helvetica, sans-serif;">
                                                                        <b class="Tit-celdas"><font face="Verdana, Arial, Helvetica, sans-serif">Estas son las
                                                                            cuotas que puedes pagar por internet
                                                                            <br>
                                                                            <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#666666">(Este documento
                                                                                NO constituye certificado)</font> <font face="Verdana, Arial, Helvetica, sans-serif"
                                                                                    color="#FFFFFF"><b><font size="1">
                                                                                        <br />
                                                                </p>
                                                                <p>
                                                                    <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#666666">Seleccione
                                                                        las cuotas a pagar</font>
                                                                    <br>
                                                                    <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#666666">(Se debe
                                                                        seleccionar las cuotas en forma correlativa, partiendo por el primer vencimiento)</font>
                                                                    <font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">
                                                                        <br />
                                                                </p>
                                                            </td>
                                                        </tr>
                                                        <tr>
                                                            <td>
                                                                <table border="0" cellpadding="0" cellspacing="0" height="246" align="left" width="899">
                                                                    <tr valign="top">
                                                                        <td width="899" height="70" colspan="1">
                                                                            <table width="900" cellspacing="0" cellpadding="0" height="72" border="1" bordercolor="#FFFFFF">
                                                                                <tr background="imagenes/fdo-cabecera-cel22.jpg">
                                                                                    <td background="imagenes/fdo-cabecera-cel22.jpg" width="30" height="30">
                                                                                        <div align="center">
                                                                                            <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                                                                Pagar</font></strong></div>
                                                                                    </td>
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
                                                                                </tr>
                                                                                <form name="form1" method="POST" action="https://siga.idma.cl/Alumnosnet/pago_servipag.asp?codcli=<%=RUT%>">
                                                                                <%
			pos = 0 
			if maxfilas >= 0 and HayRegistros = 1 then 
				do while pos <= maxfilas
				montoacancelarcuota = 0
				
				'mejora antes era solo montoacancelar = montoacancelar  + ccur(CuentasPay(19, pos))
				montoacancelar = montoacancelar  + (ccur(CuentasPay(19, pos)) - ccur(CuentasPay(25, pos))) 
				montoacancelarcuota = montoacancelarcuota  + ( ccur(CuentasPay(19, pos)) - ccur(CuentasPay(25, pos)) )  + ccur(CuentasPay(33, pos)) + ccur(CuentasPay(32, pos))
				
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
				
                CuentasPay(15, pos)=CuentasPay(35, pos)

				'strFecha="SELECT GETDATE()fecha"
				'Set RstFecha = Session("Conn").Execute(strFecha)
				'fechahoy=RstFecha("fecha")
				'
				'strdia="SELECT dbo.Fn_ValorParame('diasatr')dia "
				'Set Rstdia = Session("Conn").Execute(strdia)
				'diaAtraso=valnulo(Rstdia("dia"),NUM_)
				
				''Logica para morosos y vencidos				
				'If DateDiff("d",CuentasPay(5, pos) , fechahoy) < 0 Then	
				'	If IsDate(CuentasPay(34, pos)) Then
				'		If CuentasPay(5, pos) > CuentasPay(34, pos) Then
				'			CuentasPay(15, pos)="Prorrogado-VIGENTE"
				'		Else
				'			CuentasPay(15, pos)="VIGENTE"
				'		End If
				'	Else
				'		CuentasPay(15, pos)="VIGENTE"
				'	End If
				'ElseIf DateDiff("d",CuentasPay(5, pos) , fechahoy)>= 0 Then
				'	If DateDiff("d",CuentasPay(5, pos) , fechahoy) > diaAtraso Then
				'		If CuentasPay(5, pos) > CuentasPay(34, pos) Then
				'			CuentasPay(15, pos)="Prorrogado-MOROSO"
				'		Else
				'			CuentasPay(15, pos)="MOROSO"
				'		End If
				'	Else
				'		if CuentasPay(5, pos) > CuentasPay(34, pos) Then
				'			CuentasPay(15, pos)="Prorrogado-VENCIDO"
				'		Else
				'			CuentasPay(15, pos)="VENCIDO"
				'		End If
				'	End If
				'End IF 

                                                                                %>
                                                                                <tr bgcolor="#DBECF2">
                                                                                    <td width="30" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Arial, Helvetica, sans-serif" size="1">
                                                                                                <label>
                                                                                                    <input type="checkbox" onClick="javascript:sumaresta('id_apagar<%=pos%>',<%=Int(FormatNumber((montoacancelarcuota), 0))%>,<%=maxfilas%>,<%=CuentasPay(1, pos)%>,<%=CuentasPay(31,pos)%>,<%=IdPagoWeb%>)"
                                                                                                        name="apagar" id="id_apagar<%=pos%>" value="ON" <%=PagoIncluido(IDPAGOWEB,CuentasPay(31,pos))%>>
                                                                                                </label>
                                                                                            </font>
                                                                                        </div>
                                                                                    </td>
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
                                                                                                <%=FormatNumber(montoacancelarcuota, CantDecimales)%></font></div>
                                                                                    </td>
                                                                                    <td width="80" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Arial, Helvetica, sans-serif" size="1">
                                                                                                <%=CuentasPay(1, pos)%></font></div>
                                                                                    </td>
                                                                                    <td width="60" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Arial, Helvetica, sans-serif" size="1">
                                                                                                <%=FormatNumber(CuentasPay(7, pos), CantDecimales)%></font></div>
                                                                                    </td>
                                                                                    <td width="60" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Arial, Helvetica, sans-serif" size="1">
                                                                                                <%=FormatNumber(intrec, CantDecimales)%></font></div>
                                                                                    </td>
                                                                                    <td width="60" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Arial, Helvetica, sans-serif" size="1">
                                                                                                <%=FormatNumber(CuentasPay(33, pos), CantDecimales)%></font></div>
                                                                                    </td>
                                                                                    <td width="60" height="30">
                                                                                        <div align="center">
                                                                                            <font face="Arial, Helvetica, sans-serif" size="1">
                                                                                                <%=FormatNumber(CuentasPay(25, pos), CantDecimales)%></font></div>
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
                                                                            <table width="900" border="1">
                                                                                <tr>
                                                                                    <th background="imagenes/fdo-cabecera-cel22.jpg" width="278" bgcolor="4a5da1" scope="col">
                                                                                        <div align="right">
                                                                                            <span class="style2"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"
                                                                                                size="1">Monto Total</font> </span>
                                                                                        </div>
                                                                                    </th>
                                                                                    <td width="110" height="32" bgcolor="#DBECF2">
                                                                                        <div align="center">
                                                                                            <font face="Arial, Helvetica, sans-serif" size="2">
                                                                                                <%=FormatNumber(montoacancelar, CantDecimales)%></font></div>
                                                                                    </td>
                                                                                    <th background="imagenes/fdo-cabecera-cel22.jpg" width="151" bgcolor="4a5da1" scope="col">
                                                                                        <div align="right" class="style2">
                                                                                            <font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF" size="1">Monto a
                                                                                                cancelar</font>
                                                                                        </div>
                                                                                    </th>
                                                                                    <td width="111" bgcolor="#DBECF2">
                                                                                        <input type="text" id="TBK_MONTO" name="TBK_MONTO" size="10" style="background-color: #DBECF2;
                                                                                            color: #000000; border: none; font-family: Arial, Helvetica, sans-serif; font-size: 10pt;
                                                                                            letter-spacing: 1px; text-align: center" value="<%=FormatNumber(monto_a_pagar, CantDecimales)%>" />
                                                                                        <input type="HIDDEN" id="FILAS" name="FILAS" value="0" />
                                                                                        <!--<%=totalfilas%>"> -->
                                                                                        <input type="HIDDEN" id="TBK_ORDEN_COMPRA" name="TBK_ORDEN_COMPRA" value="<%=OC%>" />
                                                                                        <input type="HIDDEN" id="TBK_ID_SESION" name="TBK_ID_SESION" value="<%=OC%>" />
                                                                                        <input type="HIDDEN" id="CUOTA_PAGO" name="CUOTA_PAGO" value="<%=CUOTA_PAGO%>" />
                                                                                        <input type="HIDDEN" id="RUT_PAGO" name="RUT_PAGO" value="<%=RUT%>" />
                                                                                        <input type="HIDDEN" id="pagar1" name="pagar1" value="PAGAR" />
                                                                                        <input type="HIDDEN" id="cuota1" name="cuota1" value="<%=CuentaCuotas%>" />
                                                                                        <input type="HIDDEN" id="cuo" name="cuo" value="<%=cuo%>" />
                                                                                        <% cuo=session("cuo") %>
                                                                                    </td>
                                                                                </tr>
                                                                            </table>
                                                                        </td>
                                                                    </tr>
                                                                    <tr>
                                                                    <td>
                                                                    
                                                                    <table width="100%">
                                                                    <tr>
    <td width="50%" align="left" ><A onMouseOver="MM_swapImage('Image7','','Imagenes/botones/volver-on.gif',1)" onmouseout=MM_swapImgRestore() href="javascript: window.top.location.href ='menu_finanzas.asp'"><IMG src="Imagenes/botones/volver-of.gif" border="0"  name="Image7"></A></td>
 <%
strParame="SELECT dbo.Fn_ValorParame('ACTIVABPSERVIPAG')SERVIPAG,dbo.Fn_ValorParame('ACTIVABPTRANSBANK')TRANSBANK"
if BCL_ADO(strParame, rstParame) then
	SERVIPAG = rstParame("SERVIPAG")
	TRANSBANK = rstParame("TRANSBANK")
end if
 %>
 
 	<%if TRANSBANK="SI" then%>
        <td align="right" > 
            <a onClick="cuenta_cuotasTBK()" href="menu_finanzas.asp"><img src="img/boton-webpay.gif" alt="Pagar en servipag" name="boton" width="200" height="70"
            border="0" id="Img2" />
            </a>
        </td>
    <%end if %>
    
    <%if SERVIPAG="SI" then%>
        <td align="right" >
            <a onClick="cuenta_cuotas()" href="menu_finanzas.asp"><img src="img/boton-sp.gif" alt="Pagar en servipag" name="boton" width="200" height="70"
            border="0" id="Img2" />
            </a>
        </td>
    <%end if %>

                                                                        </tr>
                                                                        </table>
                                                                        
                                                                        </td>
                                                                    </tr>
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
    </table>
    <!-- fin codigo sin formato U++ -->
    </form>
    
<div id="divProceso" class="cCargando" style="visibility: hidden;">
        <div id="divProcesoMsg" class="cCargandoImg">
            <br />
            En proceso... 
            <br />
        </div>
    </div>

<style>

.cCargando {
    width: 100%;
    height: 100%;
    top: 0;
    bottom: 0;
    left: 0;
    right: 0;
    margin: auto;
    position: fixed;
    background-color: #000;
    opacity: 0.8;
    filter: alpha(opacity=80); /* Internet Explorer 8*/
    z-index: 9999;
    transition: width 2s;
    -moz-transition: width 2s; /* Firefox 4 */
    -webkit-transition: width 2s; /* Safari and Chrome */
    -o-transition: width 2s; /* Opera */
    cursor: progress;
}

.cCargandoImg {
    cursor: progress;
    position: absolute;
    top: 32%;
    right: 45%;
    left: 35%;
    filter: alpha(opacity=80); /* Internet Explorer 8*/
    opacity: 0.8;
    margin: auto;
    width: 350px;
    text-align: center;
    height: 150px;
    padding: 10px;
    background-color: #000;
    /*border: 1px solid #000;*/
    color: #ffffff;
    font-size: 2em;

}
</style>

</body>
<%ObjetosLocalizacion("pago-web-cuotas.asp")%>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
