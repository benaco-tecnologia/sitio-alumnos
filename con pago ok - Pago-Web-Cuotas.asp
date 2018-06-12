<%  response.buffer = false 
    Response.Expires = -1 
    
    var_servidor = "200.54.71.237" 
    
    Dim IdPagoWeb 
    
    IdPagoWeb = Request("IDPAGOWEB")
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
    Set miRst = Session("Conn").Execute(miSql)
    if not miRst.Eof then
        FechaVen = miRst(0)
    end if
    miRst.Close
    Set miRst = Nothing
        
    miSql = "Select Coalesce(count(*), 0) from matricula.mt_pago_web_cuotas Where IdPagoWeb = " & IdPagoWeb
    Set miRst = Session("Conn").Execute(miSql)
    if not miRst.Eof then 
        CuentaCuotas = miRst(0)
    End if
    miRst.Close
    Set miRst = Nothing
    
    miSql = "SELECT Convert(varchar, Convert(int, Round(Coalesce(matricula.fn_PagoWeb_Obtiene_monto_a_Pagar(" & IdPagoWeb & "), 0), 0) ) )"
    Set miRst = Session("Conn").Execute(miSql)
    
    if not miRst.Eof then 
        monto_a_pagar = miRst(0)
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
    Sql = "<?xml version='1.0' encoding='ISO8859-1'?><Servipag><Header><FirmaEPS>" & Firma & "</FirmaEPS><CodigoCanalPago>640" 
    Sql = Sql +        "</CodigoCanalPago><IdTxCliente>" & IDPAGOWEB & "</IdTxCliente><FechaPago>" & FechaPago & "</FechaPago><MontoTotalDeuda>" & monto_a_pagar 
    Sql = Sql +    "</MontoTotalDeuda><NumeroBoletas>1</NumeroBoletas></Header><Documentos><IdSubTrx>1</IdSubTrx><CodigoIdentificador>" & IDPAGOWEB 
    Sql = Sql +        "</CodigoIdentificador><Boleta>1</Boleta><Monto>" & monto_a_pagar 
    Sql = Sql +        "</Monto><FechaVencimiento>" & FechaVen & "</FechaVencimiento></Documentos></Servipag>"

    Session("miXML") = Sql
    
   Function PagoIncluido(idpagoweb, id_cuota)
        Dim Sql, Rst, Aux
        
        Sql = "Select 1 From matricula.mt_pago_web_cuotas Where IdPagoWeb = " & idpagoweb & " and id_cuota = " & id_cuota
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

    function vapago() {
        if (document.getElementById('cuota1').value > 0) {
            form1.submit();
        } else {
            alert('Debe seleccionar cuota a pagar');
        }

    }

    function enviar() {

        //var xml, xml1, xml2, xml3, xml4
        var props = "height=425,width=675,top=100,left=100,resizable=no,menubar=no,location=no,toolbar=no,status=no,scrollbars=no,directories=no";
        //xml = "<?xml version='1.0' encoding='ISO8859-1'?><Servipag><Header><FirmaEPS>" + '<%=Firma%>' + "</FirmaEPS><CodigoCanalPago>640";
        //xml1 = "</CodigoCanalPago><IdTxCliente>" + '<%=IDPAGOWEB%>' + "</IdTxCliente><FechaPago>" + '<%=FechaPago%>' + "</FechaPago><MontoTotalDeuda>" + '<%=monto_a_pagar%>';
        //xml2 = "</MontoTotalDeuda><NumeroBoletas>1</NumeroBoletas></Header><Documentos><IdSubTrx>1</IdSubTrx><CodigoIdentificador>" + '<%=IDPAGOWEB%>';
        //xml3 = "</CodigoIdentificador><Boleta>1</Boleta><Monto>" + '<%=monto_a_pagar%>';
        //xml4 = "</Monto><FechaVencimiento>" + '<%=FechaVen%>' + "</FechaVencimiento></Documentos></Servipag>";

        //var todo
        //todo = xml + xml1 + xml2 + xml3 + xml4;
        //document.datos.mixml.value = todo;

        //var mirul

        window.open('http://leo/alumnos/BP.asp', 'Confirmación', props);

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

codcli=session("codcli")


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
strsql= "select a.rut, a.codcarpr, c.nombre_c, a.ano from matricula.mt_alumno a, matricula.mt_carrer c where a.codcli = '"+ codcli + "' and a.codcarpr = c.codcarr"
Set Rst = Session("Conn").Execute(StrSql)
if not rst.eof then
	rutcli =rst("rut")
	codcarr =rst("codcarpr")
	nombrecarrera =rst("nombre_c")	
	anoing =rst("ano")		
end if

' CONSULTA DATOS DE ALUMNO PARA ENCABEZADO
strsql = "select codcli,dig,nombre,paterno,materno,diractual,comuna,ciudadact,fonoact,codaval,fecnac,codsissalud, codaval from matricula.mt_client  where codcli = " + rutcli 
	Set Rst=Session("Conn").Execute(StrSql)
	if not rst.eof then
		alumrut =rst("codcli")
		RUT =rst("codcli")
		alumdv =rst("dig")
		alumnombre =rst("nombre")
		alumpaterno =rst("paterno")
		alummaterno =rst("materno")
		apoderado = rst("codaval")
	end if

	' TRAE DATOS DEL APODERADO
	strsql= "select * from matricula.mt_apoder where codapod = '"+ apoderado + "'"
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


' CONSULTA DATOS DE ALUMNO PARA CTA.CTE.
strsql= "select  c.ctadoc, c.ctadocnum, b.nombre, c.codban, c.codsuc, c.fecven, gastosop, isnull(c.saldo,0) saldo, c.estado, i.descripcion as itemdesc, por1, por2, por3, "
strsql= strsql + " por4, c.estado, est.descripcion as estadodesc, c.feccancel, c.codapod, c.IntDeuda, c.monto, i.coditem, c.num_operacion, c.TipoPagCond, i.activa,"
strsql= strsql + " c.item, (c.monto - c.saldo) pago, (c.monto + c.intdeuda) montoacancelar, d.nombre as ubicacion, c.gastospro, c.intrepac, datediff(d, c.fecven,'" 
strsql= strsql + cstr(feccorte) + "') fechadiff, c.ID_CUOTA "
strsql = strsql + " ,CASE WHEN p.Interesdiario = 'N' THEN matricula.InteresMora(c.Saldo, c.FecVen, GETDATE(), -1) WHEN p.Interesdiario = 'S' THEN matricula.Interespordia(c.Saldo, c.FecVen, CONVERT(VARCHAR(8), GETDATE(), 112), -1 )	 ELSE 0 	END AS mis_interes "
strsql= strsql + " from matricula.mt_ctadoc c ,matricula.mt_item i, matricula.mt_estdoc est, matricula.mt_docum b, matricula.mt_detubi d, mt_parame p"
strsql= strsql + " where codcli = '" + rutcli + "' and saldo > 0 "
strsql= strsql + " and c.codcarr = '" + codcarr + "'"
strsql= strsql + " and c.item*=i.coditem"
strsql= strsql + " and c.estado=est.estado"
strsql= strsql + " and c.ctadoc  = b.tipodoc"
StrSql= StrSql + " and c.ubicacion = d.codubifis    "
strsql= strsql + " order by fecreg desc ,CTADOC,fecven, ctadocnum "


'response.Write(strsql)
'response.End

Set Rst = Session("Conn").Execute(StrSql)  

if not rst.eof  then 
  CuentasPay = rst.getrows  
  maxfilas = Ubound(CuentasPay, 2) 
end if 
%>
<html>
<head>
    <title>.:: PAGO DE CUOTAS POR INTERNET ::.</title>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
    <!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
    <!--<link rel="stylesheet" href="file://///Gotzilla/uddweb/admalumnos/css/tex-normales.css"      type="text/css">-->
    <link href="css/tex-normales.css" rel="stylesheet" type="text/css"/>
    <!-- <link rel="stylesheet" href="css/tablas.css" type="text/css"/> -->
    <link href="estilos_mas.css" rel="stylesheet" type="text/css">
    
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
    <form name="datos" id="datos" method="post" action="browse.asp">

<table border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right"><%CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td valign="top" >
			<% CargarTop1()%><% SubMenu()%>
			
    
    <!-- codigo sin formato U++ -->
    <!--<div align="left">-->
        <table>
            <tr>
                <table width="834" border="0">
                    <tr>
                        <td>
                            <div align="center">
                                <span class="Tx_Gral3_verde">Servipag - Transacci&oacute;n NORMAL (pesos)</span>
                                <!--<img src="http://www.vallevirtual.cl/webpay/img/webpay.gif" width="120" height="60">-->
                                <br />
                                <span class="Tx_Gral3_verde">Paso 1 (Selecci&oacute;n de cuotas a cancelar) </span>
                            </div>
                        </td>
                    </tr>
                </table>
            </tr>
            <tr valign="bottom">
                <td height="114">
                    <td colspan="2" height="30" class="tex-normales">
                        <p>
                            <span class="" width="834"><font face="Arial, Helvetica, sans-serif" >Datos del Alumno</font></span></p>
                        <table width="834" border="1">
                            <tr>
                                <th width="70" bgcolor="4A5DA1" scope="col">
                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                        Rut</font></strong>
                                </th>
                                <th width="20" bgcolor="4A5DA1" scope="col">
                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                        DV</font></strong>
                                </th>
                                <th width="130" bgcolor="4A5DA1" scope="col">
                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                        Nombres</font></strong>
                                </th>
                                <th width="130" bgcolor="4A5DA1" scope="col">
                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                        Apellido Paterno</font></strong>
                                </th>
                                <th width="130" bgcolor="4A5DA1" scope="col">
                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                        Apellido Materno</font></strong>
                                </th>
                                <th width="90" bgcolor="4A5DA1" scope="col">
                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                        C&oacute;digo Carrera</font></strong>
                                </th>
                                <th width="130" bgcolor="4A5DA1" scope="col">
                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                        Nombre Carrera</font></strong>
                                </th>
                                <th width="82" bgcolor="4A5DA1" scope="col">
                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                        A&ntilde;o Ingreso</font></strong>
                                </th>
                            </tr>
                            <tr>
                                <td width="70" height="32" bgcolor="ffc172">
                                    <div align="center">
                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                            <%=alumrut%></font></div>
                                </td>
                                <td width="20" height="32" bgcolor="ffc172">
                                    <div align="center">
                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                            <%=alumdv%></font></div>
                                </td>
                                <td width="130" height="32" bgcolor="ffc172">
                                    <div align="center">
                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                            <%=alumnombre%></font></div>
                                </td>
                                <td width="130" height="32" bgcolor="ffc172">
                                    <div align="center">
                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                            <%=alumpaterno%></font></div>
                                </td>
                                <td width="130" height="32" bgcolor="ffc172">
                                    <div align="center">
                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                            <%=alummaterno%></font></div>
                                </td>
                                <td width="90" height="32" bgcolor="ffc172">
                                    <div align="center">
                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                            <%=codcarr%></font></div>
                                </td>
                                <td width="130" height="32" bgcolor="ffc172">
                                    <div align="center">
                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                            <%=nombrecarrera%></font></div>
                                </td>
                                <td width="82" height="32" bgcolor="ffc172">
                                    <div align="center">
                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                            <%=anoing%></font></div>
                                </td>
                            </tr>
                        </table>
                        <p>
                            <span class="style5"><font face="Arial, Helvetica, sans-serif" class="Tx_Gral3_azul">
                                Antecedentes Responsable Financiero</font></span></p>
                        <table width="610" border="1">
                            <tr>
                                <th width="70" bgcolor="4A5DA1" scope="col">
                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                        Rut</font></strong>
                                </th>
                                <th width="20" bgcolor="4A5DA1" scope="col">
                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                        DV</font></strong>
                                </th>
                                <th width="130" bgcolor="4A5DA1" scope="col">
                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                        Nombres</font></strong>
                                </th>
                                <th width="130" bgcolor="4A5DA1" scope="col">
                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                        Apellido Paterno</font></strong>
                                </th>
                                <th width="130" bgcolor="4A5DA1" scope="col">
                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                        Apellido Materno</font></strong>
                                </th>
                                <th width="90" bgcolor="4A5DA1" scope="col">
                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                        Tel&eacute;fono</font></strong>
                                </th>
                            </tr>
                            <tr>
                                <td width="70" height="32" bgcolor="ffc172">
                                    <div align="center">
                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                            <%=apodrut%></font></div>
                                </td>
                                <td width="20" height="32" bgcolor="ffc172">
                                    <div align="center">
                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                            <%=apoddv%></font></div>
                                </td>
                                <td width="130" height="32" bgcolor="ffc172">
                                    <div align="center">
                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                            <%=apodnombres%></font></div>
                                </td>
                                <td width="130" height="32" bgcolor="ffc172">
                                    <div align="center">
                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                            <%=apodpaterno%></font></div>
                                </td>
                                <td width="130" height="32" bgcolor="ffc172">
                                    <div align="center">
                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                            <%=apodmaterno%></font></div>
                                </td>
                                <td width="90" height="32" bgcolor="ffc172">
                                    <div align="center">
                                        <font face="Arial, Helvetica, sans-serif" size="1">
                                            <%=apodtelefono%></font></div>
                                </td>
                            </tr>
                        </table>
                        <tr>
                            <td>
                                <p class="Apuntes">
                                    <font face="Arial, Helvetica, sans-serif" class="Tx_Gral3_azul">Estas son las cuotas
                                        que puedes pagar por Internet.</font><br>
                                    <font size="1" face="Verdana, Arial, Helvetica, sans-serif" color="#666666">(Este documento
                                        NO constituye certificado)</font> <font face="Verdana, Arial, Helvetica, sans-serif"
                                            color="#FFFFFF"><b><font size="1">
                                                <br>
                                                <!--<input type="button" name="pagar" id="bpagar" value="PAGAR CUOTAS" onclick="javascript:cuenta_cuotas()"  >-->
                                                <br />
                                                <a onclick="cuenta_cuotas()" href="#">
                                                    <!--onmouseout="MM_swapImgRestore()" onmouseover="MM_swapImage('boton-sp.gif','','img/boton-sp.gif',1)"-->
                                                    <img src="img/boton-sp.gif" alt="Pagar en servipag" name="boton" width="200" height="70"
                                                        border="0" id="boton" />
                                                </a></font></b></font>
                                </p>
                                <p class="Apuntes">
                                    <span class="Calend_DiaFeriado">Seleccione las cuotas a cancelar</span></p>
                            </td>
                        </tr>
                        <table border="0" cellpadding="0" cellspacing="0" height="246" align="left" width="899">
                            <tr valign="top">
                                <td width="899" height="70" colspan="1">
                                    <table width="900" cellspacing="0" cellpadding="0" height="72" border="1" bordercolor="#FFFFFF">
                                        <tr bgcolor="4a5da1">
                                            <td width="30" height="30" bgcolor="4a5da1">
                                                <div align="center">
                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                        Pagar</font></strong></div>
                                            </td>
                                            <td width="90" height="30" bgcolor="4a5da1">
                                                <div align="center">
                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                        Vencimiento</font></strong></div>
                                            </td>
                                            <td width="60" height="30">
                                                <div align="center">
                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                        Documento</font></strong></div>
                                            </td>
                                            <td width="80" height="30">
                                                <div align="center">
                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                        Monto a Cancelar</font></strong>
                                                </div>
                                            </td>
                                            <td width="60" height="30">
                                                <div align="center">
                                                    <font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">
                                                        Correlativo</font></b></font>
                                                </div>
                                            </td>
                                            <td width="60" height="30">
                                                <div align="center">
                                                    <strong><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif">
                                                        Monto Capital</font></strong></div>
                                            </td>
                                            <td width="60" height="30">
                                                <div align="center">
                                                    <font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">
                                                        Int.Mora</font></b></font></div>
                                            </td>
                                            <td width="80" height="30">
                                                <div align="center">
                                                    <font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">
                                                        Pagos</font></b></font></div>
                                            </td>
                                            <td width="90" height="30">
                                                <div align="center">
                                                    <font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">
                                                        Estado</font></b></font></div>
                                            </td>
                                            <td height="30" width="90">
                                                <div align="center">
                                                    <font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">
                                                        Ub icaci&oacute;n</font></b></font></div>
                                            </td>
                                        </tr>
                                        <form name="form1" method="POST" action="https://200.71.198.46/Alumnos/pago_servipag.asp?codcli=<%=RUT%>">
                                        <%
			pos = 0 
			if maxfilas > 0 then 
				do while pos <= maxfilas
				montoacancelarcuota = 0
				montoacancelar = montoacancelar  + ccur(CuentasPay(19, pos)) 
				montoacancelarcuota = montoacancelarcuota  + ccur(CuentasPay(19, pos)) 
				
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
					recargos = CuentasPay(32, pos) '0 'InteresMora(CuentasPay(19, pos), CuentasPay(5, pos), feccorte, diasdif)

					'intrec = intrec + recargos
					intrec = cdbl(recargos)
				end if				
				
			  	intrectotal = intrectotal + intrec 
            	monto = monto + ccur(CuentasPay(19, pos))
				montototal = montototal + monto
            	saldo = saldo + ccur(CuentasPay(7, pos))
				saldototal = saldototal + saldo
				
				montoacancelarcuota = montoacancelarcuota  + intrec
				montoacancelar = montoacancelar  + intrec
				' REVISA CARGO DE INTERESES

                                        %>
                                        <tr bgcolor="ffc172">
                                            <td width="30" height="30">
                                                <div align="center">
                                                    <font face="Arial, Helvetica, sans-serif" size="1">
                                                        <label>
                                                            <input type="checkbox" onclick="javascript:sumaresta('id_apagar<%=pos%>',<%=montoacancelarcuota%>,<%=maxfilas%>,<%=CuentasPay(1, pos)%>,<%=CuentasPay(31,pos)%>,<%=IdPagoWeb%>)"
                                                                name="apagar(<%=pos%>)" id="id_apagar<%=pos%>" value="ON" <%=PagoIncluido(IDPAGOWEB,CuentasPay(31,pos))%>>
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
                                                        <%=montoacancelarcuota%></font></div>
                                            </td>
                                            <td width="80" height="30">
                                                <div align="center">
                                                    <font face="Arial, Helvetica, sans-serif" size="1">
                                                        <%=CuentasPay(1, pos)%></font></div>
                                            </td>
                                            <td width="60" height="30">
                                                <div align="center">
                                                    <font face="Arial, Helvetica, sans-serif" size="1">
                                                        <%=CuentasPay(7, pos)%></font></div>
                                            </td>
                                            <td width="60" height="30">
                                                <div align="center">
                                                    <font face="Arial, Helvetica, sans-serif" size="1">
                                                        <%=intrec%></font></div>
                                            </td>
                                            <td width="60" height="30">
                                                <div align="center">
                                                    <font face="Arial, Helvetica, sans-serif" size="1">
                                                        <%=CuentasPay(25, pos)%></font></div>
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
                                            <th width="278" bgcolor="4a5da1" scope="col">
                                                <div align="right">
                                                    <span class="style2">Monto Total</span></div>
                                            </th>
                                            <td width="110" height="32" bgcolor="ffc172">
                                                <div align="center">
                                                    <font face="Arial, Helvetica, sans-serif" size="2">
                                                        <%=montoacancelar%></font></div>
                                            </td>
                                            <th width="151" bgcolor="4a5da1" scope="col">
                                                <div align="right" class="style2">
                                                    Monto a Cancelar</div>
                                            </th>
                                            <td width="111" bgcolor="ffc172">
                                                <input type="text" id="TBK_MONTO" name="TBK_MONTO" size="10" style="background-color: ffc172;
                                                    color: #000000; border: none; font-family: Arial, Helvetica, sans-serif; font-size: 10pt;
                                                    letter-spacing=1px; text-align: center" value="<%=monto_a_pagar%>" />
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
                        </table>
            </tr>
        </table>
        <!--<textarea id="mixml"></textarea>-->
     <!--</div>-->
     <!-- fin codigo sin formato U++ -->
 
     </td>
		  </tr>
		</table>
	</td>
  </tr>
</table>
     
    </form>
</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
