<%  response.buffer = false 
    Response.Expires = -1
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
<!--#INCLUDE file="analytics.asp" -->

<%
strParame="SELECT coalesce(dbo.Fn_ValorParame('USAANALYTICS'),'NO')Parame"
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
	USAANALYTICS=rstParame("Parame")
end if

if USAANALYTICS="SI" then
	analitycs() 
end if 
%>
<!--<OBJECT RUNAT=server PROGID=ADODB.Recordset id=Detalle name=Detalle > </OBJECT>-->
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=Datos name=Datos></OBJECT>
<%
Dim Rut, strSql, Estado, GlosaError, Codcarr, strsql2, sql20, sql21 ,sqlPar

if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if

if EstaHabilitadaNW (484)="S" then 
	if GetPermisoNW(484) ="N" then
		response.Redirect("MensajesBloqueos.asp")
	else
		strParame="SELECT dbo.Fn_ValorParame('BLOQUEAPAENCUESTAS')Parame"
		set rstParame= Session("Conn").Execute(strParame)		
		if not rstParame.eof then
				BLOQUEAPAENCUESTAS=rstParame("Parame")
			else
				BLOQUEAPAENCUESTAS="" 
		end if
		
		if BLOQUEAPAENCUESTAS="SI" then
			'valida si contesta o no la encuesta
			if TomadeRamo(session("Codcarr"),session("Codcli"),session("ano_ed"),session("periodo_ed"))=0 then
				session("MensajeBloqueosVarios") ="Para acceder a esta opción debe responder sus Encuestas Pendientes."
				response.Redirect("MensajeBloqueo.asp")	
			end if 
		end if 
		
	end if
else
	response.Redirect("MensajeBloqueoHabilita.asp")
end if

Audita 484,"Ingresa a Consulta Cupo"

Rut = RutSinDV(Session("RutAlum"))
codcli=session("codcli")
rutcompleto=trim((Session("logrut")))

if cint(month(date())) = 1 then 
	Mes = "01"
end if 
if cint(month(date())) = 2 then 
	Mes = "02"
end if 
if cint(month(date())) = 3 then 
	Mes = "03"
end if 
if cint(month(date())) = 4 then 
	Mes = "04"
end if 
if cint(month(date())) = 5 then 
	Mes = "05"
end if 
if cint(month(date())) = 6 then 
	Mes = "06"
end if 
if cint(month(date())) = 7 then 
	Mes = "07"
end if 
if cint(month(date())) = 8 then 
	Mes = "08"
end if 
if cint(month(date())) = 9 then 
	Mes = "09"
end if 
if cint(month(date())) > 9 then
	Mes = cint(month(date()))
end if 	

if cint(Day(date())) = 1 then 
	Dia = "01"
end if
if cint(Day(date())) = 2 then 
	Dia = "02"
end if
if cint(Day(date())) = 3 then 
	Dia = "03"
end if
if cint(Day(date())) = 4 then 
	Dia = "04"
end if
if cint(Day(date())) = 5 then 
	Dia = "05"
end if
if cint(Day(date())) = 6 then 
	Dia = "06"
end if
if cint(Day(date())) = 7 then 
	Dia = "07"
end if
if cint(Day(date())) = 8 then 
	Dia = "08"
end if
if cint(Day(date())) = 9 then 
	Dia = "09"
end if
if cint(Day(date())) > 9 then 
	Dia = cint(Day(date()))
end if


FecMat = year(date())&Mes &Dia
'response.Write(FecMat)
'response.End()

if Session("logrut") = "" then
  'rutcompleto =(Session("Rut")
   rutcompleto= Session("RutCliente") 
end if  



if session("codcarr") <> "" then 
            Sql30 = ""
			Sql30 = "Select estacad, codcarpr, jornada, ano from mt_alumno where rut = '" & Session("Rut")& "'"
			Sql30 =  Sql30 & " and codcarpr ='" & session("codcarr") & "' "
			'response.Write(sql30)
			'response.End()
			
			Set Rst30 = Session("Conn").Execute(Sql30)
			if not rst30.eof then 
				if rst30("estacad") <>"" then 
				 GlosaError = rst30("estacad")
				 Codcarr = rst30("Codcarpr")
				 Jornada = rst30("Jornada")
				 AnoPromo = rst30("Ano")
				 'else 
				 'GlosaError = "Año Actual"
			   end if 
			end if
			
			sql20 = " select c.descripcion, c.codcat from mt_client m, mt_catalumno c"
			sql20 = sql20 & " where m.categoria = c.codcat"
			sql20 = sql20 & " and m.codcli = '" & trim(Session("Rut")) & "' "	
			
			set rst20 = Session("Conn").execute(sql20)
			if not rst20.eof then
				Categoria = rst20("descripcion")
				Catalumno = rst20("codcat")
			end if 
			
			
			
			Sql21 = "select dbo.Fn_ValorParame('AnoAdmision') as AnoAdmision, dbo.Fn_ValorParame('PeriodoAdmi') as PeriodoAdmi, dbo.Fn_ValorParame('MatXPromo') as MatXPromo " ' from mt_parame"
			set Rst21 = Session("Conn").Execute(Sql21)
			if not Rst21.eof then 
				AnoAdmi = Rst21("Anoadmision")
				PeriodoAdmi = Rst21("PeriodoAdmi")
				paramMat = rst21("MatXPromo")
			end if 
			
			
else
            Sql30 = "Select estacad, codcarpr, jornada , ano from mt_alumno where rut = '" & Rut & "'"
			'Sql30 = sql30 & "matriculado ='S' "
'			'response.Write(sql30)
			'response.End()
			
			Set Rst30 = Session("Conn").Execute(Sql30)
			if not rst30.eof then 
				if rst30("estacad") <>"" then 
				 GlosaError = rst30("estacad")
				 Codcarr = rst30("codcarpr")
				 Jornada = rst30("Jornada")
				 AnoPromo = rst30("Ano")
				 'else 
'				 GlosaError = "Año Actual"
			   end if 
			end if
			
			
			sql20 = " select c.descripcion, c.codcat from mt_client m, mt_catalumno c"
			sql20 = sql20 & " where m.categoria = c.codcat"
			sql20 = sql20 & " and m.codcli = '" & trim(Session("Rut")) & "' "	
			
			set rst20 = Session("Conn").Execute(sql20)
			if not rst20.eof then
				Categoria = rst20("descripcion")
				Catalumno = rst20("codcat")
			end if 			
			sqlPar="select COALESCE(dbo.Fn_ValorParame('FinancieroNET'),0) as fi"
			set RstPar = Session("Conn").Execute(sqlPar)
			if RstPar("fi")=1 then
			Sql21 = "select dbo.Fn_ValorParame('AnoAdmision') as AnoAdmision, dbo.Fn_ValorParame('PeriodoAdmi') as PeriodoAdmi, dbo.Fn_ValorParame('MatXPromo') as MatXPromo " 
			else
			Sql21 = "SELECT AnoAdmision,PeriodoAdmi,MatXPromo FROM mt_parame"
			end if
			
			set Rst21 = Session("Conn").Execute(Sql21)
			if not Rst21.eof then 
				AnoAdmi = Rst21("AnoAdmision")
				PeriodoAdmi = Rst21("PeriodoAdmi")
				paramMat = rst21("MatXPromo")				
			end if 
end if 			

sql = ""
sql = sql & " select nombre_l from mt_carrer where codcarr = '" & trim(Codcarr) & "'" 

set rst1 = Session("Conn").Execute(sql)
if not rst1.eof then 
	CarreraNombre = rst1("nombre_l")
end if  
CodJornada=Jornada
If Jornada = "D" then 
	Jornada = "DIURNA"
end if 
If Jornada = "V" then
	Jornada = "VESPERTINA"
End if 
If Jornada = "T" then
	Jornada = "TARDE"
end if 
If Jornada = "M" then 
	Jornada = "MAÑANA"
End if 
If Jornada = "E" then 
	Jornada = "EXECUTE"
End if 
If Jornada = "R" then 
	Jornada = "REV"
End if 


Function ValNulo(varDum, intTip) 
    If VarType(varDum) = vbNull Or varDum = "" Then
        Select Case intTip
            Case STR_
                ValNulo = ""
            Case NUM_
                ValNulo = 0
            Case DAT_
                ValNulo = 0 'Null
            Case Else
                ValNulo = "0"
        End Select
    Else
        If intTip = NUM_ Then
            ValNulo = cDbl(varDum)
        Else
            ValNulo = varDum
        End If
    End If
End Function

strsql2 = ""
strsql2 =strsql2 & "select b.nombre,a.ctadocnum as Numero,a.fecven,a.monto,a.saldo "
strsql2= strsql2 & "from mt_ctadoc a, mt_docum b,mt_estdoc c, mt_item i, mt_detubi d "
strsql2= strsql2 & "where a.codcli = '" & trim(Session("Rut")) & "' "
strsql2= strsql2 & "and a.ctadoc = b.tipodoc and a.estado = c.estado  "
strsql2= strsql2 & "and a.item*=i.coditem and a.ubicacion=d.codubifis "
strsql2= strsql2 & "and a.saldo > 0 and convert(varchar,a.fecven,112) < convert(varchar,getdate(),112) and a.estado = 2 order by a.fecven asc "

'response.Write(strsql2)
'response.End()

if bcl_ado(strsql2,rst5)  then
	 'dinero = "Usted Tiene Deuda(s) de Arancel, favor regularizar situaci&oacute;n en Finanzas"
	 situalumno = "De acuerdo a nuestros antecedentes usted (NO) está habilitado para matricularse"
	 situalumno = situalumno & " en el periodo vigente.   Verifique en Registro Académico en Línea /"
	 situalumno = situalumno & " Consulta Situación Actual"
else
	 'dinero ="Usted No tiene Deudas de Arancel con la Universidad"
	 situalumno = ""
end if 


sql1= ""
sql1= sql1 &"select deuda_biblioteca from mt_client "
sql1= sql1 &"where codcli ='" & Rut & "' and coalesce(deuda_biblioteca,'N')='S'"

if bcl_ado(sql1,rst2)  then
	 'deuda = "Usted Tiene deuda Pendiente en Biblioteca, favor regularizar situación"
	 situalumno = "De acuerdo a nuestros antecedentes usted (NO) está habilitado para matricularse"
	 situalumno = situalumno & " en el periodo vigente.   Verifique en Registro Académico en Línea /"
	 situalumno = situalumno & " Consulta Situación Actual"
else
	 'deuda ="Usted No tiene Deuda en Biblioteca"
	 situaumno = ""
end if   

'response.write (paramMat & "-" & AnoPromo)

If paramMat = "N" Or AnoPromo = "" Then
        str = "select monto from mt_listaitem"
        str = str & " Where codcarr='" & trim(Codcarr) & "' "
        str = str & " and ano ='" & AnoAdmi & "' "
        str = str & " and coditem ='1' "
        str = str & " and ('" & FecMat & "' between fecinivig AND fectervig)"
		'response.write(str)
        If BCL_ado(str, rst22) Then
            GetMatricula = ValNulo(rst22("Monto"), NUM_)
        Else
            GetMatricula = 0
        End If
ElseIf paramMat = "S" Then
        str = "select matricula from mt_arancel "
        str = str & " Where codcarr='" & Codcarr & "' "
        str = str & " and ano='" & AnoAdmi & "'"
        str = str & " and periodo='" & PeriodoAdmi & "'"
        str = str & " and anoini <='" & AnoPromo & "' "
        str = str & " and anofin >='" & AnoPromo & "' "
        str = str & " and Catalumno = '" & Catalumno & "' "
        str = str & " and jornada = '" & CodJornada & "' "
        str = str & " and ('" & FecMat & "' Between fecinivig AND fectervig)"
		'response.write(str)
        If BCL_ado(str, rst22) Then
            GetMatricula = ValNulo(rst22("matricula"), NUM_)
        Else
            GetMatricula = 0
        End If
end if 

'response.Write(str)'
session("GetMatricula") = GetMatricula
    
    sql = "select monto from mt_arancel "
    sql = sql & " Where codcarr='" & Codcarr & "' "
    sql = sql & " and Ano = '" & AnoAdmi & "' "
    sql = sql & " and Periodo ='" & PeriodoAdmi & "' " 
    sql = sql & " and anoini <= '" & AnoPromo & "' "
    sql = sql & " and anofin >='" & AnoPromo & "' "
    If Catalumno <> "" Then
      sql = sql & " and CatAlumno = '" & Catalumno & "' "
    End If
    str = str & " and jornada = '" & CodJornada & "' "  
	sql = sql & " and ('" & FecMat & "' Between fecinivig AND fectervig)"
    'response.write(sql)
	'response.End()
	If BCL_ADO(sql, rst23) Then
        GetArancel = ValNulo(rst23("MONTO"), NUM_)
    Else
        GetArancel = 0
    End If

Session("GetArancel") = GetArancel

  
%>
<html>
<head>
<title><%=Session("NombrePestana")%></title>
<link rel="stylesheet" href="file://///Gotzilla/uddweb/admalumnos/css/tex-normales.css" type="text/css">
<link href="css/tex-normales.css" rel="stylesheet" type="text/css">
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=utf-8">
<style type="text/css">
<!--
.Estilo19 {
	color: #FFFFFF;
	font-weight: bold;
}
.Estilo20 {color: #FFFFFF; font-weight: bold; font-size: 16px; }
-->
</style>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<link href="css/textos.css" rel="stylesheet" type="text/css">
<style type="text/css">
<!--
.Estilo21 {color: #000000}
.Estilo22 {font-size: 10px}
-->
</style>
<script language="JavaScript" type="text/javascript">


function CargarReporte()
{
	var opciones="toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=no, resizable=no, width=1024, height=768, top=85, left=125";
window.open("cuponera.asp","",opciones);
}


</script>
</head>
<body>
<table border="0" cellpadding="0" cellspacing="0" align="left">
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
                  <table border="0" cellpadding="0" cellspacing="15" height="308" align="left" width="810">
                    <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
                    <tr>
                      <td width="574" ><img src="imagenes/titulos/T-consulta-arancel.gif" width="400" height="38"></td>
                    </tr>
                    <tr>
                     <td colspan="3" class="Estilo19"><span class="text-normal-celdas">Fecha : </span><span class="titulo">&nbsp;&nbsp; <% =formatdatetime(date(),2)%> 
                      </span></td>
                    </tr>
                     <tr>
                       <td class="Estilo19">&nbsp;</td>
                      <td width="" class="Estilo19" align="right">&nbsp;</td>
                      <td width="204" class="Estilo19" align="right"><span class="titulo"><%=AnoAdmision%>&nbsp;-&nbsp;<%=PeriodoAdmi%></td>
                    </tr>       
                    <tr valign="top"> 
                      <td colspan="3" height="8"><table width="810" border="0" cellpadding="3" cellspacing="2">
                        <tr>
                          <td width="125" height="30"  background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda"><span>Nombre:</span></td>
                          <td width="478" height="30" bgcolor="#DBECF2" class="tex-totales-celda"><%=Session("NomAlum")%></td>
                        </tr>
                        <tr>
                          <td height="30"  background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda"><span>Carrera : </span></td>
                          <td height="30" bgcolor="#DBECF2" class="tex-totales-celda"><%=CarreraNombre%></td>
                        </tr>
                        <tr>
                          <td height="30"  background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda"><span id="lblJornada" >Jornada:</span></td>
                          <td height="30" bgcolor="#DBECF2" class="tex-totales-celda"><%=Jornada%></td>
                        </tr>
                        <tr>
                          <td height="30"  background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda"><span id="lblPromoAlumn">Promoci&oacute;n Alumno: </span></td>
                          <td height="30" bgcolor="#DBECF2" class="tex-totales-celda"><%=AnoPromo%></td>
                        </tr>
                        <tr>
                          <td height="30"  background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda"><span>Categor&iacute;a:</span></td>
                          <td height="30" bgcolor="#DBECF2" class="tex-totales-celda"><%=Categoria%></td>
                        </tr>
                        <tr>
                          <td height="30"  background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda"><span >Estado Acad&eacute;mico:</span></td>
                          <td height="30" bgcolor="#DBECF2" class="tex-totales-celda"><span class="Estilo22"><%=GlosaError%></span></td>
                        </tr>
                        <tr>
                          <td height="30"  background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda"><span id="lblSituAlumn">Situaci&oacute;n Alumno: </span></td>
                          <td height="30" rowspan="2" align="justify" bgcolor="#DBECF2" class="tex-totales-celda"><%=situalumno%></td>
                        </tr>
                        <tr class="text-cabecera-celda">
                          <td height="30" background="imagenes/fdo-cabecera-cel22.jpg"  class="textos">&nbsp;</td>
                          </tr>
                         <tr>
                            <td height="30"  background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda">&nbsp;</td>
                            <td height="30" bgcolor="#DBECF2" class="tex-totales-celda">&nbsp;</td>
                        </tr>          
                        <tr>
                            <td height="30"  background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda"><span >Valor Matr&iacute;cula:</span></td>
                            <td height="30" bgcolor="#DBECF2" class="tex-totales-celda"><span class="Estilo22">$<%=GetMatricula%></span></td>
                        </tr>
                        <tr>
                            <td height="30"  background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda"><span >Valor Arancel:</span></td>
                            <td height="30" bgcolor="#DBECF2" class="tex-totales-celda"><span class="Estilo22">$<%=GetArancel%></span></td>
                        </tr>
                      </table></td>
                    </tr>
                    <tr valign="top" align="left" > 
                      <td colspan="3" height="19"> <div align="center" class="text-normal-celdas">
                        Nota: informaci&oacute;n sujeta a confirmaci&oacute;n.
                      </div>      </td>
                    </tr>
                    <tr valign="top"> 
                      <td colspan="2" height="2"><p>&nbsp;</p><a href="menu_consultas.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','Imagenes/botones/volver-on.gif',1)"><img src="Imagenes/botones/volver-of.gif" width="178" height="45" name="Image1" border="0"></a></td>
                          
                    </tr>
                  </table>
			</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>  

</body>
<%ObjetosLocalizacion("ConsultaCupo.asp")%>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->