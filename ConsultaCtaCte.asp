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
Dim Rut, strSql, Estado, GlosaError,strsql2

if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if

if EstaHabilitadaNW (482)="S" then 
	if GetPermisoNW(482) ="N" then
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
Audita 482,"Ingresa a Consulta Financiera"

Rut = RutSinDV(Session("RutAlum"))
codcli=session("codcli")
rutcompleto=trim((Session("logrut")))

if Session("logrut") = "" then
  'rutcompleto =(Session("RutCli")
   rutcompleto= Session("RutCliente")
end if  


if session("CarreraAlumno") <> "" then 
            Sql30 = ""
			Sql30 = "Select estacad from mt_alumno where rut = '" & Session("RutCli")& "'"
			Sql30 =  Sql30 & " and codcarpr ='" & session("CarreraAlumno") & "' "
			'response.Write(sql30)
			'response.End()
			
			Set Rst30 = Session("Conn").Execute(Sql30)
			if not rst30.eof then 
				if rst30("estacad") <>"" then 
				 GlosaError = rst30("estacad")
				 'else 
				 'GlosaError = "Año Actual"
			   end if 
			end if
else
            Sql30 = "Select estacad from mt_alumno where rut = '" & Rut & "'"
			'Sql30 = sql30 & "matriculado ='S' "
'			'response.Write(sql30)
			'response.End()
			
			Set Rst30 = Session("Conn").Execute(Sql30)
			if not rst30.eof then 
				if rst30("estacad") <>"" then 
				 GlosaError = rst30("estacad")
				 'else 
'				 GlosaError = "Año Actual"
			   end if 
			end if
end if 			

'Datos.Open strSql, Conn
'Estado = Trim(Datos(0))
'Select Case Estado
'	Case "ELIMINADO":
'	     GlosaError = "Ud. se encuentra Eliminado de los registros de la  Universidad."
'	Case "EGRESADO"	 
'	     GlosaError = "Su estado actual es de Egresado en nuestros registros."
'	Case "TITULADO"	 
'	     GlosaError = "Su estado actual es de Titulado en nuestros registros."
'	Case "SUSPENDIDO"	 		
'	     GlosaError = "Ud. se encuentra suspendido de la Universidad."
'	Case else
'		 GlosaError = "Vigente/Alumno Regular"
'End Select
'Datos.Close()
'Conn.Close()

'dim strsql,strsql2

'rut=trim(Session("RutAlum"))
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

'codcli=200011001
'response.Write(codcli)
'response.end

strsql2 = ""
strsql2 =strsql2 & "select sum(a.saldo) as total "
'strsql2 =strsql2 & "select b.nombre,a.ctadocnum as Numero,a.fecven,a.monto,a.saldo "
strsql2= strsql2 & "from mt_ctadoc a, mt_docum b,mt_estdoc c, mt_item i, mt_detubi d "
strsql2= strsql2 & "where a.codcli = '" & trim(Session("RutCli")) & "' "
strsql2= strsql2 & "and a.ctadoc = b.tipodoc and a.estado = c.estado  "
strsql2= strsql2 & "and a.item*=i.coditem and a.ubicacion=d.codubifis "
strsql2= strsql2 & "and a.saldo > 0 and a.estado = 2 and convert(varchar,a.fecven,112) < convert(varchar,getdate(),112) "

'strsql2= strsql2 & "order by a.fecven asc "

strsql2="SP_CONSULTA_SUMAMONTO_CCORRIENTE_PA '" & trim(Session("RutCli")) & "','" & session("codcarr") & "',null,'NO'" 
'RESPONSE.Write(STRSQL2)
'RESPONSE.End()

Set Rst = Session("Conn").Execute(StrSql2)

if not rst.eof then  
       total = valnulo(rst("MONTO"),num_) 
	    
	 else
	   total = "0" 				
end if

%>
<html>
<head>
<title><%=Session("NombrePestana")%></title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=iso-8859-1"> 
<!-- Fireworks 4.0  Dreamweaver 4.0 target.  Created Wed Nov 06 17:26:22 GMT-0300 2002-->
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
</head>
<body bgcolor="#ffffff" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
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
              <table border="0" cellpadding="0" cellspacing="15"  align="left" >
                <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
                <tr>
                  <td width="294"><img src="imagenes/titulos/T-consulta-financiera.gif" width="400" height="38"></td>
                </tr>
                <tr valign="top"> 
                  <td colspan="3"><table width="810" border="0" cellpadding="3" cellspacing="2">
                    <tr >
                      <td width="122" height="30"background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda"><span>Nombre:</span></td>
                      <td width="428" height="30" bgcolor="#DBECF2" class="tex-totales-celda"><span class="Estilo22"><font face="Arial, Helvetica, sans-serif">
                        <rutcompleto>
                        <%=Session("NomAlum")%></font></span></td>
                    </tr>
                    <tr>
                      <td height="30" background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda"><span id="lblRut">Rut: </span></td>
                      <td height="30" bgcolor="#DBECF2" class="tex-totales-celda"><span class="Estilo22"><font face="Arial, Helvetica, sans-serif"><%=rutcompleto%></font></span></td>
                    </tr>
                    <tr>
                      <td height="30" background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda"><span>Estado Acad&eacute;mico:</span></td>
                      <td height="30" bgcolor="#DBECF2" class="tex-totales-celda"><span class="Estilo22"><%=GlosaError%></span></td>
                    </tr>
                    <tr>
                      <td height="30" background="imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda"><span >Saldo Pendiente: </span></td>
                      <td height="30" bgcolor="#DBECF2" class="tex-totales-celda"><span class="Estilo22">$ <%=formatnumber (total,0)%></span></td>
                    </tr>
                  </table></td>
                </tr>
                <tr valign="top" align="left" > 
                  <td colspan="3" height="19"> <div align="center" class="text-normal-celdas">
                    Nota: informaci&oacute;n sujeta a confirmaci&oacute;n.
                  </div>      </td>
                </tr>
                <tr valign="top"> 
                   <td valign="top"><a href="menu_consultas.asp" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','Imagenes/botones/volver-on.gif',1)"><img src="Imagenes/botones/volver-of.gif" width="178" height="45" name="Image1" border="0"></a></td>
                </tr>
              </table>
			</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>              

</body>
<%ObjetosLocalizacion("ConsultaCtaCte.asp")%>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->