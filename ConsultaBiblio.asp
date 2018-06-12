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

<OBJECT RUNAT=server PROGID=ADODB.Recordset id=Datos name=Datos></OBJECT>

<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if

if EstaHabilitadaNW (481)="S" then 
	if GetPermisoNW(481) ="N" then
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
Audita 481,"Ingresa a Consulta Biblioteca"
%>
<%
Dim Rut, strSql, Estado, GlosaError,strsql2
'response.Write("RUT ALUMNO" + Session("RutAlum"))
'response.End()

Rut = RutSinDV(Session("RutAlum"))

if Session("RutAlum") = "" then
   Rut = Session("RutCli")
end if 

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
			Sql30 =  Sql30 & " and codcli ='" & (codcli) & "' "
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



'codcli=200011001
'response.Write(codcli)
'response.end
strsql2 = ""
strsql2 =strsql2 & "select b.nombre,a.ctadocnum as Numero,a.fecven,a.monto,a.saldo "
strsql2= strsql2 & "from mt_ctadoc a, mt_docum b,mt_estdoc c, mt_item i, mt_detubi d "
strsql2= strsql2 & "where a.codcli = '" & trim(Session("RutCli")) & "' "
strsql2= strsql2 & "and a.ctadoc = b.tipodoc and a.estado = c.estado  "
strsql2= strsql2 & "and a.item*=i.coditem and a.ubicacion=d.codubifis "
strsql2= strsql2 & "and a.saldo > 0 order by a.fecven asc "
'RESPONSE.Write(STRSQL)
'RESPONSE.End()

Set Rst = Session("Conn").Execute(StrSql2)


sql1= ""
sql1= sql1 &"select coalesce(deuda_biblioteca,'')deuda_biblioteca from mt_client "
sql1= sql1 &"where codcli ='" & Rut & "'"

set Rst2 = Session("Conn").Execute(Sql1)

            if not rst2.eof then
		         if rst2("deuda_biblioteca") = "" or rst2("deuda_biblioteca")= "N" then
			         deuda = rst2("deuda_biblioteca") 
					    deuda = "No tiene Deuda"		
					 else
			          deuda= rst2("deuda_biblioteca") 
					  deuda  = "Tiene deuda Pendiente, favor regularizar situación" 
			     end if 
		    end if 		   
		
   


%>

<head>
<title><%=Session("NombrePestana")%></title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; CHARSET=utf-8">
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<link href="css/textos.css" rel="stylesheet" type="text/css">
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
              <table border="0" cellpadding="0" cellspacing="15" height="248" align="left" width="745">
                <!-- fwtable fwsrc="f-der.png" fwbase="f-der.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
                <tr>
                  <td width="715"><img src="imagenes/titulos/T-consulta-biblioteca.gif" width="400" height="38"></td>
                </tr>
                <tr valign="top"> 
                  <td height="8"><table width="810" border="0" cellpadding="3" cellspacing="2">
                    <tr>
                      <td width="155" height="30" background="imagenes/fdo-cabecera-cel22.jpg" class="textos"><span class="Estilo17 Estilo18 Estilo19"><span class="text-cabecera-celda">Nombre:</span></span></td>
                      <td width="553" bgcolor="#DBECF2" class="tex-totales-celda"><span class="Estilo18"><font face="Arial, Helvetica, sans-serif">
                        <rutcompleto>
                        <%=Session("NomAlum")%></font></span></td>
                    </tr>
                    <tr>
                      <td height="30" background="imagenes/fdo-cabecera-cel22.jpg" class="textos"><span class="Estilo17 Estilo18 Estilo19"><span id="lblRut" class="text-cabecera-celda">Rut : </span></span></td>
                      <td bgcolor="#DBECF2" class="tex-totales-celda"><span class="Estilo18"><font face="Arial, Helvetica, sans-serif"><%=rutcompleto%></font></span></td>
                    </tr>
                    <tr>
                      <td height="30" background="imagenes/fdo-cabecera-cel22.jpg" class="textos"><span class="Estilo17 Estilo18 Estilo19"><span class="text-cabecera-celda">Deuda Biblioteca: </span></span></td>
                      <td bgcolor="#DBECF2" class="tex-totales-celda"><span class="Estilo18"><font face="Arial, Helvetica, sans-serif"><%=deuda%></font></span></td>
                    </tr>
                    <tr>
                      <td height="30" background="imagenes/fdo-cabecera-cel22.jpg" class="textos"><span class="Estilo17 Estilo18 Estilo19"><span class="text-cabecera-celda">Estado: </span></span></td>
                      <td bgcolor="#DBECF2" class="tex-totales-celda"><span class="Estilo18"><%=GlosaError%></span></td>
                    </tr>
                  </table></td>
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
<%ObjetosLocalizacion("ConsultaBiblio.asp")%>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->