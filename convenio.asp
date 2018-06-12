<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%
if Session("CodCli") = "" then
'  Response.Redirect("saltoinicio.htm")
end if

If (session("RutAlum"))=empty Then 
	If (session("RX"))<> Empty Then
		RutCompletoDIG=(session("RX"))
	End if
else 
	RutCompletoDIG = (session("RutAlum"))
End If 

Dim Rut, strSql, Estado, GlosaError, Codcarr, strsql2, sql20, sql21 
Rut = RutSinDV(Session("RutAlum"))
codcli=session("codcli")
rutcompleto=trim((Session("logrut")))

if cint(month(date())) = 1 then 
	Mes = "01"
	MesNombre = ucase(monthName(mes)) 
	
end if 
if cint(month(date())) = 2 then 
	Mes = "02"
	MesNombre = ucase(monthName(mes))
end if 
if cint(month(date())) = 3 then 
	Mes = "03"
	MesNombre = ucase(monthName(mes))
end if 
if cint(month(date())) = 4 then 
	Mes = "04"
	MesNombre = ucase(monthName(mes))
end if 
if cint(month(date())) = 5 then 
	Mes = "05"
	MesNombre = ucase(monthName(mes))
end if 
if cint(month(date())) = 6 then 
	Mes = "06"
	MesNombre = ucase(monthName(mes))
end if 
if cint(month(date())) = 7 then 
	Mes = "07"
	MesNombre = ucase(monthName(mes))
end if 
if cint(month(date())) = 8 then 
	Mes = "08"
	MesNombre = ucase(monthName(mes))
end if 
if cint(month(date())) = 9 then 
	Mes = "09"
	MesNombre = ucase(monthName(mes))
end if 
if cint(month(date())) > 9 then
	Mes = cint(month(date()))
	MesNombre = ucase(monthName(mes))
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
FechaCompleta = now()
Fecha = date()
'response.Write(FecMat)
'response.End()

if Session("logrut") = "" then
  'rutcompleto =(Session("RutCli")
   rutcompleto= Session("RutCliente")
end if  


if session("CarreraAlumno") <> "" then 
            Sql30 = ""
			Sql30 = "Select estacad, codcarpr, jornada, ano from mt_alumno where rut = '" & Session("RutCli")& "'"
			Sql30 =  Sql30 & " and codcarpr ='" & session("CarreraAlumno") & "' "
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
			sql20 = sql20 & " and m.codcli = '" & trim(Session("RutCli")) & "' "	
			
			set rst20 = Session("Conn").execute(sql20)
			if not rst20.eof then
				Categoria = rst20("descripcion")
				Catalumno = rst20("codcat")
			end if 
			
			Sql21 = "select AnoAdmision, PeriodoAdmi, MatXPromo from mt_parame"
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
			sql20 = sql20 & " and m.codcli = '" & trim(Session("RutCli")) & "' "	
			
			set rst20 = Session("Conn").Execute(sql20)
			if not rst20.eof then
				Categoria = rst20("descripcion")
				Catalumno = rst20("codcat")
			end if 			
			
			Sql21 = "select AnoAdmision, PeriodoAdmi, MatXPromo from mt_parame"
			set Rst21 = Session("Conn").Execute(Sql21)
			if not Rst21.eof then 
				AnoAdmi = Rst21("AnoAdmision")
				PeriodoAdmi = Rst21("PeriodoAdmi")
				paramMat = rst21("MatXPromo")				
			end if 
end if 			

sql = ""
sql = sql & " select nombre_c from mt_carrer where codcarr = '" & trim(Codcarr) & "'" 

set rst1 = Session("Conn").Execute(sql)
if not rst1.eof then 
	CarreraNombre = rst1("nombre_c")
end if  

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


    
    sql = "select monto from mt_arancel "
    sql = sql & " Where codcarr='" & Codcarr & "' "
    sql = sql & " and Ano = '" & AnoAdmi & "' "
    sql = sql & " and Periodo ='" & PeriodoAdmi & "' " 
    sql = sql & " and anoini <= '" & AnoPromo & "' "
    sql = sql & " and anofin >='" & AnoPromo & "' "
    If Catalumno <> "" Then
      sql = sql & " and CatAlumno = '" & Catalumno & "' "
    End If
    'If p_combo <> "" Then
     ' sql = sql & " and Combo = '" & p_combo & "' "
    'End If
	
	If Jornada <> "" Then
	  If Jornada = "DIURNA" Then sql = sql & " and Jornada = 'DIURNA'"
	  If jornada = "VESPERTINA" Then sql = sql & " and Jornada = 'VESPERTINA'"
	  If Jornada = "TARDE" Then sql = sql & " and Jornada = 'TARDE'"
	  If Jornada = "EXECUTE" Then sql = sql & " and Jornada = 'EXECUTE'"
	  If Jornada = "REV" Then sql = sql & " and Jornada = 'REV'"
	  If Jornada = "MAÑANA" Then sql = sql & " and Jornada = 'MAÑANA'"
	End If    
	sql = sql & " and ('" & FecMat & "' Between fecinivig AND fectervig)"
    'response.write(sql)
	'response.End()
	If BCL_ADO(sql, rst23) Then
        GetArancel = ValNulo(rst23("MONTO"), NUM_)
    Else
        GetArancel = 0
    End If

Session("GetArancel") = GetArancel

	
	sql66 = ""
	sql66 = sql66 & "select * from mt_client where codcli ='" & Trim(Session("RutCli")) &"'"
	
	If BCL_ADO(sql66, rst66) Then
        Nom_Alumno = ValNulo(rst66("Nombre"), STR_)
		Ape_Pat_Alumno = ValNulo(rst66("Paterno"), STR_)
		Ape_Mat_Alumno = ValNulo(rst66("Materno"), STR_)
		Comuna_Alumno = ValNulo(rst66("Comuna"), STR_)
		Ciudad_Alumno = ValNulo(rst66("CiudadAct"), STR_)
		Dir_Alumno = ValNulo(rst66("Diractual"), STR_)
        
    End If


  sql69 = ""
  sql69 = sql69 & "select ctapagnum, monto from mt_ctapag where ctapag ='4' and codcli ='" & Trim(Session("RutCli")) &"' and ano ='" & year(date()) & "'"
  if BCL_ADO(sql69,rst69) then
  		Folio = ValNulo(rst69("ctapagnum"), NUM_)
		MontoPesos = ValNulo(rst69("monto"), NUM_)
  end if 
  
	if Folio = "" then 
		Folio = 0
	end if  
	 
	if MontoPesos = "" then 
		MontoPesos= 0
	end if
	  
  sql68 = "select monto from mt_uf where ano = '" & year(date())& "'"
  if BCL_ADO(sql68, rst68) Then
  		UF = Valnulo(rst68("monto"), NUM_) 
  end if 
  
  UTM = MontoPesos / UF
  UTM = Round((((UTM * 10) + 0.00005) / 10), 4)

%>
<script language=Javascript> 
		window.open('pagare.asp' ,'_blank',"toolbar=0,location=0,status=0,menubar=0 ,scrollbars=yes,resizable=yes");
		//window.print()  
</script>
<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title>Portal de Alumnos en L&iacute;nea</title>
<body >
<table border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
	<td>
<table border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
	<td>
    	<table width="900" height="76">
         <tr>
          	<td colspan="2" align="right"><img align="baseline" src="Imagenes/LogoCupon/izquierda.jpg" width="146" height="98"></td>
		          <td width="621" align="center"><p><strong>UNIVERSIDAD METROPOLITANA DE CIENCIAS DE LA EDUCACI&Oacute;N</strong></p>
		   <p><strong>FONDO SOLIDARIO DE CREDITO UNIVERSITARIO</strong>   		   </p><BR></td>
          	<td width="98" class="textos" align="right" valign="top">2/2</td>
          </tr>
        <tr>
        </table>
    	<table width="861">
      <tr>
          	<td height="" colspan="2" align="center"></td>
            <td height="" colspan="2" align="center"></td>
            <td width="316" align="RIGHT" colspan="2"><p class="tex-totales-celda">FOLIO No : <%=Folio%></p>
           <p class="tex-totales-celda"><strong>No RUT DEL ALUMNO : </strong><%=RutCompletoDIG%></p></td>
          </tr>
        <tr>
      </table>        
    	<table width="900">
<tr>
          	<td colspan="2" align="center">&nbsp;</td>	          
         <td width="258" align="center"><p><strong>CONVENIO</strong></p></td>
          	<td width="318" class="textos">&nbsp;</td>
          </tr>
        <tr>
      </table>  
		<table width="830" border="0" cellpadding="0" cellspacing="0" align="center">
         <tr>
      		<td width="830" align="justify"><BR><p align="justify">En Santiago, a 31/03/<%=year(fecha)%>, entre la Universidad Metropolitana de Ciencias de la Educaci&oacute;n RUT No 60.910.047-8, en adelante la Universidad, representada por su Rector don JAIME ESPINOSA ARAYA, c&eacute;dula de identidad No 6.069.050-2, ambos domiciliados para todos los efectos en AV. Jos&eacute; P. Alessandri 774, y don(a) <%=Ape_Pat_alumno%>&nbsp;<%=Ape_Mat_Alumno%>&nbsp;<%=RTRIM(Nom_Alumno)%>, RUT <%=RutCompletoDIG%>, en adelante el alumno, domiciliado en <%=RTRIM(Dir_Alumno)%>, Comuna de <%=RTRIM(Comuna_Alumno)%>, Ciudad de <%=Ciudad_Alumno%> se ha celebrado el siguiente convenio: </p><p></p></td>
        </tr>
        <tr>
      		<td width="830" align="justify"><p align="justify">1.   La Universidad se compromete a proporcionar, habido cumplimiento por parte del alumno de la reglamentaci&oacute;n interna de la instituci&oacute;n, los servicios educacionales correspondientes al Programa o carrera en que el alumno se encuentre matriculado. </p></td>
        </tr>
        <tr>
      		<td width="830" align="justify"><p align="justify">2.  La Universidad es due&ntilde;a de un patrimonio denominado Fondo Solidario de Cr&eacute;dito Universitario, se&ntilde;alado en la Ley No 19.287, sucesor y continuador legal del Fondo de Cr&eacute;dito Universitario, creado en virtud de los art&iacute;culos No 70 y siguientes de la Ley No 18.591 del 3 de Enero de 1987. </p>
            </td>
        </tr>
        <tr>
      		<td width="830" align="justify"><p align="justify">3.  Con cargo a dicho fondo, la Universidad viene en otorgar al alumno un cr&eacute;dito cuyo monto asciende a $ <%=UTM%>, Unidades Tributarias Mensuales, fijada por ley y actualizada por el S.I.I. o el organismo que lo reemplace, seg&uacute;n el valor de la misma para el mes de Marzo del presente a&ntilde;o, destinado a financiar total o parcialmente, el arancel de matr&iacute;cula de la Carrera o programa de la Universidad en que est&eacute; matriculado. Dicho cr&eacute;dito, devengar&aacute; un inter&eacute;s anual del 2% desde la fecha de suscripci&oacute;n de los instrumentos representativos del cr&eacute;dito. </p>
            </td>
        </tr> 
        <tr>
      		<td width="830" align="justify"><p align="justify">4.  Para el otorgamiento efectivo de dicho cr&eacute;dito, el alumno suscribir&aacute; un pagar&eacute; ante Notario P&uacute;blico, destinado a garantizar el cumplimiento de la obligaci&oacute;n contra&iacute;da. </p>
            </td>
        </tr>          
        <tr>
      		<td width="830" align="justify"><p align="justify">5.  El otorgamiento y las condiciones para la devoluci&oacute;n del cr&eacute;dito, se regir&aacute;n por lo dispuesto en la Ley No 19.287 del 04 de Febrero de 1994 y sus reglamentos, normativa que forma parte integrante del presente convenio, la que el alumno declara conocer y aceptar. </p>
            </td>
        </tr> 
        <tr>
      		<td width="830" align="justify"><p align="justify">6.  La obligaci&oacute;n de pago se har&aacute; exigible una vez transcurridos dos a&ntilde;os desde el egreso del alumno de la Universidad por haber cursado sus estudios completos, est&eacute; o no en posesi&oacute;n del T&iacute;tulo Profesional o grado respectivo, o transcurridos 2 a&ntilde;os consecutivos sin que el alumno se matriculare en alguna de las instituciones a que se refiere el art. 70 de la Ley 18.591. Al momento de hacerse exigible la obligaci&oacute;n, el alumno deber&aacute; pagar anualmente una suma equivalente al 5% del total de sus ingresos que haya obtenido en el a&ntilde;o inmediatamente anterior expresado en Unidades Tributarias Mensuales, correspondientes a cada uno de los meses en que se percibieron los ingresos. Todo lo anterior sin perjuicio de lo se&ntilde;alado en el art. 10 de la Ley No 19.287. Si el alumno no acreditare sus ingresos ante el Administrador General del Fondo Solidario de Cr&eacute;dito Universitario de la Universidad, en la forma establecida en la Ley No 19.287, a m&aacute;s tardar el &uacute;ltimo d&iacute;a h&aacute;bil del mes de Mayo del a&ntilde;o en que le corresponda efectuar el pago, se le fijar&aacute; una cuota equivalente al mayor valor entre en doble del pago anual anterior o el 20% del saldo deudor. </p>
            </td>
        </tr> 
        <tr>
      		<td width="830" align="justify"><p align="justify">7.  Si se determinare que el alumno falt&oacute; a la verdad en la informaci&oacute;n proporcionada sobre los ingresos percibidos, el total de la deuda se har&aacute; exigible de inmediato y devengar&aacute; el inter&eacute;s penal establecido en el art. 15 de la Ley 19.287, sin perjuicio de la responsabilidad penal que correspondiere en conformidad a lo dispuesto en el art&iacute;culo 210 del C&oacute;digo Penal. En virtud de lo dispuesto en el Art. 6 de la Ley No 19.287, se deja constancia que, en caso de que el alumno haya faltado a la verdad en los antecedentes proporcionados a la universidad para acreditar su condici&oacute;n socioecon&oacute;mica, la cual ha sido factor determinante para el otorgamiento de este cr&eacute;dito, perder&aacute; autom&aacute;tica e irrevocablemente, el derecho a obtener cr&eacute;dito universitario para el financiamiento de sus estudios ante cualquier instituci&oacute;n que lo otorgue, y se le har&aacute; exigible de inmediato el total del cr&eacute;dito que da cuenta del presente convenio, deveng&aacute;ndose adem&aacute;s el inter&eacute;s penal establecido en el Art. 15 de la Ley No 19.287 , sin perjuicio de la responsabilidad penal que correspondiere, en conformidad a lo dispuesto en el Art. 210 del C&oacute;digo Penal. </p></td>
        </tr> 
        <tr>
      		<td width="830" align="justify"><p align="justify">8.  La suma adeudada, se pagar&aacute; en el equivalente en moneda nacional por el valor de la Unidad Tributaria Mensual fijada por Ley y actualizada por el Servicio de Impuestos Internos, o el organismo que lo reemplace, vigente a la &eacute;poca del pago efectivo. </p></td>
        </tr> 
        <tr>
      		<td width="830" align="justify"><p align="justify">9.  Para todos los efectos del presente convenio, las partes fijan su domicilio en la comuna y ciudad de Santiago, someti&eacute;ndose a la jurisdicci&oacute;n de sus Tribunales Ordinarios de Justicia. </p></td>
        </tr> 
        <tr>
      		<td width="830" align="justify"><p align="justify">10.  El presente instrumento se firma en dos ejemplares de igual tenor y fecha, quedando uno de ellos en poder de cada una de las partes.      		  </p>   </td>
        </tr>  
        <tr>
      		<td width="830" align="justify"><p align="justify">10.  El presente instrumento se firma en dos ejemplares de igual tenor y fecha, quedando uno de ellos en poder de cada una de las partes.      		  </p>   </td>
        </tr>
        <tr>
      		<td width="830" align="justify"><p align="justify">11. La personer&iacute;a de don(a) JAIME ESPINOSA ARAYA, consta en el Decreto Supremo No 258 del a&ntilde;o 2009 otorgado por el Ministerio de Educaci&oacute;n</p>
      		  <p align="justify">&nbsp;</p>
      		  <p align="justify">&nbsp;</p></td>
        </tr>                                                                                                            
		</table>
        <table width="861">
      <tr>
      		<td width="29"></td>
          	<td width="350" align="center"><p>_________________________________</p>
          	  <p align="center">SR.JAIME ESPINOSA ARAYA</p>
          	  <p align="center">6.069.050-2</p>
          	  <p align="center">RECTOR</p></td>
            <td width="85" height="" align="center"></td>
			<td width="377" align="center"><p>_________________________________</p>
          	  <p align="center"><%=Ape_Pat_alumno%>&nbsp;<%=Ape_Mat_Alumno%>&nbsp;<%=RTRIM(Nom_Alumno)%></p>
          	  <p align="center"><%=RutCompletoDIG%></p>
          	  <p align="center">ALUMNO</p></td>
          </tr>
        <tr>
      </table>        
	</td>
  </tr>
</table>
</body>
</html>
<script language=Javascript> 
		//window.open('pagare.asp' ,'_blank',"toolbar=0,location=0,status=0,menubar=0 ,scrollbars=yes,resizable=yes");
		//window.print()  
</script>
<%
'response.Buffer "true"
'Response.AddHeader "Content-Disposition","attachment; filename=pagare.doc"
'Response.ContentType = "application/vnd.word"
'Response.Buffer = True
'Response.ContentType = "application/vnd.ms-word"
'Response.AddHeader "content-disposition", "inline; filename = ASP_Word_Doc.doc" 
%>
<!--#INCLUDE file="include/desconexion.inc" -->