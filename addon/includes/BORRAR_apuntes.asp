<%
dim TablaRamos
dim TablaProfesores
dim TablaDocumentos
dim HeadDoctos
dim BodyDoctos
dim FooterDoctos
dim Bloque

Set Rs1 = Session("Conn").Execute("SELECT DISTINCT ano FROM ra_nota WHERE codcli='" & Session ("CarreraEnCurso") & "' ORDER BY ano") 
if Not Rs1.Bof or Not Rs1.Eof then
	Rs1.MoveFirst
	While Not Rs1.Eof
		Bloque = Bloque & 	"<table width='100%'  border='1' cellpadding='5' cellspacing='0'>" &_
							"<tr><td width='100%' class='caption'><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onclick=""expandit('" & Rs1("ANO") &"');"">Documentos Publicados A&ntilde;o " & Rs1("ANO") & "</a></td></tr>" &_
							"<tr><td valign='top' style='display: none;' id='" & Rs1("ANO") & "'>"

Semestres_1y2 =  "<table width='100%'  border='1' cellpadding='0' cellspacing='0'>" &_
				"<tr>" &_
				"<td width='5%' align='center' valign='top'><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onclick=""expandit('" & Rs1("ANO") & "_I');"">I  Semestre</a></td>" &_
				"<td width='95%' valign='top'>"
TablaRamos=""
TablaProfesores=""
TablaDocumentos=""

Semestre1="<table width='100%'  border='1' cellpadding='0' cellspacing='0' id='" & Rs1("ANO") &"_I' style='display: none;'><tr><td valign='top'>"
Set Rs2 = Session("Conn").Execute("SELECT CODRAMO,ANO,PERIODO from ra_nota where codcli='" & Session ("CarreraEnCurso") & "' and ano='" & Rs1("ANO") & "' and periodo='1' order by ano,periodo")
if Not Rs2.Bof or Not Rs2.Eof _
	then
		Rs2.MoveFirst
		While Not Rs2.Eof
			Set Rs3 = Session("Conn").Execute("SELECT NOMBRE from ra_ramo where codramo='" & Rs2("CODRAMO") & "'")
			TablaRamos = TablaRamos & "<table width='100%'  border='1' cellpadding='3' cellspacing='0' class='smalltext'>" &_
									  "<tr><td bgcolor='#FFFFCC'><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onclick=""expandit('" & Rs2("CODRAMO") & "');"">(" & Rs2("CODRAMO") & "):" & Rs3("NOMBRE") & "</a></td>" &_
									  "</tr><tr><td valign='top' style='display: none;' id='" & Rs2("CODRAMO") & "'>"									  
									  TablaProfesores=""
Set Rs4 = oCmd.Execute("select DISTINCT CodProfe from ca_doctos where CodigoRamo='" & Rs2("CODRAMO") & "'")
if Not Rs4.Bof or Not Rs4.Eof _
	then
		Rs4.MoveFirst
		While Not Rs4.Eof
			Set Rs6=Session("Conn").Execute("select * from ra_profes where codprof='" & Rs4("CodProfe") & "'")
			if Not Rs6.Bof or Not Rs6.Eof _ 
				then
					TablaProfesores = TablaProfesores & "<table width='100%'  border='1' cellpadding='0' cellspacing='0' class='smalltext'>" &_
									"<tr><td class='maintitle'><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onClick=""expandit('" & Rs2("CODRAMO") & "_" & Rs6("CODPROF") & "');"">" & Rs6("NOMBRES") & " " & Rs6("AP_PATER") & " " & Rs6("AP_MATER") & "</a></td>" &_
									"</tr><tr><td valign='top' style='display: none;' id='" & Rs2("CODRAMO") & "_" & Rs6("CodProf") & "'>"
TablaDocumentos=""
HeadDoctos=""
BodyDoctos=""
FooterDoctos=""
Set Rs5=oCmd.Execute("select * from ca_doctos where CodigoRamo='" & Rs2("CODRAMO") & "' and CodProfe='" & Rs6("CodProf") & "'")
if Not Rs5.Bof or Not Rs5.Eof _
	then
		Rs5.MoveFirst
		HeadDoctos="<table width='100%'  border='1' cellpadding='2' cellspacing='0'><tr><td width='7%' align='center' valign='middle'>Fecha Publicaci&oacute;n</td><td colspan='2'>Detalle de Documentaci&oacute;n / Apuntes </td></tr>"			
		While Not Rs5.Eof
			BodyDoctos = BodyDoctos & "<tr><td align='center' valign='middle'>" & Rs5("FechaUpLoad") & "</td><td width='89%'>" & Rs5("DetalleDocumento") & "</td><td width='4%'><a href='doctos/" & Rs2("CODRAMO") & "/" & Rs6("CODPROF") & "/" & Rs5("IdArchivo") & "'>xxx<img src='addon/imagenes/boton_bajar2.gif' width='37' height='27' border='none'>xx</a></td></tr>"
			Rs5.MoveNext
		Wend
		FooterDoctos = "</table>"
	else
		BodyDoctos = BodyDoctos & "<!--(Sin Datos)//-->"
		FooterDoctos = ""
End if
TablaDocumentos=HeadDoctos & BodyDoctos & FooterDoctos
					TablaProfesores = TablaProfesores & TablaDocumentos & "</td></tr></table>"
				else
					TablaProfesores = TablaProfesores & "<!--(Sin Datos)//-->"
			end if
			Rs4.MoveNext
		Wend
	else
		TablaProfesores = TablaProfesores & "<!--(Sin Datos)//-->"
End If
			TablaRamos = TablaRamos & TablaProfesores & "</td></tr></table>"
			Rs2.MoveNext
		Wend
	else
		TablaRamos ="<!--(Sin Datos)//-->"
End if
Semestre1 = Semestre1 & TablaRamos & "</td></tr></table>"

Semestres_1y2 = Semestres_1y2 & Semestre1 & "</td></tr><tr>" &_
				"<td width='5%' align='center' valign='top'><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onclick=""expandit('" & Rs1("ANO") & "_II');"">II  Semestre</a></td>" &_
				"<td width='95%' valign='top'>"
TablaRamos=""
TablaProfesores=""
TablaDocumentos=""

Semestre2 = "<table width='100%'  border='1' cellpadding='0' cellspacing='0' id='" & Rs1("ANO") &"_II' style='display: none;'><tr><td valign='top'>"

Set Rs2 = Session("Conn").Execute("SELECT CODRAMO,ANO,PERIODO from ra_nota where codcli='" & Session ("CarreraEnCurso") & "' and ano='" & Rs1("ANO") & "' and periodo='2' order by ano,periodo")
if Not Rs2.Bof or Not Rs2.Eof _
	then
		Rs2.MoveFirst
		While Not Rs2.Eof
			Set Rs3 = Session("Conn").Execute("SELECT NOMBRE from ra_ramo where codramo='" & Rs2("CODRAMO") & "'")
			TablaRamos = TablaRamos & "<table width='100%'  border='1' cellpadding='3' cellspacing='0' class='smalltext'>" &_
									  "<tr><td bgcolor='#FFFFCC'><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onclick=""expandit('" & Rs2("CODRAMO") & "');"">(" & Rs2("CODRAMO") & "):" & Rs3("NOMBRE") & "</a></td>" &_
									  "</tr><tr><td valign='top' style='display: none;' id='" & Rs2("CODRAMO") & "'>"
									  TablaProfesores=""
Set Rs4 = oCmd.Execute("select DISTINCT CodProfe from ca_doctos where CodigoRamo='" & Rs2("CODRAMO") & "'")
if Not Rs4.Bof or Not Rs4.Eof _
	then
		Rs4.MoveFirst
		While Not Rs4.Eof
			Set Rs6=Session("Conn").Execute("select * from ra_profes where codprof='" & Rs4("CodProfe") & "'")
			if Not Rs6.Bof or Not Rs6.Eof _ 
				then
					TablaProfesores = TablaProfesores & "<table width='100%'  border='1' cellpadding='0' cellspacing='0' class='smalltext'>" &_
									"<tr><td class='maintitle'><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onClick=""expandit('" & Rs2("CODRAMO") & "_" & Rs6("CODPROF") & "');"">" & Rs6("NOMBRES") & " " & Rs6("AP_PATER") & " " & Rs6("AP_MATER") & "</a></td>" &_
									"</tr><tr><td valign='top' style='display: none;' id='" & Rs2("CODRAMO") & "_" & Rs6("CodProf") & "'>"

TablaDocumentos=""
HeadDoctos=""
BodyDoctos=""
FooterDoctos=""
Set Rs5=oCmd.Execute("select * from ca_doctos where CodigoRamo='" & Rs2("CODRAMO") & "' and CodProfe='" & Rs6("CodProf") & "'")
if Not Rs5.Bof or Not Rs5.Eof _
	then
		Rs5.MoveFirst
		HeadDoctos="<table width='100%'  border='1' cellpadding='2' cellspacing='0'><tr><td width='7%' align='center' valign='middle'>Fecha Publicaci&oacute;n</td><td colspan='2'>Detalle de Documentaci&oacute;n / Apuntes </td></tr>"			
		While Not Rs5.Eof
			BodyDoctos = BodyDoctos & "<tr><td align='center' valign='middle'>" & Rs5("FechaUpLoad") & "</td><td width='89%'>" & Rs5("DetalleDocumento") & "</td><td width='4%'><a href='#'><img src='addon/imagenes/boton_bajar2.gif' width='37' height='27' border='none'></a></td></tr>"
			Rs5.MoveNext
		Wend
		FooterDoctos = "</table>"
	else
		BodyDoctos = BodyDoctos & "<!--(Sin Datos)//-->"
		FooterDoctos = ""
End if
TablaDocumentos=HeadDoctos & BodyDoctos & FooterDoctos

					TablaProfesores = TablaProfesores & TablaDocumentos & "</td></tr></table>"
				else
					TablaProfesores = TablaProfesores & "<!--(Sin Datos)//-->"
			end if
			Rs4.MoveNext
		Wend
	else
		TablaProfesores = TablaProfesores & "<!--(Sin Datos)//-->"
End If

			TablaRamos = TablaRamos & TablaProfesores & "</td></tr></table>"
			Rs2.MoveNext
		Wend
End if

Semestre2 = Semestre2 & TablaRamos & "</td></tr></table>"
Semestres_1y2 = Semestres_1y2 & Semestre2 & "</td></tr></table>"
		Bloque = Bloque & Semestres_1y2 & "</td></tr></table>"
		Rs1.MoveNext
	wend
end if
%>