
<!--#include file="funciones.inc" -->
<%
if Session("eMailAlumno")="" _
	then
		eMail_go = "<table width='100%'  border='0' cellpadding='2' cellspacing='0'><tr><td width='9%'><img src='addon/imagenes/alerta2.gif' width='109' height='109'></td><td width='91%' align='center' valign='middle'>NO PUEDE ENVIAR MAIL SIN UNA DIRECCION DE CORREO VALIDA<br>ACTUALICE SU INFORMACION EN PAGINA WEB TOMA DE RAMOS</td></tr></table>"
	else
		Ramos = Session("matRamos")
		Set mt_carrer=dbMat.Execute("select CODCARR, DIRECTOR from mt_carrer where CODCARR='" & Session("CodigoCarrera") & "'")
		if Not mt_carrer.Bof or Not mt_carrer.Eof _
			then
				'// ACTUALIZAR AKI Set ca_mail_jc=oCmd.Execute("select CodCarrera, Mail from ca_Mail_jc where CodCarrera='" & Session("CodigoCarrera") & "'")
				JC="<input name='chkMail_JC' type='checkbox' id='chkMail_JC' value='aki_va_el_mail'>&nbsp;JEFE CARRERA - " & mt_carrer("DIRECTOR")
			else
				JC="* Sin Jefe Carrera Envie Mail a la Siguiente Direccion<br><input name='chkMail_JC' type='checkbox' id='chkMail_JC' value='hellreiser9@hotmail.com'>Administracion caENAC"
		End if
		
		Set mt_carrer = Nothing
'//HAB		Set ca_Mail_jc = Nothing		

eMail_go =  "<form method='post' action='main.asp?caID=5' id='frmMail' name='frmMail'>" &_
			"<table width='100%'  border='2' cellpadding='5' cellspacing='2' bordercolor='#7B869A'>" &_
			"<tr>" &_
			"	<td valign='top'>" & MsgMailOk &"<table width='100%'  border='2' cellpadding='5' cellspacing='0'>" &_
			"		<tr><td width='7%' align='right' valign='middle'>DE:</td>" &_
			"			<td colspan='4' class='calResaltado'>" & Session("eMailAlumno") & "</td>" &_
			"		</tr>" &_
			"		<tr>" &_
			"			<td align='right' valign='middle'>Para:</td>" &_
			"			<td colspan='4' valign='top'><table width='100%'  border='0' cellpadding='0' cellspacing='5'>" &_
			"											<tr><td valign='middle'>" & JC & "</td></tr>" &_
			"											<tr><td valign='top'><table width='100%'  border='1' cellpadding='0' cellspacing='0'>" &_
			"																	<tr>" &_
			"																		<td align='center' class='button'><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onclick=""expandit('tblMail');"">PROFESORES</a></td>" &_
			"																	</tr>" &_
			"																	<tr><td valign='top' id='tblMail' style='display: none;' name='tblMail'>"
ChkBoxes = ""
for i = Lbound(Ramos) to  UBound(Ramos)
	if Len(Ramos(i))<>0 _
		then
			ThisCadena=Ramos(i)
			Seccion=Left(ThisCadena,InStr(ThisCadena,"-")-1): ThisCadena=Mid(ThisCadena,InStr(ThisCadena,"-")+1)
			CodRamo=Left(ThisCadena,InStr(ThisCadena,"-")-1): ThisCadena=Mid(ThisCadena,InStr(ThisCadena,"-")+1)
			AnoCurse=Left(ThisCadena,InStr(ThisCadena,"-")-1): ThisCadena=Mid(ThisCadena,InStr(ThisCadena,"-")+1)
			Periodo=ThisCadena
			
			Set ra_horprof = dbMat.Execute("select DISTINCT CODPROF from ra_horprof where codramo='" & CodRamo & "' and ano='" & AnoCurse & "' and periodo='" & Periodo & "' and codsecc='" & Seccion & "'")
			if Not ra_horProf.Bof or Not ra_horprof.Eof _
				then
					Set ra_profes = dbMat.Execute("select CODPROF, AP_PATER, AP_MATER, NOMBRES, MAIL from ra_profes where CODPROF='" & ra_horprof("CODPROF") & "' and MAIL is not null")
					if Not ra_profes.Bof or Not ra_profes.Eof _
						then
							if InStr(ra_profes("AP_PATER"),"NO DEFINIDO")=0 _
								then
									if EsEmail(ra_profes("MAIL")) _
										then
											ChkBoxes = ChkBoxes & "<input name='Nomina[]' id='Nomina[]' type='checkbox' value='" & ra_profes("MAIL") & "'>" & CodRamo & " - " & ra_profes("NOMBRES") & " " & ra_profes("AP_PATER") & " " & ra_profes("AP_MATER") & "<br>"
									End If
								else
									ChkBoxes = ChkBoxes & "<!--()//-->"
							end if
						else
							ChkBoxes = ChkBoxes & "<!--()//-->"
					End if			
				else
					ChkBoxes = ChkBoxes & "<!--()//-->"
			End if
		
			Set ra_horprof = Nothing
			Set ra_profes = Nothing
		else			
			ChkBoxes = ChkBoxes & "<!--()//-->"			
	End If
next 

eMail_go = eMail_go & ChkBoxes & "</td></tr></table></td></tr></table></td></tr>" &_
								 "<tr><td align='right' valign='middle'>Asunto: </td><td colspan='4' valign='middle'><input name='Asunto' type='text' id='Asunto' size='70' maxlength='150'></td>" &_
								 "</tr><tr><td align='right' valign='middle'>CONTENIDO:</td><td colspan='4' valign='top'><textarea name='Contenido' cols='50' rows='10' id='Contenido'></textarea></td>" &_
								 "</tr><tr align='center'><td colspan='5' valign='middle'><input name='goGo' type='submit' id='goGo' value='Enviar Correo...'></td>" &_
								 "</tr></table></td></tr></table></form>"
End if 
%>