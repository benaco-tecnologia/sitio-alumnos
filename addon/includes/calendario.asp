<%
'pagina que recibirá la fecha seleccionada por el usuario
'Const URLDestino = "OtraPagina.asp" 

Dim ThisEventos
Dim MyMonth 'Month of calendar
Dim MyYear 'Year of calendar
Dim FirstDay 'First day of the month. 1 = Monday
Dim CurrentDay 'Used to print dates in calendar
Dim Col 'Calendar column
Dim Row 'Calendar row
Dim calEventos
dim SqlExec
Dim LosEventos
Dim Conter

MyMonth = Request.Querystring("Month")
MyYear = Request.Querystring("Year")

If IsEmpty(MyMonth) then MyMonth = Month(Date)
if IsEmpty(MyYear) then MyYear = Year(Date)

Call ShowHeader (MyMonth, MyYear)

FirstDay = WeekDay(DateSerial(MyYear, MyMonth, 1)) -1
CurrentDay = 1

'Let's build the calendar
For Row = 0 to 5
	For Col = 0 to 6
		If Row = 0 and Col < FirstDay then
			calEventos=calEventos & "<td>&nbsp;</td>"
		elseif CurrentDay > LastDay(MyMonth, MyYear) then
			calEventos=calEventos & "<td>&nbsp;</td>"
		else
			calEventos=calEventos & "<td"
			'// 
			'// Lee DB con los eventos paracada dia
			'//
			SqlExec="SELECT * FROM CA_Eventos WHERE FechaEvento='"  & CurrentDay & "-" & cInt(MyMonth) & "-" & cInt(MyYear) & "'"
			Set ThisEventos = oCmd.Execute(SqlExec)
			if not ThisEventos.Bof or Not ThisEventos.Eof then		'//			if cInt(MyYear) = Year(Date) and cInt(MyMonth) = Month(Date) and CurrentDay = Day(Date) then '// SI ES LA FECHA ACTUAL ...AKI DEBERIA LEER LA DB 
				calEventos=calEventos & " align='center' onMouseOver=""this.style.backgroundColor='#A4C0DC'"" onMouseOut=""this.style.backgroundColor='#D9DFE5'""><a style='text-align:center; cursor:hand; cursor:pointer; text-decoration:none;' onclick=""expandit('" & CurrentDay & "');""><img src='addon/imagenes/iconsearch.gif' width='20' height='20'></a>"
				LosEventos=LosEventos & "<table width='100%'  border='1' cellpadding='0' cellspacing='0' id='" & CurrentDay & "' style='display: none;' class='poll'><tr><td colspan='2' align='center'>ACTIVADES Y EVENTOS PARA EL DIA " & CurrentDay & "</td></tr>"
				'// LEE DB
				Conter=0
				ThisEventos.MoveFirst
				While Not ThisEventos.Eof
					Conter = Conter + 1
					LosEventos = LosEventos & "<tr><td width='3%' align='center'>" & Conter & "</td><td width='97%'>" & ThisEventos("DetalleEvento") & "</td></tr>"
					ThisEventos.MoveNext
				Wend				
				LosEventos = LosEventos & "</tr></table>"
			else 
				calEventos=calEventos & " align='center'>" 
			end if
			
			if cInt(MyYear) = Year(Date) and cInt(MyMonth) = Month(Date) and CurrentDay = Day(Date) then 
				calEventos=calEventos & "<div>" 
			else
				calEventos=calEventos & "<div>" 
			end if
			calEventos=calEventos & CurrentDay & "</div></td>"			
			CurrentDay = CurrentDay + 1
		End If
	Next
	calEventos=calEventos & "</tr>"
Next
calEventos=calEventos & "</table>" & LosEventos

'------ Sub and functions

Sub ShowHeader(MyMonth,MyYear)
	calEventos="<table border='2' cellspacing='0' cellpadding='15' width='100%' align='center' class='tbox'>" & _
			   "	<tr align='center'>" & _
			   "		<td colspan='7'>" & _ 
			   "			<table border='0' cellspacing='1' cellpadding='1' width='100%'>" & _
			   "				<tr>" & _
			   "					<td align='left'>" & _
			   "						<a href = 'main.asp?caID=4&"
	
	if MyMonth - 1 = 0 then
		calEventos=calEventos & "month=12&year=" & MyYear - 1
		else
			calEventos=calEventos & "month=" & MyMonth - 1 & "&year=" & MyYear
	end if
	
	calEventos=calEventos & "'><img src='addon/imagenes/iconarrowleft.gif' width='8' height='13' border='0'><img src='addon/imagenes/iconarrowleft.gif' width='8' height='13' border='0'></a>" & _
							"<span class='calEncabe'><b>&nbsp;&nbsp;" & Ucase(MonthName(MyMonth)) & "&nbsp;&nbsp;</b></span>" & _
							"<a href = 'main.asp?caID=4&"
	
	if MyMonth + 1 = 13 then 
		calEventos=calEventos & "month=1&year=" & MyYear + 1
		else
			calEventos=calEventos & "month=" & MyMonth + 1 & "&year=" & MyYear
	end if
	
	calEventos=calEventos & "'><img src='addon/imagenes/iconarrowright.gif' width='8' height='13' border='0'><img src='addon/imagenes/iconarrowright.gif' width='8' height='13' border='0'></a>" & _
							"</td><td align='center'><div><b>ENAC - CARITAS CHILE</b></div>" & _
							"</td><td align='right'>"
	
	calEventos=calEventos & "<a href = 'main.asp?caID=4&" & _
							"month=" & MyMonth & "&year=" & MyYear -1 & _
							"'><img src='addon/imagenes/iconarrowleft.gif' width='8' height='13' border='0'><img src='addon/imagenes/iconarrowleft.gif' width='8' height='13' border='0'></a>"
	
	calEventos=calEventos & "<span class='calEncabe'><b>&nbsp;&nbsp;" & MyYear & "&nbsp;&nbsp;</b></span>" & _
							"<a href = 'main.asp?caID=4&" & _
							"month=" & MyMonth & "&year=" & MyYear + 1 & _
							"'><span class='calSimbolo'><img src='addon/imagenes/iconarrowright.gif' width='8' height='13' border='0'><img src='addon/imagenes/iconarrowright.gif' width='8' height='13' border='0'></span></a>" & _
							"</td></tr></table></td></tr><tr align='center'>" & _
							"<td><div class='calDias'>Domingo</div></td>" & _
							"<td><div class='calDias'>Lunes</div></td>" & _
							"<td><div class='calDias'>Martes</div></td>" & _
							"<td><div class='calDias'>Miercoles</div></td>" & _
							"<td><div class='calDias'>Jueves</div></td>" & _
							"<td><div class='calDias'>Viernes</div></td>" & _
							"<td><div class='calDias'>Sabado</div></td>" & _
							"</tr>"
End Sub

Function MonthName(MyMonth)
	Select Case MyMonth
		Case 1
			MonthName = "Enero"
		Case 2
			MonthName = "Febrero"
		Case 3
			MonthName = "Marzo"
		Case 4
			MonthName = "Abril"
		Case 5
			MonthName = "Mayo"
		Case 6
			MonthName = "Junio"
		Case 7
			MonthName = "Julio"
		Case 8
			MonthName = "Agosto"
		Case 9
			MonthName = "Septiembre"
		Case 10
			MonthName = "Octubre"
		Case 11
			MonthName = "Noviembre"
		Case 12
			MonthName = "Diciembre"
		Case Else
			MonthName = "ERROR!"
	End Select
End Function

Function LastDay(MyMonth, MyYear)
' Returns the last day of the month. Takes into account leap years
' Usage: LastDay(Month, Year)
' Example: LastDay(12,2000) or LastDay(12) or Lastday

	Select Case MyMonth
		Case 1, 3, 5, 7, 8, 10, 12
			LastDay = 31

		Case 4, 6, 9, 11
			LastDay = 30

		Case 2
			If IsDate(MyYear & "-" & MyMonth & "-" & "29") Then LastDay = 29 Else LastDay = 28

		Case Else
			LastDay = 0
	End Select
End Function
%>