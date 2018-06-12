<%
    Response.Expires = -1
    Session.Timeout = 999
    Response.Buffer = false
    Server.ScriptTimeout = 99999
%>
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%
dim lc_Pregunta,lc_Respuesta,lc_Alumno,lc_Profesor,lc_TipoProfesor,lc_Ramo,lc_Seccion,lc_Ano, lc_Periodo,lc_Rut,lc_Observacion,j
dim num_Respuestas
dim CantidadPreguntas

lc_Alumno=Session("Codcli")
lc_Profesor=session("CodProf")
lc_TipoProfesor=session("TipoDoc")
lc_Ramo=session("CodRamo")
lc_Seccion=session("CodSecc")
lc_Ano=session("Ano_Ed")
lc_Periodo=session("Periodo_Ed")
lc_Observacion=trim(request("txtObs"))

dim Rt
dim str 
dim tr
'cuenta la cantidad de preguntas y las deja en la variable cantidad
Str="select count(*) as Cantidad from ra_preguntas "
set tr = Session("Conn").execute (str)
CantidadPreguntas=tr("Cantidad")
	'for i = 1 to 25
	  'response.write("C=" & request("COMBO" & i) & "<br>")
'	  response.write("T=" & request("Texto" & i) & "<br>")
	             'lc_Respuesta=trim(request("Texto" & ii))
'	next
'	Response.end
	ii=1
	for ii=1 to int(CantidadPreguntas)
		lc_Pregunta= trim(Request("P" & ii))
		lc_Respuesta=trim(request("Texto" & ii))
			'trim(request("Texto" & ii))
			'response.Write(lc_Respuesta)
			
			if lc_Respuesta <> "" and lc_Respuesta <> 0 and lc_Ramo <> "" then
     		    GrabaRespuesta lc_Pregunta,lc_Respuesta,lc_Alumno,lc_Profesor,lc_TipoProfesor,lc_Ramo,lc_Seccion,lc_Ano,lc_Periodo,lc_Observacion
				'Response.Write("Grabando respuesta    " & ii) & "<br>" & lc_Respuesta & "-" & lc_Pregunta
				Limpiar
			 else
				'Response.Write("No Grabo respuesta    " & ii) & "<br>" & lc_Respuesta & "-" & lc_Pregunta
				
				Limpiar
			end if
	next 
' actualizar estado
' Si total de Respuesta es igual al numero de Preguntas , Actualiza a Profesor Encuestado
		 sql="Select Count(*) from ra_resultados where Ano='" & lc_Ano & "'"
		 sql=sql & " and Periodo='" & lc_Periodo & "' and CodCli='" & lc_Alumno & "'"
		 sql=sql & " and Rut='" & session("Rut") & "' and CodRamo='" & lc_Ramo & "'"
		 sql=sql & " and CodSecc='" & lc_Seccion & "' and CodProf='" & lc_Profesor & "'"
		 sql=sql & " and TipoDocente='" & lc_TipoProfesor & "'"
		'response.Write(sql)
	     set Rt = Session("Conn").execute (sql)
         if not Rt.eof then
            num_Respuestas=valnulo(rt(0),num_)	
			'response.Write(num_Respuestas & "<br>" & TotalRespuesta())					
			if num_Respuestas = TotalRespuesta() then
			    Strsql= " Update RA_ENCUESTADOS set Evaluado='S'"
				Strsql=Strsql & " WHERE  ANO = " & valnulo(lc_Ano,NUM_) & ""
                Strsql=Strsql & " AND  PERIODO = " & valnulo(lc_Periodo,NUM_) & " " 
                Strsql=Strsql & " AND  CODSEC = " & lc_Seccion & "" 
                Strsql=Strsql & " AND  CODRAMO ='" & lc_Ramo & "'"
                Strsql=Strsql & " AND  CODPROF ='" & lc_Profesor & "'"
				Strsql=Strsql & " AND  CODCLI ='" & lc_Alumno & "'"

				set Rt = Session("Conn").execute (strsql)
			end if         
         end if
'response.Write("Paso todo el Ciclo") & "<br>"
%>


<%
Sub GrabaRespuesta(Codpregunta,Codrespuesta,Alumno,Profesor,TipoProfesor,Ramo,Seccion,Ano,Periodo,Observacion)
dim rs
dim sql
'    response.Write("Paso por funcion")
	  'set rs=Server.CreateObject ("ADODB.Recordset")
	  if Codrespuesta <> ""  or Codrespuesta <> "0" then
	     sql="Select * from ra_resultados where Ano='" & Ano & "'"
		 sql=sql & " and Periodo='" & Periodo & "' and CodCli='" & Alumno & "'"
		 sql=sql & " and Rut='" & session("Rut") & "' and CodRamo='" & Ramo & "'"
		 sql=sql & " and CodSecc='" & Seccion & "' and CodProf='" & Profesor & "'"
		 sql=sql & " and TipoDocente='" & TipoProfesor & "' and NroPregunta=" & Codpregunta & ""
		 set rs = Session("Conn").execute (sql)
		 if rs.eof then         'Inserta
               sql="Insert Into ra_Resultados (Ano,Periodo,CodCli,Rut,"
	           sql=sql & "CodRamo,CodSecc,CodProf,TipoDocente,NroPregunta,Respuesta) values("
	           sql=sql & "'" & Ano & "','" & Periodo & "' ,'" & Alumno & "','" & session("Rut") & "',"
	           sql=sql & "'" & Ramo & "','" & Seccion & "' ,'" & Profesor & "',"
 	           sql=sql & "'" & TipoProfesor & "'," & Codpregunta & " ," & valnulo(Codrespuesta,num_) & ")"
	           'response.Write(sql)
	     else                 'Actualiza
		       sql="Update ra_resultados set Respuesta=" & valnulo(Codrespuesta,num_) & " where Ano='" & Ano & "'"
		       sql=sql & " and Periodo='" & Periodo & "' and CodCli='" & Alumno & "'"
		       sql=sql & " and Rut='" & session("Rut") & "' and CodRamo='" & Ramo & "'"
		       sql=sql & " and CodSecc='" & Seccion & "' and CodProf='" & Profesor & "'"
		       sql=sql & " and TipoDocente='" & TipoProfesor & "' and NroPregunta=" & Codpregunta & ""		     
		 end if 		   
		 set rs = Session("Conn").execute (sql)
        end if
		StrSql= "UPDATE RA_Encuestados SET OBSERVACION='" & EliminaPalabras(Observacion) & "'"
		Strsql=Strsql & " WHERE  ANO = " & valnulo(Ano,NUM_) & ""
		Strsql=Strsql & " AND  PERIODO = " & valnulo(Periodo,NUM_)  & " " 
		Strsql=Strsql & " AND  CODSEC = " & Seccion & "" 
		Strsql=Strsql & " AND  CODRAMO ='" & Ramo & "'"
		Strsql=Strsql & " AND  CODPROF ='" & Profesor & "'"
		Strsql=Strsql & " AND CodCli='" & Alumno & "'"
	  
	 set rs=Session("Conn").execute(StrSql)
	'response.Write("Grabò en funcion exitosamente..?")
End Sub
sub Limpiar()
dim lc_Pregunta,lc_Respuesta,lc_Alumno,lc_Profesor,lc_TipoProfesor,lc_Ramo,lc_Seccion,lc_Ano,lc_Periodo,lc_Observacion
end sub
%>
<!--<script>window.top.mainFrame.location.href='ed.asp';</script>-->
<script>window.top.mainFrame.location.href='ed.htm';</script>
