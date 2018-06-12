<%
    Response.Expires = -1
    Session.Timeout = 999
    Response.Buffer = false
    Server.ScriptTimeout = 99999
%>
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="include/guardalog.inc" -->
<%
dim lc_Pregunta,lc_Respuesta,lc_Alumno,lc_Profesor,lc_TipoProfesor,lc_Ramo,lc_Seccion,lc_Ano, lc_Periodo,lc_Rut,lc_Observacion,j,lc_encuesta
dim num_Respuestas
dim CantidadPreguntas

lc_Alumno=Session("Codcli")
lc_Profesor=session("CodProf")
lc_TipoProfesor=session("TipoDoc")
lc_Ramo=session("CodRamo")
lc_Seccion=session("CodSecc")
lc_Ano=session("Ano_Ed")
'JJJ lc_Periodo=session("Periodo_Ed")
lc_Periodo=session("Periodo_pas")


lc_Observacion=trim(request("txtObs"))

dim Rt
dim str 
dim tr
'Str="select count(*) as Cantidad from ra_preguntas "
Str="SELECT COUNT(*) as Cantidad FROM RA_ENCUESTA_DETALLE WHERE NRO_ENCUESTA = '" & Session("NroEncuesta") & "' "

set tr = Session("Conn").execute (str)
CantidadPreguntas=tr("Cantidad")
	'for i = 1 to 100
	'  response.write("C=" & request("COMBO" & i) & "<br>")
	'  response.write("T=" & request("Texto" & i) & "<br>")
	'next
	'Response.end
	ii=1
	for ii=1 to int(CantidadPreguntas)
		lc_Pregunta= trim(Request("P" & ii))
		lc_Respuesta=trim(request("Combo" & ii))
		lc_encuesta = trim(request("Enc" & ii))
		
	    'response.Write("lc_Pregunta")
		'response.Write(lc_Pregunta)
        'response.Write("lc_Respuesta")
		'response.Write(lc_Respuesta)
		'response.Write("lc_encuesta")
		'response.Write(lc_encuesta)
		'response.End()
		
			'trim(request("Texto" & ii))
			if lc_Respuesta <> "" and lc_Ramo <> "" then
			'response.Write(Session("NroEncuesta"))
			'response.write(lc_encuesta)
			'response.End()
     		    GrabaRespuesta lc_Pregunta,lc_Respuesta,lc_Alumno,lc_Profesor,lc_TipoProfesor,lc_Ramo,lc_Seccion,lc_Ano,lc_Periodo,lc_Observacion,lc_encuesta
				'Response.Write("Grabando respuesta    " & ii) & "-" & lc_Respuesta & "-" & lc_Pregunta & "<br>"
				Limpiar
			 else
				 'Response.Write("No Grabo respuesta    " & ii) & "<br>" & lc_Respuesta & "-" & lc_Pregunta
				Limpiar
			end if
	next 
	
'response.End()
' actualizar estado
' Si total de Respuesta es igual al numero de Preguntas , Actualiza a Profesor Encuestado
		 sql=""
		 sql="Select Count(*) from ra_resultados where Ano='" & lc_Ano & "'"
		 sql=sql & " and Periodo='" & lc_Periodo & "' and CodCli='" & lc_Alumno & "'"
		 sql=sql & " and Rut='" & session("Rut") & "' and CodRamo='" & lc_Ramo & "'"
		 sql=sql & " and CodSecc='" & lc_Seccion & "' and CodProf='" & lc_Profesor & "'"
		 sql=sql & " and TipoDocente='" & lc_TipoProfesor & "'"
		 sql=sql & " AND  encuesta ='" & Session("NroEncuesta") & "'"
	    'response.Write(sql)
		'response.End
		'response.Write(sql)
		'response.End
	     set Rt = Session("Conn").execute (sql)
         if not Rt.eof then
            num_Respuestas=valnulo(rt(0),num_)	
			'response.Write(num_Respuestas & "<br>" & TotalRespuesta())					
			if num_Respuestas = TotalRespuestasEncuesta(Session("NroEncuesta")) then
			    Strsql= " Update RA_ENCUESTADOS set Evaluado='SI'"
				Strsql=Strsql & " WHERE  ANO = " & valnulo(lc_Ano,NUM_) & ""
                Strsql=Strsql & " AND  PERIODO = " & valnulo(lc_Periodo,NUM_) & " " 
                Strsql=Strsql & " AND  CODSECC = " & lc_Seccion & "" 
                Strsql=Strsql & " AND  CODRAMO ='" & lc_Ramo & "'"
                Strsql=Strsql & " AND  CODPROF ='" & lc_Profesor & "'"
				Strsql=Strsql & " AND  CODCLI ='" & lc_Alumno & "'"
				'Verificar que este correcto lo de abajo
				Strsql=Strsql & " AND  encuesta ='" & Session("NroEncuesta") & "'"
				
				set Rt = Session("Conn").execute (strsql)
				l=guarda_log(Session("id_usuario"),Session("CodCli"),475,"Graba encuesta Docente Ramo:"&lc_Ramo&", Seccion:"&lc_Seccion &", Año y Periodo:"&lc_Periodo&"/"&lc_Ano &", Profesor:"&lc_Profesor&"")
			end if         
         end if
'response.Write("Paso todo el Ciclo") & "<br>"
%>


<%
Sub GrabaRespuesta(Codpregunta,Codrespuesta,Alumno,Profesor,TipoProfesor,Ramo,Seccion,Ano,Periodo,Observacion,encuesta)
dim rs
dim sql
'    response.Write("Paso por funcion")
	  'set rs=Server.CreateObject ("ADODB.Recordset")
	  if Codrespuesta <> ""  or Codrespuesta <> "0" then
	
		 textolibre = EsTextoLibre(Codpregunta)
	  
	     sql="Select * from ra_resultados where Ano='" & Ano & "'"
		 sql=sql & " and Periodo='" & Periodo & "' and CodCli='" & Alumno & "'"
		 sql=sql & " and Rut='" & session("Rut") & "' and CodRamo='" & Ramo & "'"
		 sql=sql & " and CodSecc='" & Seccion & "' and CodProf='" & Profesor & "'"
		 sql=sql & " and TipoDocente='" & TipoProfesor & "' and NroPregunta=" & Codpregunta & " "
		 sql=sql & " AND  encuesta ='" & Session("NroEncuesta") & "'"

		 set rs = Session("Conn").execute (sql)
		 if rs.eof then         'Inserta
               			sql = ""
			sql = "Insert Into RA_RESULTADOS " & _ 
			    " (Ano, Periodo, CodCli, Rut, Codramo, " &_ 
			    " CodSecc, CodProf, TipoDocente, NroPregunta, Respuesta, " & _ 
			    " Encuesta, Fecha,textolibre) "
			sql = sql & "Values('" & Ano & "', "
			sql = sql & "	    '" & Periodo & "', "
			sql = sql & "		'" & Alumno & "', "
			sql = sql & "		'" & Session("Rut") & "', "
			sql = sql & "		'" & Ramo & "', "			
			sql = sql & "		" & Seccion & ", "
			sql = sql & "		'" & Profesor & "', "
			sql = sql & "		'" & TipoProfesor & "', "		
			sql = sql & "		'" & Codpregunta & "', "
			
			if textolibre = "SI" then
				sql = sql & "		NULL, "
			else
				sql = sql & "		'" & Codrespuesta & "', "
			end if 
			sql = sql & "		'" & encuesta & "', getdate() "	 
			if textolibre = "SI" then
				sql = sql & ",'"& Codrespuesta &"' )"	   
			else
				sql = sql & ",NULL )"	   
			end if 
			   	'	 response.Write(sql)
'		 response.End()
			   'sql="Insert Into ra_Resultados (Ano,Periodo,CodCli,Rut,"
	           'sql=sql & "CodRamo,CodSecc,CodProf,TipoDocente,NroPregunta,Respuesta) values("
	           'sql=sql & "'" & Ano & "','" & Periodo & "' ,'" & Alumno & "','" & session("Rut") & "',"
	           'sql=sql & "'" & Ramo & "','" & Seccion & "' ,'" & Profesor & "',"
 	           'sql=sql & "'" & TipoProfesor & "'," & Codpregunta & " ," & valnulo(Codrespuesta,num_) & ")"
	           'response.Write(sql)
	     else                 'Actualiza
		 	   if textolibre = "SI" then
			   		sql="Update ra_resultados set textolibre='" & valnulo(Codrespuesta,str_) & "'"
			   else
			   		sql="Update ra_resultados set Respuesta=" & valnulo(Codrespuesta,num_) & ""
			   end if  
		       sql=sql & " where Ano='" & Ano & "' and Periodo='" & Periodo & "' and CodCli='" & Alumno & "'"
		       sql=sql & " and Rut='" & session("Rut") & "' and CodRamo='" & Ramo & "'"
		       sql=sql & " and CodSecc='" & Seccion & "' and CodProf='" & Profesor & "'"
		       sql=sql & " and TipoDocente='" & TipoProfesor & "' and NroPregunta=" & Codpregunta & ""		     
		 end if 		   
		 

		set rs = Session("Conn").execute (sql)
		'response.Write(sql)
		
        end if
		StrSql= "UPDATE RA_Encuestados SET OBSERVACION=LEFT('" & EliminaPalabras(Observacion) & "',500)"
		Strsql=Strsql & " WHERE  ANO = " & valnulo(Ano,NUM_) & ""
		Strsql=Strsql & " AND  PERIODO = " & valnulo(Periodo,NUM_)  & " " 
		Strsql=Strsql & " AND  CODSECC = " & Seccion & "" 
		Strsql=Strsql & " AND  CODRAMO ='" & Ramo & "'"
		Strsql=Strsql & " AND  CODPROF ='" & Profesor & "'"
		Strsql=Strsql & " AND CodCli='" & Alumno & "'"
	  
	 set rs=Session("Conn").execute(StrSql)
	'response.Write("Grabò en funcion exitosamente..?")
	
End Sub
sub Limpiar()
dim lc_Pregunta,lc_Respuesta,lc_Alumno,lc_Profesor,lc_TipoProfesor,lc_Ramo,lc_Seccion,lc_Ano,lc_Periodo,lc_Observacion,lc_encuesta
end sub

function EsTextoLibre(codpregunta)
Dim rstResp, sql
	EsTextoLibre = ""
	
	sql= "select coalesce(tr.TextoLibre,'NO')TextoLibre  "
	sql=sql & " from RA_PREGUNTAS p,RA_TIPORESPUESTA TR "
	sql=sql & " where tr.CODTIPORESPUESTA = p.CODRESPUESTA "
	sql=sql & " and p.numero=" & codpregunta & "" 
    
   	set rstResp = Session("Conn").execute (sql)
    if not rstResp.eof then
       EsTextoLibre = rstResp("TextoLibre")
    end if	  
end function 

%>
<!--<script>window.top.mainFrame.location.href='ed.asp';</script>-->
<script>window.top.mainFrame.location.href='ed.htm';</script>
<!--#INCLUDE file="include/desconexion.inc" -->
