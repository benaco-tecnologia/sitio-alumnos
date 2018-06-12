<% 
   Response.Expires = -1 
%>
<!--#INCLUDE file="include/conexion.inc" -->
<!--#INCLUDE file="include/funciones.inc" -->
<!--#INCLUDE file="include/periodos.inc" -->
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=Login name=Login></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=LoginAux name=LoginAux></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=LoginB name=LoginB></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=LoginPeriodo name=LoginPeriodo></OBJECT>
<%
'response.Write( Session("MiValor"))
'response.End()
Dim IdRut, IdRutAux, Clave, ClaveEncript, ClaveDesenc, strSql, strSqlAux, CodCli,SW,strb,IdRutB,sTRSQLP,Estado,RX
Dim msg,Bloqueo,Permiso,opcion1,EstadoMatriculado, mySQL,myRST,Carr,MatriculaSemestral,StrPreValida,StrCar
dim GetCar,CarreraAlumno, Sstr,Rrst, Rs,rstParame,mystr,claveNW,usrNW,sindig,perfilNW
Opcion1="INT_001"


'Session("MiValor")=0
if Session("MiValor")=0 then ' o cuando es vacío también pasa
'response.Write("entro 1")
	if Session("FromNet") = 1 then		
		IdRut = RutSinDV(EliminaPalabras(Trim(Session("RQRut"))))
		Session("RutCli")=idRut
		usrNW = Session("RQRut")
		Session("usrNW") = usrNW	
		Session("RutCliente")=(Trim(Session("RQRut")))
		ClaveEncript = Session("RQpass")
		Session("MiClave")=(Session("RQpass"))
		Session("ClaveEncriptada")=ClaveEncript	
		claveNW = Session("MiClave")
		usrNW = Replace(usrNW, ".", "")
		usrNW = Replace(usrNW, "-", "")	
		Clave=Session("RQpass")							
	else
		IdRut = RutSinDV(EliminaPalabras(Trim(Request("logrut"))))
		Session("RutCli")=idRut
		usrNW = Request("logrut")
		Session("usrNW") = usrNW	
		Session("RutCliente")=(Trim(Request("logrut")))
		ClaveEncript = Request("logclave")
		Session("MiClave")=(Request("logclave"))
		Session("ClaveEncriptada")=ClaveEncript	
		claveNW = Session("MiClave")
		usrNW = Replace(usrNW, ".", "")
		usrNW = Replace(usrNW, "-", "")	
		Clave=Request("logclave")
	end if 	
	Session("RutClienteR")=Session("RutCliente")
	Session("RutClienteR")= Replace(Session("RutClienteR"), ".", "")
	Session("RutClienteR")= Replace(Session("RutClienteR"), "-", "")		
	

	StrPreValida = " SELECT * FROM ca_usuarios WHERE us_consuser = '" & usrNW & "'"
	Set rstU = Session("Conn").execute(StrPreValida)
    if rstU.eof then		
		'Response.Redirect("MensajeNoExiste.asp")
		session("MensajeLogin")="Lo siento, el Rut que ha Ingresado no se encuentra en nuestros registros....!!!\nVuelva a Intentar"
		Response.Redirect "alumnosfinning.asp" 
		Response.End()
	end if
	
	StrPreValida = " SELECT * FROM ca_usuarios WHERE us_consuser = '" & usrNW & "'  AND dbo.Encrypt('" & Clave & "' ) = us_password "	
	Set rstV = Session("Conn").execute(StrPreValida)
    if rstV.eof then
	 	'Response.Redirect "ClaveIncorrecta.asp" 
		session("MensajeLogin")="Clave incorrecta ..!!! Vuelva a Intentar\n(Recuerde que el campo password es sensible a mayúsculas y minúsculas)"
		Response.Redirect "alumnosfinning.asp" 
	  	Response.End()	
	end if

		
Session("MiValor")=1
end if
'codigo a prueba

Clave = Trim(Request("logclave"))
	if Session("MiValor")=1 then
	

 		
 
		if GetVariasCarreras(idRut)="Verdadero" and session("CarreraAlumno") = "" then
		   Session("GetCarr")=1
		   session("RutAlum")=Trim(Request("logrut"))
		   Response.Redirect "EligeCarrera.asp"
		   response.End
		else			
		   Session("GetCarr")=0
		   
		end if
		Session("MiValor")=2
    end if

'if Session("GetCarr") = 0 then

'			StrCar = " SELECT codcarpr FROM mt_alumno WHERE rut = '" & Session("RutCli") & "' "	
'			if bcl_ado(StrCar,rst) then
'			session("CarreraAlumno")=Valnulo(rst(0), Str_)
'			end if
'end if


'response.Write(session("CarreraAlumno"))
'response.End()

'	response.write("tiene mas de una carrera")
'	response.end
		   
	 GetCar= Session("GetCarr")
	 CarreraAlumno=session("CarreraAlumno")
	 IdRut=Session("RutCli")
	usrNW =Session("usrNW") 
	claveNW = Session("MiClave")
	'Session("RutAlum") = Session("RutCli")
	 'Response.Write(CarreraAlumno)
	
	 strSql ="select matriculasemestral, coalesce(websituactual, 'N') as websituactual from mt_parame"
	 strSql ="SELECT ("&_
"SELECT Valor FROM MT_PARAME_DET WHERE IdParametro = 'matriculasemestral' AND   GETDATE() BETWEEN FechaInicio AND FechaTermino) AS matriculasemestral,"&_
"(SELECT  coalesce(Valor, 'N') as websituactual FROM MT_PARAME_DET WHERE IdParametro = 'websituactual' AND   GETDATE() BETWEEN FechaInicio AND FechaTermino) AS websituactual"

	 'strSql ="select matriculasemestral from mt_parame"
	'response.Write(strsql)
	'response.End()
		if bcl_ado(strsql,rstParame) then
			MatriculaSemestral = trim(valnulo(rstParame("matriculasemestral"),str_))
			SituacionActual = trim(valnulo(rstParame("websituactual"),str_))
			rstParame.close
		end if

	Session("SituacionActual")= SituacionActual
	'Session("SituacionActual")= "N"

'	if CarreraAlumno="" then
'		strSql = "Select top 1 a.rut, c.nombre as nom, c.paterno as app, a.passweb, a.passchanged, b.codcli, b.codsede,b.codcarpr,b.codpestud, 'b.ano_mat, b.ano, b.jornada,b.codcarpr,b.periodo_mat , b.nivel" &_
'					 " From mt_webmail a with (nolock), mt_alumno b with (nolock), mt_client c with (nolock)" &_ 
'					 " Where a.rut = '" & IdRut & "' And a.rut = b.rut  and c.codcli = a.rut order by b.ano desc"
'	else
'  		strSql = "Select top 1 a.rut, c.nombre as nom, c.paterno as app, a.passweb, a.passchanged, b.codcli, b.codsede,b.codcarpr,b.codpestud, 'b.ano_mat, b.ano, b.jornada,b.codcarpr,b.periodo_mat, b.nivel " &_
'					 " From mt_webmail a with (nolock), mt_alumno b with (nolock), mt_client c with (nolock)" &_ 
'					 " Where a.rut = '" & IdRut & "' and b.codcarpr= '"& CarreraAlumno &"' And a.rut = b.rut  and c.codcli = a.rut " &_ 
'					 " and b.estacad = '" & trim(Session("EstadoCarrera")) &"' order by b.ano desc"
'	end if

usrNW = Replace(usrNW, ".", "")
usrNW = Replace(usrNW, "-", "")	
Session("usrNW") = usrNW
	
'response.Write(usrNW & Trim(Request("logrut")))
'response.End()


'Session("usrNW")
'	response.Write("mi valor lol")
'	response.Write(Session("MiValor"))
'	response.End()


	if CarreraAlumno="" then
		strSql = "Select top 1 b.rut, c.nombre as nom, c.paterno as app, a.us_password, '' as passchanged, b.codcli, b.codsede,b.codcarpr,b.codpestud, " &_
"b.ano_mat, b.ano, b.jornada,b.codcarpr,b.periodo_mat, b.nivel,a.id_usuario,c.mail_inst " &_
"From ca_usuarios a with (nolock), mt_alumno b with (nolock), mt_client c with (nolock) ,tg_personas d " &_
"Where a.us_consuser = '" & usrNW & "'  And dbo.Fn_RutSinDig(d.pe_rut) = b.rut  and c.codcli = b.rut AND a.id_persona =d.id_persona " &_
"order by b.ano DESC"
else
	strSql = "Select top 1 b.rut, c.nombre as nom, c.paterno as app, a.us_password, '' as passchanged, b.codcli, b.codsede,b.codcarpr,b.codpestud, " &_
"b.ano_mat, b.ano, b.jornada,b.codcarpr,b.periodo_mat, b.nivel,a.id_usuario,c.mail_inst " &_
"From ca_usuarios a with (nolock), mt_alumno b with (nolock), mt_client c with (nolock) ,tg_personas d " &_
"Where a.us_consuser = '" & usrNW & "' and b.codcarpr= '"& CarreraAlumno &"' And dbo.Fn_RutSinDig(d.pe_rut) = b.rut  and c.codcli = b.rut" &_
" and a.id_persona = d.id_persona and b.estacad = '" & trim(Session("EstadoCarrera")) &"' order by b.ano DESC"

end if



'SEGUIR AQUI !!!	
'response.Write(strsql)
'response.End

'	response.write("tiene mas de una carrera")
'	response.end

Dim claveD
Dim encriptada
claveNW = Session("MiClave")
'response.Write(claveNW)
'response.End()
mystr = "select dbo.Encrypt('"&claveNW&"') as clave"
Set rst = Session("Conn").execute(mystr)
encriptada = Rst("clave")

'response.write(strSql)
'response.end()
Login.Open strSql,Session("Conn")

	
'response.write(Login("us_password"))
'response.end()

if Login.eof and login.bof then
	Session("MiValor")=0
	'Response.Redirect("MensajeNoExiste.asp")
	session("MensajeLogin")="Lo siento, el Rut que ha Ingresado no se encuentra en nuestros registros....!!!\nVuelva a Intentar"
	Response.Redirect "alumnosfinning.asp" 
	Response.End()
end if

Carr=login("codcarpr")
Session("id_usuario") = login("id_usuario")

'response.write(login("codcarpr") + login("us_password")) 
'response.End()

'response.Write(MatriculaSemestral) 
'response.End()
IdRut = Login("rut")
if CarreraAlumno ="" then
	estado=GetEstado2(IdRut)
else
	estado=GetEstado(IdRut,CarreraAlumno)
end if

'response.Write(estado) & "</br>"

session("estado")=estado
ClaveEncript=Session("ClaveEncriptada")


'response.Write("perfil:" & perfilNW &"claveencriptada"&claveencript&" y el estado "&estado)
'response.end

	  
If Login.Eof Then
	  IdRutAux = Session("usrNW")' Trim(Request("logrut"))
	  Session("RutAlum") = IdRutAux
	  'Login.Close()
	  'Session("Conn").Close()
	  'Session("Logueado") = "0"
	  Session("Logueado") = "1"
	  session("EstadoMatriculado")="NO MATRICULADO"	
	  'Response.Redirect("NoMatriculado.asp")
	  'response.End()   
	  'Secomentò esto el 21 octubre del 2003
	  

	  
Else ''' ********** VERIFICAR QUE EL PRIMER LOGIN SEA CON NUMERO DE MATRICULA

session("EstadoMatriculado")="MATRICULADO"
GetPeriodoEd  IdRut , Carr
GetBuscaRamosInasistentes  IdRut , Carr

GetHabilitaBotonEncuenta IdRut, Carr 






	'strSql ="select b.ano,b.periodo,a.NumMaxPer from mt_alumno a ,mt_carreraperiodo b "
	'strSql= strSql & " where a.codcarpr= b.codcarr  "
	'strSql= strSql & " and a.codpestud=b.codpestud "
	'strSql= strSql & " and a.rut='" & IdRut & "'"
	''response.write(strSql) & "</br>"
	''response.End
	'Set Login2 = Session("Conn").execute(strSql)
	'if login2.eof then
	'   Response.Redirect "MensajesNoapertura.asp"
	'   response.End
	'end if
	'Session("ano") = Login2("Ano")
	'Session("Periodo")= Login2("Periodo")                                                                                                        

NumMaxPer = session("NumeroMaximoPrueba")

'COMENTADO 2-8-2011 POR QUE NO EXISTE TMPTEST
'ResetInscrip CodCli, Session("ano"), Session("Periodo")




'Login2.Close()
'
'mySQL="Select num_max_periodo from mt_carrer where codcarr='"& Carr &"'"
'   Set myRST= Session("Conn").execute(mySQL)
'	if myRST("num_max_periodo") <> 1 then
'		if NumMaxPer = 1 then	    
'				  Session("Ano_Ed") = cdbl(Session("ano")) - 1
'				  Session("Periodo_Ed")=NumMaxPer 
'		 elseif NumMaxPer > 1 then 
'				 if cdbl(Session("periodo"))=1 then
'					Session("Ano_Ed") = cdbl(Session("ano")) - 1
'					Session("Periodo_Ed")=NumMaxPer 'prueba..
'				 else
'					Session("Periodo_Ed") = cdbl(Session("Periodo")) - 1
'					Session("Ano_Ed") = cdbl(Session("ano"))
'				 end if	
'		 end if
'		 			session("NumeroMaximoPrueba")=NumMaxPer	
'	else
'	 				Session("Ano_Ed") = cdbl(Session("ano")) - 1
'					Session("Periodo_Ed")=myRST("num_max_periodo") 'prueba..				
'		 			session("NumeroMaximoPrueba")=myRST("num_max_periodo")
'	end if
'response.Write(NumMaxPer) & "</br>" 
'response.Write("NumeroMaximo de Periodo") & "</br>"
'response.Write(Session("Periodo_Ed"))& "</br>"
'response.Write(Session("Ano_Ed")) & "</br>"
'response.Write(Session("Ano")) & "</br>"
'response.Write(Login("ano_mat")) & "</br>"
'response.end 
   
   'if Trim(Login("ano_mat")) <> trim(Session("Ano")) then
    '  	    'Login.Close()
'		    'Session("Conn").Close() 
	'		session("SW")=1
	'		Session("NomAlum") = Trim(Login(1)) & " " & Trim(Login(2))
	'	  	IdRutAux = Trim(Request("logrut"))
	'	  	Session("RutAlum") = IdRutAux
	'	  	Session("Rut") = IdRut
	''	  	Session("CodCli") = Login("CodCli")
		'  	session("Codsede")=Login("Codsede")
		  	'session("Codcarr")=login("Codcarpr")
		 ' 	Session("codpestud")=login("Codpestud")
          '	Session("AnoAlumno") = login("ano")
           ' Session("Jornada") = login("Jornada")
		    'Response.Redirect "atencion.asp?Error=4&Logueo=1" ' No Matriculado
		    'Comente esto el 21 de Octubre del 2003
     		'Session("Logueado") = "0"
			'Session("Logueado") = "1"
			'session("EstadoMatriculado")="N"
			'session("m_bloque")="<blockquote><div align='justify'>Conforme a nuestros registros, a la fecha Usted, presenta a la fecha estado No Matriculado.<br><br>Para regularizar esta situación debe concurrir a las oficinas de Registro y Control de sus respectivas Sedes.<br><br>Si a la fecha usted ya regularizó su situación, rogamos dar aviso a la dirección de biblioteca  de la Universidad.</div></blockquote>"
		    ''Response.End	             
'    end if


	strMatriculado="Pr_Valida_Matricula_Alumno  '" & Carr & "','" & Session("RutCli") & "'"
	'response.Write(strMatriculado)
	'response.End()
	
	if bcl_ado(strMatriculado,rstMatriculado) then
		if valnulo(rstMatriculado(0),STR_)="SI" then
			session("EstadoMatriculado")="MATRICULADO"
		else
			Session("EstadoMatriculado")="NO MATRICULADO"		
		end if 
	else
		Session("EstadoMatriculado")="NO MATRICULADO"		
	END IF 
	
	session("SW")=1
	Session("NomAlum") = Trim(Login(1)) & " " & Trim(Login(2))
	IdRutAux = Session("usrNW")' Trim(Request("logrut"))
	Session("RutAlum") = IdRutAux
	Session("Rut") = IdRut
	Session("CodCli") = Login("CodCli")
	Session("mail_inst") = Login("mail_inst")
	session("Codsede")=Login("Codsede")
	session("Codcarr")=login("Codcarpr")
	Session("codpestud")=login("Codpestud")
	Session("AnoAlumno") = login("ano")
	Session("Jornada") = login("Jornada")
	Session("Nivel") = login("Nivel")
	
	'Response.Redirect "atencion.asp?Error=4&Logueo=1" ' No Matriculado
	'Comente esto el 21 de Octubre del 2003
	'Session("Logueado") = "0"
	Session("Logueado") = "1"

	if Session("EstadoMatriculado")="NO MATRICULADO" then
		session("m_bloque")="<blockquote><div align='justify'>Conforme a nuestros registros, a la fecha usted, presenta estado No Matriculado.<br><br>Para regularizar esta situación debe concurrir a las oficinas de Registro Curricular.<br><br></div></blockquote>"
	end if		



'	strsqlAD="select dbo.Fn_ValorParame('ANOADMISION')AS ano,dbo.Fn_ValorParame('PERIODOADMISION') as periodo "
'	response.Write(strsqlAD)	
'	if bcl_ado(strsqlAD,rstParameAD) then
'		ANOAD = valnulo(rstParameAD("ano"),NUM_)
'		PERIODOAD = valnulo(rstParameAD("periodo"),NUM_)
'	END IF 
'	
'	'response.Write(	ANOAD & "  " &PERIODOAD)
'	'response.End()
'	if trim(MatriculaSemestral)<>"S" then
'	
'	   'response.Write("Matricula Anual")
'	   if Trim(Login("ano_mat")) <> trim(ANOAD) then
'			'Login.Close()
'			'Session("Conn").Close() 
'			session("SW")=1
'			Session("NomAlum") = Trim(Login(1)) & " " & Trim(Login(2))
'			IdRutAux = Session("usrNW")' Trim(Request("logrut"))
'			Session("RutAlum") = IdRutAux
'			Session("Rut") = IdRut
'			Session("CodCli") = Login("CodCli")
'			Session("mail_inst") = Login("mail_inst")
'			session("Codsede")=Login("Codsede")
'			session("Codcarr")=login("Codcarpr")
'			Session("codpestud")=login("Codpestud")
'			Session("AnoAlumno") = login("ano")
'			Session("Jornada") = login("Jornada")
'			Session("Nivel") = login("Nivel")
'			'Response.Redirect "atencion.asp?Error=4&Logueo=1" ' No Matriculado
'			'Comente esto el 21 de Octubre del 2003
'			'Session("Logueado") = "0"
'			Session("Logueado") = "1"
'			
'			session("EstadoMatriculado")="NO MATRICULADO"
'			if Trim(Login("ano_mat")) > trim(ANOAD) then
'				session("EstadoMatriculado")="MATRICULADO"
'			end if
'			
'			session("m_bloque")="<blockquote><div align='justify'>Conforme a nuestros registros, a la fecha usted, presenta estado No Matriculado.<br><br>Para regularizar esta situación debe concurrir a las oficinas de Registro Curricular de sus respectivas Escuelas.<br><br></div></blockquote>"
'			'Response.End	             
'		end if
'	else
'	   'response.Write("Matricula semestral")  & "</br>"
'	   if Trim(Login("ano_mat")) <> trim(ANOAD) or Trim(Login("periodo_mat"))<> trim(PERIODOAD) then
'	   'response.Write("No está matriculado")		
'			'Login.Close()
'			'Session("Conn").Close() 
'			session("SW")=1
'			Session("NomAlum") = Trim(Login(1)) & " " & Trim(Login(2))
'			IdRutAux = Session("usrNW")' Trim(Request("logrut"))
'			Session("RutAlum") = IdRutAux
'			Session("Rut") = IdRut
'			Session("CodCli") = Login("CodCli")
'			Session("mail_inst") = Login("mail_inst")
'			session("Codsede")=Login("Codsede")
'			session("Codcarr")=login("Codcarpr")
'			Session("codpestud")=login("Codpestud")
'			Session("AnoAlumno") = login("ano")
'			Session("Jornada") = login("Jornada")
'			Session("Nivel") = login("Nivel")
'			'Response.Redirect "atencion.asp?Error=4&Logueo=1" ' No Matriculado
'			'Comente esto el 21 de Octubre del 2003
'			'Session("Logueado") = "0"
'			Session("Logueado") = "1"
'			
'			session("EstadoMatriculado")="NO MATRICULADO"
'			if Trim(Login("ano_mat")) > trim(ANOAD) then
'  			session("EstadoMatriculado")="MATRICULADO"
'		 end if
'			 if Trim(Login("ano_mat")) = trim(ANOAD) and Trim(Login("periodo_mat"))> trim(PERIODOAD) then
'			 session("EstadoMatriculado")="MATRICULADO"
'			 end if
'			 
'			session("m_bloque")="<blockquote><div align='justify'>Conforme a nuestros registros, a la fecha usted, presenta estado No Matriculado.<br><br>Para regularizar esta situación debe concurrir a las oficinas de Registro Curriculas de sus respectivas Escuelas.<br><br></div></blockquote>"
'			'Response.End	       
'		end if
'	end if
	'response.End()
	'Analisis de Biblioteca..	
	
'response.Write("antes biblioteca")
'response.End()

if BloqueoBibliotecaCarrera(Login("CodCarpr"))="S" then
		if BloqueoBibliotecaAlumno(IdRut)="S" then
			''JJJ
			session("m_bloque")="<blockquote><div align='justify'>Conforme a nuestros registros, a la fecha usted, presenta deuda de Biblioteca.<br><br>Para regularizar esta situación debe concurrir a la Biblioteca Institucional.<br><br></div></blockquote>"
			session("BloqueoB")="S"
			Session("Logueado") = "1"
		else
			session("BloqueoB")="N"
			Session("Logueado") = "1"
		end if
else
'response.Write("No Tiene para Bloqueos")
   	session("BloqueoB")="N"
	Session("Logueado") = "1"
end if	
'BLOQUEO BLIBLITECA .NET LISTO




'response.Write(session("BloqueoB"))	& "</br>"
'response.Write(session("Logueado"))	& "</br>"
'response.End
'Analisis de mt_bloqueos
'response.Write("")
'response.End()



'response.Write("1er Punto de Interrupcion")
'response.End()




	if CarreraValidaDeuda(Login("CodCarpr")) then
	        'response.Write("Pasa por Validacion de Analiza Alumno")
   	        'IdRutAux = Trim(Request("logrut"))
                    
		    Sstr=" exec sp_alumno_deuda_ra '"& (IdRut) &"'"
            Session("Conn").Execute Sstr

		    Sstr=" Select Deuda from mt_client Where Codcli = '"& (IdRut) &"'"
			Set Rrst = Session("Conn").execute(Sstr)
			Bloqueo=valnulo(Rrst("DEUDA"), num_)

			if Bloqueo <> 0 then
				session("MIBLOQUEO")=valnulo(Rrst(0), num_)
			end if
			
'	     	Bloqueo=AlumnoConDeuda(IdRut)
			
		    'Login.Close()
		    'Session("Conn").Close() 
		    'Response.Redirect "atencion.asp?Error=3&Logueo=1" ' BF
			
			'response.Write(Bloqueo)
			'response.end()
			
			if Bloqueo=1 then
				session("m_bloque")="<blockquote><div align='justify'>Conforme a nuestros registros, a la fecha usted, presenta deuda con la Institucion.<br><br>Para regularizar esta situación deben concurrir al Departamento de Tesoreria y Cobranzas de la Institución.<br><br><br><br></div></blockquote>"
				Session("Logueado") = "1"
				session("BloqueoF")="CON BLOQUEO"
				session("BloqueoB")="SIN BLOQUEO"
			end if
			if Bloqueo=2 then
				session("m_bloque")="<blockquote><div align='justify'>Conforme a nuestros registros, a la fecha usted, presenta deuda de Biblioteca.<br><br>Para regularizar esta situación debe concurrir a la Biblioteca Institucional.<br><br></div></blockquote>"
				Session("Logueado") = "1"
				session("BloqueoF")="SIN BLOQUEO"
				session("BloqueoB")="CON BLOQUEO"
	    	end if
			if Bloqueo=3 then
				session("m_bloque")="<blockquote><div align='justify'>Conforme a nuestros registros, a la fecha usted, presenta deuda de Biblioteca y deuda con la Institución.<br><br>Para regularizar esta situación deben concurrir a la Biblioteca Institucional y al Departamento de Tesorería y Cobranzas.<br><br>Si a la fecha usted ya regularizó su situación, rogamos nos disculpe y de aviso a la oficina antes señalada..</div></blockquote>"
				Session("Logueado") = "1"
				session("BloqueoF")="CON BLOQUEO"
				session("BloqueoB")="CON BLOQUEO"
			end if
			if Bloqueo=0 then
				Session("Logueado") = "1"
				session("BloqueoF")="SIN BLOQUEO"
				session("BloqueoB")="SIN BLOQUEO"
			end if
			if Bloqueo="" then
				Session("Logueado") = "1"
				session("BloqueoF")="SIN BLOQUEO"
				session("BloqueoB")="SIN BLOQUEO"
			end if
	 else
				'response.Write("Pasa por Validacion de Analiza Alumno")
	 			Session("Logueado") = "1"
				session("BloqueoF")="SIN BLOQUEO"
				session("BloqueoB")="SIN BLOQUEO"
	end if

'response.Write("antes de validar bloq administrativo ")

	if BloqueoAdministrativoCarrera(Login("CodCarpr"))="S" then
		  
		if BloqueoAdministrativoAlumno(IdRut)<>"" then
			session("m_bloque")="<blockquote><div align='justify'>Conforme a nuestros registros, a la fecha usted, tiene documentos pendiente.<br><br>Para regularizar esta situación deben concurrir a las oficinas de Registro Curricular de sus respectivas Escuelas.<br><br>Si a la fecha usted ya regularizó su situación, rogamos nos disculpe y de aviso a la oficina antes señalada..</div></blockquote>"
			session("BloqueoA")="CON BLOQUEO"
		else
			session("BloqueoA")="SIN BLOQUEO"
		end if
	else
		session("BloqueoA")="SIN BLOQUEO"
	end if	

			 	
    Strsql = "select encuestaantes from mt_carrer where codcarr='" & trim(Login("CodCarpr")) & "'"  	
	if bcl_ado(strsql,rs) then
		if trim(valnulo(rs("encuestaantes"),str_)) = "S" then
	  		session("Cerrada") = "N"
		else
			session("Cerrada") = "S"
		end if
		rs.close
	end if



'response.Write(session("BloqueoF")) & "</br>"
'response.Write(session("BloqueoB")) & "</br>"
'response.Write(session("Logueado")) & "</br>"
if "Verificar si se usara esta secuencia" = "ya que ya no se usa ra_webmail" then
	If Valnulo(Login(4), STR_) <> "S" Then
	'Response.Write("Pasa por aqui ya que seguramente es un N em Cambio de Clave")
	'Response.End
	'strSqlAux = "Select rut  from mt_alumno where rut = '" & IdRut & "' and estacad = 'VIGENTE'"
	'comentè esto el dìa 14 de Enero...por que estamos tratando de integrar a todos los estados..y la carrera
	   	if CarreraAlumno <> "" then
		   strSqlAux = "Select rut  from mt_alumno with (nolock) where rut = '" & IdRut & "' and codcarpr='"& CarreraAlumno &"'"
        else
		   strSqlAux = "Select rut  from mt_alumno with (nolock) where rut = '" & IdRut & "'"
		end if
	
LoginAux.Open strSqlAux,Session("Conn")
		

       If LoginAux.Eof Then
'	      response.Write("Paso por aquì efectivamnete")
		  IdRutAux = Session("usrNW")' Trim(Request("logrut"))
		  Session("RutAlum") = IdRutAux
	      'LoginAux.Close()
	      'Session("Conn").Close()
		  Session("Logueado") = "1"	
		  session("EstadoMatriculado")="NO MATRICULADO"
		  'Response.Redirect("NoNumMatricula.asp")
		  'response.End()
	   End If
		 Clave=Session("MiClave")
		 'response.Write(Trim(LoginAux("rut")))
		 'response.End
		 
		If Trim(LoginAux("rut")) = Ucase(Trim(Clave)) Then '' **** AL SER IGUALES DEBE ENVIARSE A PAGINA DE CAMBIO DE CLAVE
          'Response.Write("Son Iguales")
	      Session("Logueado") = "1"
		  LoginAux.Close()
   		  Session("NomAlum") = Trim(Login(1)) & " " & Trim(Login(2))
		  IdRutAux = Session("usrNW")' Trim(Request("logrut"))
		  Session("RutAlum") = IdRutAux
		  Session("Rut") = IdRut
		  Session("CodCli") = Login("CodCli")
		  Session("mail_inst") = Login("mail_inst")
		  session("Codsede")=Login("Codsede")
		  session("Codcarr")=login("Codcarpr")
		  Session("codpestud")=login("Codpestud")
          Session("AnoAlumno") = login("ano")		  
          Session("Jornada") = login("Jornada")
		  GetPeriodoActivo
  		  Session("RutAlum") = IdRutAux
		  IF Session("RutAlum") = "" THEN 
   		  		'Session("RutAlum")=Session("RutCliente")
	  	  END IF 
		  Login.Close()
          Response.redirect("CambioClave.asp?CambioClave=1") 
		  Response.End()
		  'Session("Conn") = Conn
		  'Conn.Close() 
	   Else
'	   	  Response.Write("No Son Iguales por alguna Razon")
'		  response.End
	       LoginAux.Close()
  		   Login.Close()
		   Session("Conn").Close() 
		   'Response.Redirect "atencion.asp?Error=2&Logueo=1" ' Clave invalida
   		   Session("MiValor")=0
   		   Session("Logueado") = "0"
		   response.Write("aca se cae")
		   response.End()
		   session("MensajeLogin")="Clave incorrecta ..!!! Vuelva a Intentar\n(Recuerde que el campo password es sensible a mayúsculas y minúsculas)"
		Response.Redirect "alumnosfinning.asp" 
	  	Response.End()
	   End If	  
	End If   

end if

'    If UCase(Login(3)) = UCase(ClaveEncript) Then
'response.Write(Login("us_password")&encriptada)
'response.end()
    If UCase(Login("us_password")) = UCase(encriptada) Then
       Session("Logueado") = "1"   
	Else
	   'Response.Redirect "atencion.asp?Error=2&Logueo=1" ' Clave invalida
   	   Session("MiValor")=0
   	   Session("Logueado") = "0"
'	   response.Write("aca se cae")
'	   response.End()
   	   session("MensajeLogin")="Clave incorrecta ..!!! Vuelva a Intentar\n(Recuerde que el campo password es sensible a mayúsculas y minúsculas)"
		Response.Redirect "alumnosfinning.asp" 
	  	Response.End()
	End If      
End If	


Session("NomAlum") = Trim(Login(1)) & " " & Trim(Login(2))
IdRutAux = Session("usrNW")' Trim(Request("logrut"))
Session("RutAlum") = IdRutAux
CodCli = Trim(Login(5))
Session("Rut") = IdRut
Session("CodCli") = CodCli
Session("mail_inst") = Login("mail_inst")
session("Codsede")=Login("Codsede")
session("Codcarr")=login("Codcarpr")
Session("codpestud")=login("Codpestud")
Session("anoalumno") = login("ano")
Session("Jornada") = login("Jornada")
Session("Nivel") = login("Nivel")

Login.Close()
GetPeriodoActivo
StrSql = "delete from tmpSolici where Codcli = '" & Codcli & "' "
Session("Conn").execute StrSql
StrSql = "delete from tmpTest where Codcli = '" & Codcli & "' "
Session("Conn").execute StrSql
'Traemos el Permiso
carrera=Session("codcarr")
estado=session("estado")
bloqueoF=session("BloqueoF")
bloqueoB=session("BloqueoB")
Nivel=0
EstadoMatriculado=session("EstadoMatriculado")
bloqueoA=session("BloqueoA")
'response.write("carrera"&carrera&"estado"& estado&"bloqueof"& bloqueoF&"bloqueob"&bloqueoB&"estadomatriculados"&EstadoMatriculado&"bloqueoA"+bloqueoA)
'response.End()
perfilNW = GetPerfil(carrera, estado, bloqueoF,bloqueoB,EstadoMatriculado, bloqueoA)
Session("perfilNW") = perfilNW
'response.Write("perfil:" & perfilNW)
'response.End()
'********************************************
		'response.Write(carrera) & "<br>"
		'response.Write(estado) & "<br>"	
		'response.Write(bloqueoF) & "<br>"	
		'response.Write(bloqueoB) & "<br>"
		'response.Write(opcion1) & "<br>"
		'response.Write(nivel) & "<br>"
		'response.Write(EstadoMatriculado) & "<br>"
'********************************************

'Permiso=GetPermiso(carrera,estado,bloqueoF,bloqueoB,opcion1,nivel,EstadoMatriculado,bloqueoA)
if(Session("FromNet") = 1) then
	response.Redirect(Session("Destino"))
	response.Redirect("RedirectNet.asp") 
	response.End()
end if


	
'response.Write(carrera) & "</br>"
'response.Write("-Estado")
'response.Write(estado) & "</br>"
'response.Write("-BF")
'response.Write(bloqueoF) & "</br>"
'response.Write("-BB")
'response.Write(bloqueoB) & "</br>"
'response.Write("-Ma")
'response.Write(EstadoMatriculado) & "</br>"
'response.Write("-BA")
'response.Write(bloqueoA) & "</br>"
'response.Write("-Permiso")
'response.Write(session("Nivel")) & "</br>"
'response.End()

'response.Write(Permiso) & " este es el Permiso"
'response.End
'response.Write(session("ano_ed"))& "</br>"
'response.Write(session("periodo_ed")) & "</br>"'
'response.end()
if Permiso = "" then 
	Permiso = 0
end if 

%>

<% if Session("destino") = "PAGO_CUOTAS" then  %>
    <script>
        window.top.location.href = 'PagoWeb_CreaIntencion.asp';
    </script>
<%end if %>

<%permiso = 1
if Permiso=1 then%>

<%Session("MiValor")=0%>

	<%
	Audita 580,"Ingresa a Portal de Alumnos Login" 
	
	 if Session("SituacionActual") = "N" then %>
        <script>
        window.top.location.href='menu_tomaderamos.asp';
        //window.top.location.href='SituActual.asp';
        </script>
     <%else%>
        <script>
        //window.top.location.href='menu_tomaderamos.asp';
        window.top.location.href='SituActual.asp';
        </script>     
     
     <%end if%>
<%else%>
<%Session("MiValor")=0%>
<script>//window.top.location.href='MensajesBloqueosInicial.asp'; 
//esta funcion lo que tiene que envìa a una pagina pero en forma Entera..
</script> 
	<%
	if Session("SituacionActual") = "N" then %>
        <script>
        window.top.location.href='MensajesBloqueosInicial.asp';
        //window.top.location.href='SituActual.asp';
        </script> 
     <%else%>
        <script>
        //window.top.location.href='menu_tomaderamos.asp';
        window.top.location.href='SituActualBloqueoInicial.asp';
        </script>   
	 <%end if%>
<%end if%>


