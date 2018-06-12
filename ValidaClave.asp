<% 
   Response.Expires = -1 
%>
<!--#INCLUDE file="include/conexion.inc" -->
<!--#INCLUDE file="include/funciones.inc" -->
<!--#INCLUDE file="include/periodos.inc" --> 
<object runat="server" progid="ADODB.Recordset" id="Login" name="Login">
</object>
<object runat="server" progid="ADODB.Recordset" id="LoginEstacad" name="LoginEstacad">
</object>
<object runat="server" progid="ADODB.Recordset" id="LoginAux" name="LoginAux">
</object>
<object runat="server" progid="ADODB.Recordset" id="LoginB" name="LoginB">
</object>
<object runat="server" progid="ADODB.Recordset" id="LoginPeriodo" name="LoginPeriodo"> 
</object>
<%
Dim IdRut, IdRutAux, Clave, ClaveEncript, ClaveDesenc, strSql, strSqlAux, CodCli,SW,strb,IdRutB,sTRSQLP,Estado,RX
Dim msg,Bloqueo,Permiso,opcion1,EstadoMatriculado, mySQL,myRST,Carr,MatriculaSemestral,StrPreValida,StrCar
dim GetCar,CarreraAlumno, Sstr,Rrst, Rs, RstV,rstParame,mystr,claveNW,usrNW,sindig,perfilNW
NoExiste = 0
CarreraAlumno=""
Opcion1="INT_001"
session("m_bloque") =""


'Titulo Localizacion PA


strNP = "sp_trae_localizacion 'validaclave.asp',NULL"
set rstNP= Session("Conn").Execute(strNP)
if not rstNP.eof then
	Session("NombrePestana") = rstNP("VALOR")
else
	Session("NombrePestana") = "Portal de Alumnos en L&iacute;nea"
end if 

	session("LOCALIZACION") ="CHILE"
	strL="select ct_constanteValor from tg_constantes WHERE ct_constante='LOCALIZACION'"
	if BCL_ADO(strL, rstL) then
		session("LOCALIZACION") = rstL("ct_constanteValor")
	end if

if Session("MiValor")=0 then ' o cuando es vacío también pasa
'response.Write("entro 1")
	if Session("FromNet") = 1 then	
		IdRut = RutSinDV(EliminaPalabras(Trim(Session("RQRut"))))
		Session("RutCli")=idRut
		Session("RutCliente")=(Trim(Session("RQRut")))		
		strAL ="select top 1 1 from mt_alumno where rut ='"& Session("RutCli") &"'" 
		'response.Write(strAL)
		'response.End() 
		if bcl_ado(strAL,rstAL) then
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
			
			'nuevo para el redirect con carrera
			CarreraAlumno = Session("RQCodcarr")
			session("CarreraAlumno") = CarreraAlumno
			Session("EstadoCarrera") = GetEstado(IdRut,CarreraAlumno)
		end if
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
		
		if trim(ucase(session("TIPOVALIDACIONRUT"))) ="PERUANA" then 'Nuevo digitoP
			usrNW = usrNW & "0"
		end if
	
	end if 		
	Session("RutClienteR")=Session("RutCliente")
	Session("RutClienteR")= Replace(Session("RutClienteR"), ".", "")
	Session("RutClienteR")= Replace(Session("RutClienteR"), "-", "")		
	
	

	StrPreValida = " SELECT * FROM ca_usuarios WHERE us_consuser = '" & usrNW & "'"
	
	Set rstU = Session("Conn").execute(StrPreValida)
    if rstU.eof then		
		NoExiste = 1
	end if
	
	StrPreValida = " SELECT * FROM ca_usuarios WHERE us_consuser = '" & usrNW & "'  AND dbo.Encrypt('" & Clave & "' ) = us_password "		
	
	Set rstV = Session("Conn").execute(StrPreValida)	
    if rstV.eof then
        
        strParame="SELECT dbo.Fn_ValorParame('PORTADA_PROPIA_PORTAL_ALUMNO')Parame"
	    set rstParame= Session("Conn").Execute(strParame)
	    if not rstParame.eof then
		        salirnuevo=rstParame("Parame")
	    else
		        salirnuevo=""
	    end if 
	    if(salirnuevo <> "" and not isNull(salirnuevo)) then
	            pagina=rstParame("Parame")
                Response.Write("<script language='javascript'>alert('Clave incorrecta ..!!! Vuelva a Intentar.'); window.location.href='"  + pagina + "';</script>")
                response.End() 
	    else
				 'comentado x la mejora del cambio de pass
	             'Response.Redirect "ClaveIncorrecta.asp"  
	  	         'Response.End()	
	    end if 
        
	 	
	end if

		
Session("MiValor")=1 
end if

Clave = Trim(Request("logclave"))
	
	 GetCar= Session("GetCarr")
	 CarreraAlumno=session("CarreraAlumno")
	 IdRut=Session("RutCli")
	 claveNW = Session("MiClave")


	 strSql ="select matriculasemestral, coalesce(websituactual, 'N') as websituactual from mt_parame"
	 strSql ="SELECT dbo.Fn_ValorParame('matriculasemestral')matriculasemestral,	dbo.Fn_ValorParame('websituactual')websituactual"

		if bcl_ado(strsql,rstParame) then
			MatriculaSemestral = trim(valnulo(rstParame("matriculasemestral"),str_))
			SituacionActual = trim(valnulo(rstParame("websituactual"),str_))
			rstParame.close
		end if

	Session("SituacionActual")= SituacionActual
	
if Session("usrNWRC")<>"" then
	usrNW = Session("usrNWRC")   
end if 

usrNW = Replace(usrNW, ".", "")
usrNW = Replace(usrNW, "-", "")	
Session("usrNW") = usrNW

	strParameNM="SELECT coalesce(dbo.Fn_ValorParame('FILTRAESTADOVIGENTEPA'),'NO')Parame"
	set rstParameNM= Session("Conn").Execute(strParameNM)	
		
	if not rstParameNM.eof then
		FILTRAESTADOVIGENTEPA=rstParameNM("Parame")
	end if 

	strSql= "SP_DATOS_PORTAL_ALUMNO '" & usrNW & "','"& CarreraAlumno &"','" & trim(Session("EstadoCarrera")) &"','SI'"
		
Dim claveD
Dim encriptada
claveNW = Session("MiClave")
mystr = "select dbo.Encrypt('"&claveNW&"') as clave"

Set rst = Session("Conn").execute(mystr)
encriptada = Rst("clave")

Login.Open strSql,Session("Conn")

if Login.eof and login.bof then

	strSql= "SP_DATOS_PORTAL_ALUMNO '" & usrNW & "','"& CarreraAlumno &"','" & trim(Session("EstadoCarrera")) &"','NO'"
	
	LoginEstacad.Open strSql,Session("Conn")

	if LoginEstacad.eof and LoginEstacad.bof then
		Session("MiValor")=0
		NoExiste = 1
	else
		if FILTRAESTADOVIGENTEPA ="SI" then
			Session("MiValor")=0
			session("MensajeBloqueosVarios") ="Estimado alumno, el acceso se encuentra deshabilitado para alumnos no vigentes."
			response.Redirect("MensajeBloqueo.asp")
			response.End()
		else
			Session("MiValor")=0
			NoExiste = 1
		end if 
		
	end if 

end if

if NoExiste = 1 then 'Nuevo

	strMatWebTr="SELECT coalesce(dbo.Fn_ValorParame('USAMATRICULAWEB_TR'),'NO')Parame"
	set rstMatWebTr= Session("Conn").Execute(strMatWebTr)		
	if not rstMatWebTr.eof then
		USAMATRICULAWEB_TR=rstMatWebTr("Parame")
	end if
	
	if USAMATRICULAWEB_TR = "SI" then

		destinoRedirect ="menu_finanzas.asp"
		
		if Session("FromNet") = 1 then	
			LoginPostul = RutSinDV(EliminaPalabras(Trim(Session("RQRut"))))
			RutX = Session("RQRut")
			ClavePostul = Session("RQpass")
			destinoRedirect = Session("Destino")
			Session("MiClave")=(Session("RQpass"))
			Session("RutClienteR") = RutX
		else
			LoginPostul = EliminaPalabras(Trim(Session("RutCli")))
			RutX = Session("RutCli")
			ClavePostul = Request("logclave")
			Session("RutCli") = Session("RutCli") & "0"
		end if 
		
		strVP = "SP_VALIDA_LOGIN_POSTUL_PA '"& Session("RutCli") & "','"& ClavePostul & "'"	
		'response.Write(strVP)
		'response.End()
		if bcl_ado(strVP,rstVP) then
			Session("NomAlum")  = rstVP("NomAlum")
			Session("RutAlum")  = Trim(RutX)
			Session("Rut")      = Trim(RutX)
			Session("CodCli")   = "POSTULANTE"
			Session("mail_inst")= rstVP("mail_inst")
			session("Codsede")  = rstVP("Codsede")
			session("Codcarr")  = rstVP("Codcarr")
			Session("codpestud")= rstVP("Codpestud")
			Session("anoalumno")= rstVP("ano")
			Session("Jornada")  = rstVP("Jornada")
			Session("Nivel")    = rstVP("Nivel")		
			Session("perfilNW") = rstVP("perfil")
			Session("Logueado") = 1
			
			strPago="SP_VALIDA_PAGO_CUOTA_BASICA '"& Session("RutCli") &"' ,'"& session("codcarr") &"'"
			if BCL_ADO(strPago, rstPago) then
				PAGO = rstPago("PAGO")
			end if
			
			if PAGO = "NO" then
				session("MensajeBloqueosVarios") ="Para acceder a esta opci&oacute;n debe pagar su cuota b&aacute;sica de matr&iacute;cula."
				destinoRedirect = "MensajeBloqueo.asp"
			else
				destinoRedirect = "RedirectNet.asp?PaginaRedirect=matriculaweb.aspx"
			end if
														
			response.Redirect(destinoRedirect)
			response.End()
		else
			response.Redirect("MensajeNoExiste.asp")
			response.End()
		end if
	else
		response.Redirect("MensajeNoExiste.asp")
		response.End()
	end if
end if 



Carr=login("codcarpr")
Session("id_usuario") = login("id_usuario")

IdRut = Login("rut")

if CarreraAlumno ="" then
	estado=GetEstado2(IdRut)
else
	estado=GetEstado(IdRut,CarreraAlumno)
end if

session("estado")=estado
ClaveEncript=Session("ClaveEncriptada")
	  
If Login.Eof Then
	  IdRutAux = Session("usrNW")' Trim(Request("logrut"))
	  Session("RutAlum") = IdRutAux
	  Session("Logueado") = "1"
	  session("EstadoMatriculado")="NO MATRICULADO"	
Else ''' ********** VERIFICAR QUE EL PRIMER LOGIN SEA CON NUMERO DE MATRICULA


session("EstadoMatriculado")="MATRICULADO"
GetPeriodoEd  IdRut , Carr
GetBuscaRamosInasistentes  IdRut , Carr

GetHabilitaBotonEncuenta IdRut, Carr                                                                                                    

NumMaxPer = session("NumeroMaximoPrueba")

	strMatriculado="Pr_Valida_Matricula_Alumno  '" & Carr & "','" & IdRut & "','" & Login("codcli") & "'"
	
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
	
	Session("Logueado") = "1"

	if Session("EstadoMatriculado")="NO MATRICULADO" then
	
		strParameNM="SELECT dbo.Fn_ValorParame('MUESTRAPERIODONOMATRICULADO')Parame"
		set rstParameNM= Session("Conn").Execute(strParameNM)		
		if not rstParameNM.eof then
				MUESTRAPERIODONOMATRICULADO=rstParameNM("Parame")
			else
				MUESTRAPERIODONOMATRICULADO="" 
		end if
		
		oficina="Registro Curricular"
		strI="select TOP 1 COALESCE(CODINST,'')CODINST from RA_INSTITUCION"
		if bcl_ado(strI,Irst) then
			inst = Irst("CODINST")
			if inst ="UAC" then
				oficina="Matricula" 
			end if 			
		end if
		
		if MUESTRAPERIODONOMATRICULADO ="SI" then 
		
			strProc ="SP_TRAEANOPERIODOMAT '"& session("codcli") &"'"  
			if bcl_ado(strProc,rstProc) then
				ANOTM = rstProc("ANO")
				PERIODOTM = rstProc("PERIODO")
			end if
		
			session("m_bloque")="Conforme a nuestros registros, a la fecha usted, presenta estado No Matriculado para el periodo "& ANOTM &"/"& PERIODOTM &".<br>Para regularizar esta situación debe concurrir a las oficinas de "& oficina &".<br>"
			
		else		
			session("m_bloque")="Conforme a nuestros registros, a la fecha usted, presenta estado No Matriculado.<br>Para regularizar esta situación debe concurrir a las oficinas de "& oficina &".<br>"
		end if 
	end if	
	
if BloqueoBibliotecaCarrera(Login("CodCarpr"))="S" then
		if BloqueoBibliotecaAlumno(IdRut)="S" then
			''JJJ
			session("m_bloque")="Conforme a nuestros registros, a la fecha usted, presenta deuda de Biblioteca.<br>Para regularizar esta situación debe concurrir a la Biblioteca Institucional.<br>"
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


	if CarreraValidaDeuda(Login("CodCarpr")) then                    
		    Sstr=" exec sp_alumno_deuda '"& (IdRut) &"'"
            Session("Conn").Execute Sstr

		    Sstr=" Select Deuda from mt_client Where Codcli = '"& (IdRut) &"'"
			Set Rrst = Session("Conn").execute(Sstr)
			Bloqueo=valnulo(Rrst("DEUDA"), num_)
			
			if Bloqueo <> 0 then
				session("MIBLOQUEO")=valnulo(Rrst(0), num_)
			end if
			
			if Bloqueo=1 then
				session("m_bloque")="Conforme a nuestros registros, a la fecha usted, presenta deuda con la Institucion.<br>Para regularizar esta situación deben concurrir al Departamento de Tesoreria y Cobranzas de la Institución.<br>"
				Session("Logueado") = "1"
				session("BloqueoF")="CON BLOQUEO"
				session("BloqueoB")="SIN BLOQUEO"
			end if
			if Bloqueo=2 then
				session("m_bloque")="Conforme a nuestros registros, a la fecha usted, presenta deuda de Biblioteca.<br>Para regularizar esta situación debe concurrir a la Biblioteca Institucional.<br>"
				Session("Logueado") = "1"
				session("BloqueoF")="SIN BLOQUEO"
				session("BloqueoB")="CON BLOQUEO"
	    	end if
			if Bloqueo=3 then
				session("m_bloque")="Conforme a nuestros registros, a la fecha usted, presenta deuda de Biblioteca y deuda con la Institución.<br>Para regularizar esta situación deben concurrir a la Biblioteca Institucional y al Departamento de Tesorería y Cobranzas.<br>Si a la fecha usted ya regularizó su situación, rogamos nos disculpe y de aviso a la oficina antes señalada.<br>"
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

'response.Write("session(BloqueoF): " & session("BloqueoF"))
'response.Write("session(BloqueoB): " & session("BloqueoB"))
'response.end


	if BloqueoAdministrativoCarrera(Login("CodCarpr"))="S" then
		  
		if BloqueoAdministrativoAlumno(IdRut)<>"" then
			session("m_bloque")="Conforme a nuestros registros, a la fecha usted, tiene documentos pendiente.<br>Para regularizar esta situación deben concurrir a las oficinas de Registro Curricular de sus respectivas Escuelas.<br>Si a la fecha usted ya regularizó su situación, rogamos nos disculpe y de aviso a la oficina antes señalada.<br>"
			session("BloqueoA")="CON BLOQUEO"
		else
			session("BloqueoA")="SIN BLOQUEO"
		end if
	else
		session("BloqueoA")="SIN BLOQUEO"
	end if	

	'Mejora de BloqueAdm.EA
	if session("BloqueoA")="SIN BLOQUEO" then
		Strsql = "SP_VALIDABLOQUEO_ADM_PA '"& session("codcli") &"'"  	
		if bcl_ado(strsql,RstV) then
			if RstV("TIENEBLOQUEOADM") = "SI" then
				session("BloqueoA")="CON BLOQUEO"
				session("m_bloque")="Conforme a nuestros registros, a la fecha usted, presenta un bloqueo administrativo.<br>Para regularizar esta situación deben concurrir a la Institución.<br>"
			else
				session("BloqueoA")="SIN BLOQUEO"
			end if
			RstV.close
		end if
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

	'saca el id_usuario
	strIUSer="SELECT id_usuario FROM ca_usuarios WHERE us_consuser = CONVERT(VARCHAR,'" & usrNW & "')"
	Set rstIUSer = Session("Conn").execute(strIUSer)
	
	If UCase(Login("us_password")) = UCase(encriptada) Then
		Session("Logueado") = "1"    
	
		'Mejoras politicas de Password, SI
		str_validaBloqExp ="sp_valida_us_bloqueo_expiracion @id_usuario='" & rstIUSer("id_usuario") & "',@claveCorrecta='SI',@salida='',@PortalAlumno='SI'"
		'str_validaBloqExp ="select 0 salida"
		Set Rst_validaBloqExp = Session("Conn").execute(str_validaBloqExp)
			
		if not Rst_validaBloqExp.eof  then
			salida = Rst_validaBloqExp("salida")
			
			if left(salida,1) = "4" then 'cuando expira la clave lo envia a cambiar la pass
				Response.Write("<script language='javascript'>window.top.location.href = 'cambioclaveexterno.asp';</script>") 			
				response.End()
			end if
			
			if left(salida,1) = "1" then
				MensajeClave = "- El usuario esta bloqueado."
				Response.Write("<script language='javascript'>alert('"& MensajeClave &"');window.top.location.href = 'alumnos.asp';</script>")
			elseif left(salida,1) = "2" then
				MensajeClave = "- La fecha de expiración del usuario está vencida."
				Response.Write("<script language='javascript'>alert('"& MensajeClave &"');window.top.location.href = 'alumnos.asp';</script>")
			elseif left(salida,1) = "4" then	
				MensajeClave = "- Su clave ha expirado."
				Response.Write("<script language='javascript'>alert('"& MensajeClave &"');window.top.location.href = 'alumnos.asp';</script>")
			elseif left(salida,1) = "6" then
				dias = right(salida,1)
				if dias = "1" then
					dias = dias & " día."
				else
					dias = dias & " días."
				end if 	
				Response.Write("<script language='javascript'>alert('- Su clave expirará en " & dias & "');</script>")
			end if 
		end if 
			   
	Else
   	   Session("MiValor")=0
   	   Session("Logueado") = "0"
	   
	   'Mejoras politicas de Password, NO
	   '+1 intentos fallidos
	   strIFail="UPDATE ca_usuarios SET us_intentosfallidos = COALESCE(us_intentosfallidos,0)+1 WHERE id_usuario='" & rstIUSer("id_usuario") & "'"
       Session("Conn").Execute strIFail
	   
	   str_validaBloqExp ="sp_valida_us_bloqueo_expiracion @id_usuario='" & rstIUSer("id_usuario") & "',@claveCorrecta='NO',@salida='',@PortalAlumno='SI'"
	   Set Rst_validaBloqExp = Session("Conn").execute(str_validaBloqExp)
	   
	   if not Rst_validaBloqExp.eof  then
	   		salida = Rst_validaBloqExp("salida")
			
			if left(salida,1) = "0" then ' cuando tiene el parametro desactivado de las politicas de claves
				MensajeClave = "- Clave incorrecta .Vuelva a Intentar \n(Recuerde que el campo password es sensible a mayúsculas y minúsculas)."
			elseif left(salida,1) = "1" then
				MensajeClave = "- El usuario esta bloqueado."
			elseif left(salida,1) = "3" then			
				MensajeClave = "- Ha excedido el límite de intentos para ingresar."
			elseif left(salida,1) = "5" then			
				intentos = right(salida,1)
				if intentos = "1" then
					intentos = intentos & " intento para ingresar."
				else
					intentos = intentos & " intentos para ingresar."
				end if 			 				
				MensajeClave = "- La clave es incorrecta, tiene " & intentos & ""
			end if 		 
			Response.Write("<script language='javascript'>alert('"& MensajeClave &"');window.top.location.href = 'alumnos.asp';</script>") 			
			response.End()  
	   end if 
	End If      
End If	
   
if Session("MiValor")=1 then
	if GetVariasCarreras(idRut)="Verdadero" and session("CarreraAlumno") = "" then
	   Session("GetCarr")=1
	   session("RutAlum")=Trim(Session("RutCliente"))
	   Response.Redirect "EligeCarrera.asp"
	   response.End 
	else			
	   Session("GetCarr")=0
	   
	end if
	Session("MiValor")=2
end if

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


if(Session("FromNet") = 1) then
	response.Redirect(Session("Destino"))
	response.Redirect("RedirectNet.asp") 
	response.End()
end if

session("TomadeRamo")=TomadeRamo(session("Codcarr"),session("Codcli"),session("ano_ed"),session("periodo_ed"))

'Se comenta para que predomine el parametro websituactul

'strParame="SELECT dbo.Fn_ValorParame('BLOQUEAPAENCUESTAS')Parame"
'set rstParame= Session("Conn").Execute(strParame)		
'if not rstParame.eof then
'		BLOQUEAPAENCUESTAS=rstParame("Parame")
'	else
'		BLOQUEAPAENCUESTAS="" 
'end if

'if session("TomadeRamo")="0"  and BLOQUEAPAENCUESTAS="SI"then
'	Session("SituacionActual")="S"
'end if 

if Permiso = "" then 
	Permiso = 0
end if 

BLOQUEOD=""

strParame="SELECT coalesce(dbo.Fn_ValorParame('ACTIVABLOQUEODOCUMENTOS'),'')Parame"
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
	ACTIVABLOQUEODOCUMENTOS=rstParame("Parame")
end if 

if ACTIVABLOQUEODOCUMENTOS = "SI"  then
	strFD="SELECT dbo.FN_VALIDA_DOCUMENTOS_ALUMNO('"& IdRut &"')valor"
	set rstFD= Session("Conn").Execute(strFD)
	if not rstFD.eof then
		BLOQUEO = rstFD("valor")
	end if 
	
	if BLOQUEO <> 0 then
		session("m_bloque") ="Para acceder al portal debe entregar todos los documentos pendientes."
		response.Redirect("MensajesBloqueosInicial.asp") 
	end if
end if 

'Mejora vc validacion de mails
strvma ="SP_EXIJE_MAIL_PA"
if bcl_ado(strvma,rstvma) then
	if rstvma(0) = "SI" then
		strvmac ="SP_CONSULTA_CAMBIO_MAIL '"& SESSION("Rutcli") &"'"
		if bcl_ado(strvmac,rstvmac) then
			if rstvmac(0) = "NO" then
				response.Redirect("cambiomail.asp")
			end if 	
		end if 
	end if 	
end if 

%>
<% if Session("pagar") = "1" then  %>

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
    window.top.location.href = 'menu_tomaderamos.asp';
    //window.top.location.href='SituActual.asp';
</script>

<%else%>

<script>
    //window.top.location.href='menu_tomaderamos.asp';
    window.top.location.href = 'SituActual.asp';
</script>

<%end if%>
<%else%>
<%Session("MiValor")=0%>

<script>    //window.top.location.href='MensajesBloqueosInicial.asp'; 
    //esta funcion lo que tiene que envìa a una pagina pero en forma Entera..
</script>

<%
	if Session("SituacionActual") = "N" then %>

<script>
    window.top.location.href = 'MensajesBloqueosInicial.asp';
    //window.top.location.href='SituActual.asp';
</script>

<%else%>

<script>
    //window.top.location.href='menu_tomaderamos.asp';
    window.top.location.href = 'SituActualBloqueoInicial.asp';
</script>

<%end if%>
<%end if%>
