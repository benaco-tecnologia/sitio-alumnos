<%
	Response.Expires = -1
%>
<!--#include file="addon/includes/funciones.inc" -->
<!--#include file="addon/includes/libreria.asp" -->
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

<%
Dim IdLoginRut, IdLoginPassWord
Dim Login, MT_CLIENT, MT_ALUMNO, RA_ALUMNO
Dim fso, carpeta, tempSite
Dim Ramos(130)
Dim Conter, i

'IdLoginRut=RutSinDV(EliminaPalabras(Trim(Session("RutCliente"))))
IdLoginRut = Session("Rut")
IdLoginPassWord=Encripta(EliminaPalabras(Trim(Session("MiClave"))))


'response.Write(Session("MiClave") & " y " & Session("usrNW"))
'response.end

'Set Login = getAlumno(IdLoginRut, IdLoginPassWord)
Set Login = getAlumnoNW(Session("usrNW"), Session("MiClave"))

'response.Write(IdLoginRut) 
'response.End()

if Not Login.Bof or Not Login.Eof _
	then
	    '// Cargar Datos
		set MT_CLIENT = get_mtClient(IdLoginRut) 
		set MT_ALUMNO = Session("Conn").Execute("SELECT CODCLI,CODCARPR FROM mt_alumno WHERE codcli='" & session("codcli") & "'")'get_mtAlumno(IdLoginRut)
		if MT_ALUMNO.EOF _
			then
				response.Redirect("NoAccess.asp")
				response.End()
		end if
        Session ("LoginErr")="0"
        Session ("OnLine")="1"
        Session ("LoginRutOnLine")=IdLoginRut
        Session ("LoginPassWord")=IdLoginPassWord
		Session ("IdAlumno")=MT_CLIENT("NOMBRE") & " " & MT_CLIENT("PATERNO") & " " & MT_CLIENT("MATERNO")
		Session ("eMailAlumno")=MT_CLIENT("MAIL")
		Session ("CarreraEnCurso")=MT_ALUMNO("CODCLI")
		Session ("CodigoCarrera")=MT_ALUMNO("CODCARPR")
		'// Determina Ramos y Secciones
		set RA_ALUMNO=get_raAlumno(MT_ALUMNO("CODCLI"))
		if Not RA_ALUMNO.Bof or Not RA_ALUMNO.Eof _			
			then
				Conter=0
				RA_ALUMNO.MoveFirst
				While Not RA_ALUMNO.Eof
					Ramos(Conter) = RA_ALUMNO("CODSECC") & "-" & RA_ALUMNO("CODRAMO") & "-" & RA_ALUMNO("ANO") & "-" & RA_ALUMNO("PERIODO")
					Conter = Conter + 1
					RA_ALUMNO.MoveNext
				Wend
				Session("matRamos")=Ramos
			else
				Ramos(0)="X"
				Session("matRamos")=Ramos
		end if

		Set RA_ALUMNO = Nothing
		Set MT_CLIENT = Nothing
		Set MT_ALUMNO = Nothing
		Set Login = Nothing
		
        Response.Redirect ("main.asp?cmd=1")
        Response.End		
	else
		//'// Mensaje Error
		Session ("LoginErr")="1"
		Response.Redirect ("MensajeSinDatos.asp")
		Response.End		
End if
%>