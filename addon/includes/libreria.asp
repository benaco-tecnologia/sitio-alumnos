<%
Function SqlSeccion()
	if Session("Seccion") = 1 _
		then
			SqlSeccion = "Where Seccion IN (1,2,3,4)"
		else
			SqlSeccion = "Where Seccion = '" & Session("Seccion") & "'"
	end if
	
end function
Function CreateFolderFile(ThisFolderFile)
	Dim fs
	Set fs = Server.CreateObject("Scripting.FileSystemObject")
	if (Not fs.FolderExists(ThisFolderFile)) then
		Set fol = fs.CreateFolder(ThisFolderFile)
		Set fol = Nothing
	end if
End function

Function deleteFile(archivo) 
    dim fs 
    Set fs = Server.CreateObject("Scripting.FileSystemObject") 
    if fs.FileExists(archivo) then 
        fs.DeleteFile(archivo) 
        deleteFile = true 
    else 
        deleteFile = false 
    end if 
    Set fs = Nothing 
End function

Function getGUID()
    Set typ = CreateObject("Scriptlet.TypeLib") 
    guidStr = cstr(typ.GUID) 
   getGUID = LEFT(guidStr, 38) 
    set typ = nothing 
End Function

Function cr2br(s) 
   cr2br = Replace (s, vbCrLf, "<BR>") 
End Function

Function getClasificacionId(Inx)
	dim Clasificacion(6)
	Clasificacion(1)="ADMINISTRADORES"
	Clasificacion(2)="RELACIONES PUBLICAS"
	Clasificacion(3)="FINANZAS"
	Clasificacion(4)="DIRECCION ACADEMICA"
	Clasificacion(5)="D.A.E"
	Clasificacion(6)="PASTORAL"	
	getClasificacionId=Clasificacion(cInt(Inx))
End Function

Function getAdministradores()
	Set getAdministradores = Session("Conn").Execute("SELECT * FROM ca_UDS ORDER BY Seccion")
End Function

Function getExisteUser(User)
	Set rs =  Session("Conn").Execute("IF exists(Select * From ca_UDS Where Usuario='" & User & "') BEGIN Select Existe = 1 END ELSE BEGIN Select Existe = 0 END")
	if (cInt(rs("Existe"))) _
		then
			getExisteUser = true
		else
			getExisteUser = false
	end if 
End Function

Function getAdmin(caID)
	Set getAdmin = Session("Conn").Execute("Select * From ca_UDS Where caID='" & caID & "'")
End Function

Function getUserAdmin(User, Pass, Seccion)
	Set getUserAdmin = Session("Conn").Execute("Select * From ca_UDS Where Usuario='" & User & "' And Pass='" & Pass & "' And Seccion='" & Seccion & "'")
End Function

Function getUserAdminSinSeccion(User, Pass)
	Set getUserAdminSinSeccion = Session("Conn").Execute("Select * From ca_UDS Where Usuario='" & User & "' And Pass='" & Pass & "'")
End Function

Function InsertAdmin(NombreUsuario, User, Pass, Seccion)
	Set InsertAdmin = Session("Conn").Execute("Insert Into ca_UDS(Nombre, Usuario, Pass, Seccion, FechaMov, caID) Values('" & NombreUsuario & "', '" & User & "', '" & Pass & "', '" & Seccion & "',getdate(), newid())")
End Function

Function UpdateAdmin(caID, NombreUsuario, User, Pass, Seccion)
	Set UpdateAdmin = Session("Conn").Execute("Update ca_UDS Set Nombre='" & NombreUsuario & "', Usuario='" & User & "', Pass='" & Pass & "', Seccion='" & Seccion & "', FechaMov=getdate() Where caID='" & caID & "'")
End Function

Function DeleteAdmin(caID)
	Set DeleteAdmin = Session("Conn").Execute("Delete From ca_UDS Where caID='" & caID & "'")
End Function

Function getAlumno(IdLoginRut, IdLoginPassWord)
'//	Set getAlumno = Session("Conn").Execute("SELECT * FROM CA_Access WHERE rut='" & IdLoginRut & "' AND syspass='" & IdLoginPassWord & "'")
	Set getAlumno = Session("Conn").Execute("SELECT * FROM mt_webmail WHERE rut='" & IdLoginRut & "' AND passweb='" & IdLoginPassWord & "'")
End Function

Function getAlumnoNW(IdLoginRut, IdLoginPassWord)
'//	Set getAlumno = Session("Conn").Execute("SELECT * FROM CA_Access WHERE rut='" & IdLoginRut & "' AND syspass='" & IdLoginPassWord & "'")
'response.Write(IdLoginRut + " = " + IdLoginPassWord)
'response.End()

	Set getAlumnoNW = Session("Conn").Execute("SELECT * FROM ca_usuarios WHERE us_consuser='" & IdLoginRut & "' AND us_password= dbo.Encrypt('" & IdLoginPassWord & "')")
	
End Function

Function get_mtClient(IdLoginRut)
	Set get_mtClient = Session("Conn").Execute("SELECT * FROM mt_client WHERE codcli='" & IdLoginRut & "'")
End Function

Function get_mtAlumno(IdLoginRut)
	Set get_mtAlumno = Session("Conn").Execute("SELECT * FROM mt_alumno WHERE rut='" & IdLoginRut & "' And EstAcad='VIGENTE'")
End Function

Function get_raAlumno(CodCli)
	Set get_raAlumno = Session("Conn").Execute("select a.CODSECC, a.RAMOEQUIV as CODRAMO, a.ANO, a.PERIODO from ra_nota as a where a.codcli='" & CodCli & "' and a.ANO='" & year(date()) & "' order by a.ano, a.periodo")
End Function

Function getMensajesScroll()
	Set getMensajesScroll = Session("Conn").Execute("Select * From ca_Scroll " & SqlSeccion())
End Function

Function getScrollMsg(ID)
	Set getScrollMsg = Session("Conn").Execute("Select * From ca_Scroll Where ID='" & ID & "'")
End Function

Function InsertScroll(IDUSER, FechaInicio, FechaTermino, Activado, Mensaje)
	Set rs = getAdmin(IDUSER)
	Set InsertScroll = Session("Conn").Execute("Insert Into ca_Scroll(ID, Seccion, idUser, Mensaje, Activo, FechaInicial, FechaFinal, FechaActualizacion) Values(newid(), '" & rs("Seccion") & "', '" & IDUSER & "', '" & Mensaje & "', '" & Activado & "', '" & FechaInicio & "', '" & FechaTermino & "', '" & date() & "')")
End Function

Function UpdateScroll(IDmsg, IDUSER, FechaInicio, FechaTermino, Activado, Mensaje)
	Set rs = getAdmin(IDUSER)
	Set UpdateScroll = Session("Conn").Execute("Update ca_Scroll Set Seccion='" & rs("Seccion") & "', idUser='" & IDUSER & "', Mensaje='" & Mensaje & "', Activo='" & Activado & "', FechaInicial='" & FechaInicio & "', FechaFinal='" & FechaTermino & "', FechaActualizacion='" & date() & "' Where id='" & IDmsg & "'")
End Function

Function DeleteMsgScroll(IDmsg)
	Set DeleteMsgScroll = Session("Conn").Execute("Delete From ca_Scroll Where id='" & IDmsg & "'")
End Function

Function SwitchMsgScroll(IDmsg, SwOnOff)
	Set SwitchMsgScroll = Session("Conn").Execute("Update ca_Scroll Set Activo='" & SwOnOff & "' Where id='" & IDmsg & "'")
End Function

Function getMsgScrollActivos()
	Set getMsgScrollActivos = Session("Conn").Execute("Select * From ca_Scroll Where Activo='1' and getdate()>=FechaInicial and getdate()<=FechaFinal")
End Function

Function getNoticias()
	Set getNoticias = Session("Conn").Execute("Select * From ca_Noticias " & SqlSeccion() & " order by FechaActualizacion DESC")
End Function

Function getNoticia(ID)
	Set getNoticia = Session("Conn").Execute("Select * From ca_Noticias Where ID='" & ID & "'")
End Function

Function InsertNoticia(IDUSER, FechaInicio, FechaTermino, Activado, Encabezado, txtNoticia)
	Set rs = getAdmin(IDUSER)
	Set InsertNoticia = Session("Conn").Execute("Insert Into ca_Noticias(ID, Seccion, idUser, Encabezado, Activo, FechaInicial, FechaFinal, FechaActualizacion, Noticia) Values(newid(), '" & rs("Seccion") & "', '" & IDUSER & "', '" & Encabezado & "', '" & Activado & "', '" & FechaInicio & "', '" & FechaTermino & "', '" & date() & "', '" & txtNoticia & "')")
End Function

Function UpdateNoticia(IDmsg, IDUSER, FechaInicio, FechaTermino, Activado, Encabezado, txtNoticia)
	Set rs = getAdmin(IDUSER)
	Set UpdateNoticia = Session("Conn").Execute("Update ca_Noticias Set Seccion='" & rs("Seccion") & "', idUser='" & IDUSER & "', Encabezado='" & Encabezado & "', Activo='" & Activado & "', FechaInicial='" & FechaInicio & "', FechaFinal='" & FechaTermino & "', FechaActualizacion='" & date() & "', Noticia='" & txtNoticia & "' Where id='" & IDmsg & "'")
End Function

Function DeleteNoticia(IDmsg)
	Set DeleteNoticia = Session("Conn").Execute("Delete From ca_Noticias Where id='" & IDmsg & "'")
End Function

Function SwitchNoticia(IDmsg, SwOnOff)
	Set SwitchNoticia = Session("Conn").Execute("Update ca_Noticias Set Activo='" & SwOnOff & "' Where id='" & IDmsg & "'")
End Function

Function getTodasNoticias()
	Set getTodasNoticias = Session("Conn").Execute("Select * From ca_Noticias Where Activo='1' and getdate()>=FechaInicial and getdate()<=FechaFinal order by FechaActualizacion, FechaInicial DESC")
End Function

Function getEventos()
	Set getEventos = Session("Conn").Execute("Select * From ca_Eventos " & SqlSeccion())
End Function

Function getEvento(ID)
	Set getEvento = Session("Conn").Execute("Select * From ca_Eventos Where ID='" & ID & "'")
End Function

Function InsertEvento(IDUSER, FechaInicio, FechaTermino, Activado, Mensaje)
	if Activado = "" Then Activado=0 end if
	Set rs = getAdmin(IDUSER)
	Set InsertEvento = Session("Conn").Execute("Insert Into ca_Eventos(ID, Seccion, idUser, Mensaje, Activo, FechaInicial, FechaFinal, FechaActualizacion) Values(newid(), '" & rs("Seccion") & "', '" & IDUSER & "', '" & Mensaje & "', '" & Activado & "', '" & FechaInicio & "', '" & FechaTermino & "', '" & date() & "')")
End Function

Function UpdateEvento(IDmsg, IDUSER, FechaInicio, FechaTermino, Activado, Mensaje)
	if Activado = "" Then Activado=0 end if
	Set rs = getAdmin(IDUSER)
	Set UpdateEvento = Session("Conn").Execute("Update ca_Eventos Set Seccion='" & rs("Seccion") & "', idUser='" & IDUSER & "', Mensaje='" & Mensaje & "', Activo='" & Activado & "', FechaInicial='" & FechaInicio & "', FechaFinal='" & FechaTermino & "', FechaActualizacion='" & date() & "' Where id='" & IDmsg & "'")
End Function

Function DeleteEvento(IDmsg)
	Set DeleteEvento = Session("Conn").Execute("Delete From ca_Eventos Where id='" & IDmsg & "'")
End Function

Function SwitchEvento(IDmsg, SwOnOff)
	Set SwitchEvento = Session("Conn").Execute("Update ca_Eventos Set Activo='" & SwOnOff & "' Where id='" & IDmsg & "'")
End Function

Function getEventosDia(ThisDay)
	Set getEventosDia = Session("Conn").Execute("Select * From ca_Eventos Where Activo='1' and '" & ThisDay & "'>=FechaInicial and '" & ThisDay & "'<=FechaFinal")
End Function

Function getProfesor(IDrut)
	Set getProfesor = Session("Conn").Execute("Select * From ra_profes Where Rut='" & IDrut & "'")
End Function

Function getNominaProfesoresSelect()
	Set getNominaProfesoresSelect = Session("Conn").Execute("Select * From ra_profes Where Rut > 1 Order By Ap_Pater, Ap_Mater, Nombres")
End Function

Function getDoctosDelProfe(ID)
	Set getDoctosDelProfe = Session("Conn").Execute("Select * From ca_Doctos Where CodProfe='" & ID & "'")
End Function

Function getDoctoDelProfe(ID)
	Set getDoctoDelProfe = Session("Conn").Execute("Select * From ca_Doctos Where IdConter='" & ID & "'")
End Function

Function getCarrerasPorProfesor(CodigoProfe)
	Set getCarrerasPorProfesor = Session("Conn").Execute("Select * From mt_carrer Where CodCarr IN(Select DISTINCT CodCarr From ra_horprof Where CodProf = '" & CodigoProfe & "' AND Ano >= year(getdate()))")
End Function

Function getRamosPorProfesor(CodigoProfe)
	Set getRamosPorProfesor = Session("Conn").Execute("Select * From ra_ramo Where codramo IN(Select DISTINCT CodRamo From ra_horprof Where CodProf = '" & CodigoProfe & "' AND Ano >= year(getdate()) AND CodCarr IN(Select DISTINCT CodCarr From ra_horprof Where CodProf = '" & CodigoProfe & "' AND Ano >= year(getdate())))")
End Function

Function getCodRamosPorProfesor(CodigoProfe)
	Set getCodRamosPorProfesor = Session("Conn").Execute("Select DISTINCT CodRamo, CodCarr From ra_horprof Where CodProf = '" & CodigoProfe & "' AND Ano >= year(getdate()) AND CodCarr IN(Select DISTINCT CodCarr From ra_horprof Where CodProf = '" & CodigoProfe & "' AND Ano >= year(getdate()))")
End Function

Function getGUIDProfe(ThisProfe)
	Set rsProfe = Session("Conn").Execute("if exists(select * from ca_profes where rutprofe='"&ThisProfe&"') begin select GUID=1 end else begin select GUID=0 end")
	if cStr(rsProfe("GUID"))="0" _
		then
			Set rsProfe = Session("Conn").Execute("Insert Into ca_profes values('"&ThisProfe&"', newid())")
	end if 
	Set getGUIDProfe = Session("Conn").Execute("Select * from ca_Profes Where RutProfe='"&ThisProfe&"'")
End Function

Function insertNuevoMaterial(Id_Conter, CodigoProfe, CodigoCarrera, CodigoRamo, TipoDocumento, Referencia, DetalleDocumento, IdArchivo1, IdArchivo2, IdArchivo3, FechaUpLoad, YearPublicacion, Semestre, FechaActualizacion, Activo, UniqueProfe, FileSubidos)
	Set insertNuevoMaterial = Session("Conn").Execute("Insert Into ca_doctos Values('"&IdConter&"', '"&CodigoProfe&"', '"&CodigoCarrera&"', '"&CodigoRamo&"', '"&TipoDocumento&"', '"&Referencia&"', '"&DetalleDocumento&"', '"&IdArchivo1&"', '"&IdArchivo2&"', '"&IdArchivo3&"', '"&FechaUpLoad&"', '"&YearPublicacion&"', '"&Semestre&"', '"&FechaActualizacion&"', '"&Activo&"', '"&UniqueProfe&"')")
	Set insertNuevoMaterial = Nothing
End Function

Function getRamo(CodigoRamo)
	Set getRamo = Session("Conn").Execute("Select * From ra_ramo Where CodRamo='" & CodigoRamo & "'")
End Function

Function getCarrera(CodigoCarrera)
	Set getCarrera= Session("Conn").Execute("Select * From mt_carrer Where CodCarr='" & CodigoCarrera & "'")
End Function

Function DeleteDoctoPublicado(ThisDocument)
	Set rs = getDoctoDelProfe(ThisDocument)
	Set DeleteDoctoPublicado = Session("Conn").Execute("Delete From ca_Doctos Where IdConter='" & ThisDocument & "'")
	t=deleteFile("c:\InetPub\caENAC\doctos\" & rs("CodigoRamo") & "\" & rs("CodProfeFolder") & "\" & rs("AnoPublicacion") & "\S" & rs("PeriodoPublicacion") & "\" & rs("IdArchivo1"))
	t=deleteFile("c:\InetPub\caENAC\doctos\" & rs("CodigoRamo") & "\" & rs("CodProfeFolder") & "\" & rs("AnoPublicacion") & "\S" & rs("PeriodoPublicacion") & "\" & rs("IdArchivo2"))
	t=deleteFile("c:\InetPub\caENAC\doctos\" & rs("CodigoRamo") & "\" & rs("CodProfeFolder") & "\" & rs("AnoPublicacion") & "\S" & rs("PeriodoPublicacion") & "\" & rs("IdArchivo3"))
	Set rs = Nothing
	Set DeleteDoctoPublicado = Nothing
End Function

Function updateMaterial(ID, Referencia, Detalle, Activo)
	response.Write "Update ca_Doctos Set Referencia='" & Referencia & "', DetalleDocumento='" & Detalle & "', Activo='" & Activo & "' Where IdConter='" & ID & "<hr><hr>"
	Set updateMaterial = Session("Conn").Execute("Update ca_Doctos Set Referencia='" & Referencia & "', DetalleDocumento='" & Detalle & "', Activo='" & Activo & "' Where IdConter='" & ID & "'")
	Set updateMaterial = Nothing
End Function

Function SwitchApunte(ID, SwOnOff)
	Set SwitchApunte = Session("Conn").Execute("Update ca_Doctos Set Activo='" & SwOnOff & "' Where IdConter='" & ID & "'")
End Function

function ConterProfesRamos(CodigoRamo, Periodo, Codprof)
'response.Write(CodigoRamo + Periodo)
'response.End()
	str="Select Count(*) as Conter from ca_doctos where codramo='" & CodigoRamo & "' And PeriodoPublicacion='" & Periodo & "' And Activo='1' "
	if codprof <> "" then
		str = str & " and codprof ='"& Codprof &"'" 
	end if

	Set rs = Session("Conn").Execute(str) 
	ConterProfesRamos = rs("Conter")
end function

function getTodasLasCarreras()
	Set getTodasLasCarreras = Session("Conn").Execute("Select * From mt_carrer where CodCarr like '%I%' or CodCarr like '%C%' order by Nombre_C")
end function

Function getMailsRemitentes(Seccion)
	Set getMailsRemitentes = Session("Conn").Execute("Select * From ca_Mails Where Seccion='" & Seccion & "'")
End Function

Function getMails()
	Set getMails = Session("Conn").Execute("Select * From ca_Mails " & SqlSeccion())
End Function

Function getMail(ID)
	Set getMail = Session("Conn").Execute("Select * From ca_Mails Where ID='" & ID & "'")
End Function

Function InsertMail(caID, MailRemitente, NombreRemitente, ChkActivo)
	Set InsertMail = Session("Conn").Execute("Insert into ca_Mails(ID, Seccion, Mail, Nombre, IdUSer, Activo) Values(newid(), '" & Session("Seccion") & "', '" & MailRemitente & "', '" & NombreRemitente & "', '" & caID & "', '" & ChkActivo & "')")
End Function

Function UpdateMail(IDmsg, MailRemitente, NombreRemitente)
	Set UpdateMail = Session("Conn").Execute("Update ca_Mails Set Mail='" & MailRemitente & "', Nombre='" & NombreRemitente & "' Where ID='" & IDmsg & "'")
End Function

Function getMailsProfesores()
	Set getMailsProfesores = Session("Conn").Execute("Select CodProf, AP_PATER, AP_MATER, NOMBRES, MAIL From ra_profes where CodProf in (Select DISTINCT CodProf From ra_horprof) and mail <>''")
End Function

Function getMailsCoordinadores()
	Set getMailsCoordinadores = Session("Conn").Execute("Select CodProf, AP_PATER, AP_MATER, NOMBRES, MAIL From ra_profes Where Cargo='24' and Mail<>''")
End Function

function getAlumnosPorUnaCarrera(CodCarrera)
	Set getAlumnosPorUnaCarrera = Session("Conn").Execute("Select * From mt_client where CodCli in(Select Rut From mt_alumno Where EstAcad = 'VIGENTE' And CodCarPR='" & CodCarrera & "') And Mail<>''")
end function
%>