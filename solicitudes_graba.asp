<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->

<%
	Dim mySmartUpload
	Dim i, f, Archivos(2),ArchivosOriginal(2)

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set mySmartUpload = Server.CreateObject("aspSmartUpload.SmartUpload")
	mySmartUpload.Upload  

	strParame="SELECT dbo.Fn_ValorParame('rutaIngSolicitudes')Parame"
	set rstParame= Session("Conn").Execute(strParame)		
	if not rstParame.eof then
			Raiz=rstParame("Parame")
		else
			Raiz=""
	end if

	'//Sube Archivos
	'intCount = mySmartUpload.Save(Raiz)

	SecuenciaSolicitud = GetSecuenciaSolicitud
	
	'Sube Archivos
	i = 0
	For each File In mySmartUpload.Files
		if Len(File.Filename) > 0 then

			Archivos(i) =  Cstr(left(replace(cstr(Time()),":",""),4)) + Cstr(mySmartUpload.Form("txtSolicitud")) + "0" + Cstr(SecuenciaSolicitud) + "_" + File.Filename
			mySmartUpload.Files("Docto" & i+1).SaveAs(Raiz & Archivos(i))
			i = i + 1
		end if
	next	
	
	'InsertaSolicitud	
	Str = "pa08_RA_MOVRES_in_2 @NUMSOLI =0, "
	Str = Str & "@TIPOSOLI='"& mySmartUpload.Form("txtSolicitud") &"', "
	Str = Str & "@CODROL='"& Session("perfilNW") & "', "
	Str = Str & "@USUARIO='"& SacaUsuario() &"', "
	Str = Str & "@FEC_ASIG='"& mySmartUpload.Form("txtFecha") &"', "
	Str = Str & "@ESTADO='PENDIENTE', "
	Str = Str & "@GLOSA= 'Alumno: " & session("codcli") &" ## "& LimpiarTexto(mySmartUpload.Form("txtGlosa")) &"', "
	Str = Str & "@CODCLI='"& session("codcli") &"', "
	Str = Str & "@CODROLDEST='"& GetRolDestinoSolicitud &"', "
	Str = Str & "@GLOSAADICIONAL='', "
	Str = Str & "@MOTIVO='"& LimpiarTexto(mySmartUpload.Form("txtMotivo")) &"', "
	Str = Str & "@CODCARR='"& session("codcarr") &"', "
	Str = Str & "@SECUENCIA="& SecuenciaSolicitud &", " 
	Str = Str & "@CODSEDE='"& session("codsede") &"', "
	Str = Str & "@ESTADOANT='---------------------', "

	if mySmartUpload.Form("txtCarreraDes") = "" or mySmartUpload.Form("txtCarreraDes") = "0" then
		Str = Str & "@CODCARRDEST='"& session("codcarr") &"', "
	else
		Str = Str & "@CODCARRDEST='"& mySmartUpload.Form("txtCarreraDes") &"', "
	end if 
	
	if mySmartUpload.Form("txtCodsedeDes") = "" or mySmartUpload.Form("txtCodsedeDes") = "0"then	
		Str = Str & "@SEDEDEST='"& session("codsede") &"', " 
	else
		Str = Str & "@SEDEDEST='"& mySmartUpload.Form("txtCodsedeDes") &"', "
	end if 
	
	Str = Str & "@adjunto1='"& Archivos(0) &"', "
	Str = Str & "@adjunto2='"& Archivos(1) &"', "
	Str = Str & "@adjunto3='"& Archivos(2) &"'  " 
	
	Session("Conn").Execute(Str) 

	response.Redirect("ingresodesolicitudes.asp?M=OK")
	response.End()
	
Function GetRolDestinoSolicitud()
	dim str 
	dim Rst

	str="pa08_RA_DETRES_sel @TIPOSOLI= '"& mySmartUpload.Form("txtSolicitud") &"', @CODROL = '"& Session("perfilNW") & "' "
	set Rst= Session("Conn").Execute(str)
	
	if not Rst.eof then
		GetRolDestinoSolicitud=Rst("CODROLDEST")
	else
		GetRolDestinoSolicitud="0"
	end if

	Rst.close()

End Function 

Function GetSecuenciaSolicitud()
	dim str 
	dim Rst

	sec = 0
	GetSecuenciaSolicitud = 0
	
	str="pa08_RA_DETRES_sel @TIPOSOLI= '"& mySmartUpload.Form("txtSolicitud") &"', @CODROL = '"& Session("perfilNW") & "' "
	
	set Rst= Session("Conn").Execute(str)
	
	While Not Rst.Eof
	
		if GetSecuenciaSolicitud = 0 then
			secuencia = valnulo(Rst("SECUENCIA"),NUM_)
			if secuencia >= sec then
				GetSecuenciaSolicitud = secuencia
			end if 
		end if 
		
		Rst.movenext
	wend
	
	Rst.close()
End Function 
	
%>