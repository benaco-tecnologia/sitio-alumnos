<%
Function BorrarArchivo(archivo) 
    dim fs 
    Set fs = Server.CreateObject("Scripting.FileSystemObject") 
    if fs.FileExists(archivo) then 
        fs.DeleteFile(archivo) 
        BorrarArchivo = true 
    else 
        BorrarArchivo = false 
    end if 
    Set fs = Nothing 
End function

Function SecureSQLex(strVar) 
      dim banned, final, i 
      banned = array("select", "drop", ";", "--", "insert","delete", "xp_") 
      for i = 0 to uBound(banned) 
            str = replace(str, banned(i), "") 
      next 
      final = replace(str,"'","''") 
      secureSQLex= final 
End Function

Function EsEmail(email) 
   dim ubicacion_arroba 
   EsEmail = true 
   If Len(email) < 5 then EsEmail = false 
   If InStr(email, "@") = 0 then EsEmail = false 
   If InStr(email, ".") = 0 then EsEmail = false 
   ubicacion_arroba = InStr(email, "@") 
   If (ubicacion_arroba <> 0) and (InStr(email, ".") <> 0) then 
      If InStr(email, "@") > InStr(ubicacion_arroba, email, ".") then EsEmail = false 
   End If 
end function

Function Encripta(sPasswd)
    Dim fc_b1, fc_b2, fc_b3, fc_b4, fc_b5, fc_password
    Dim fn_b1, fn_b2, fn_b3, fn_b4, fn_b5

    If Len(sPasswd) < 5 Then
      sPasswd = Space(5 - Len(sPasswd)) + sPasswd
    End If
    fc_b1 = Mid(sPasswd, 1, 1)
    fc_b2 = Mid(sPasswd, 2, 1)
    fc_b3 = Mid(sPasswd, 3, 1)
    fc_b4 = Mid(sPasswd, 4, 1)
    fc_b5 = Mid(sPasswd, 5, 1)
    fn_b1 = Asc(fc_b1)
    fn_b2 = Asc(fc_b2)
    fn_b3 = Asc(fc_b3)
    fn_b4 = Asc(fc_b4)
    fn_b5 = Asc(fc_b5)
    fn_b1 = fn_b1 + 3
    fn_b2 = fn_b2 + 5
    fn_b3 = fn_b3 + 2
    fn_b4 = fn_b4 + 4
    fn_b5 = fn_b5 + 1
    fc_b1 = Chr(fn_b1)
    fc_b2 = Chr(fn_b2)
    fc_b3 = Chr(fn_b3)
    fc_b4 = Chr(fn_b4)
    fc_b5 = Chr(fn_b5)
    Encripta = fc_b1 & fc_b4 & fc_b5 & fc_b3 & fc_b2
End Function

Function Desencripta(sPasswd)
    Dim fc_b1, fc_b2, fc_b3, fc_b4, fc_b5, fc_password
    Dim fn_b1, fn_b2, fn_b3, fn_b4, fn_b5

    fc_b1 = Mid(sPasswd, 1, 1)
    fc_b4 = Mid(sPasswd, 2, 1)
    fc_b5 = Mid(sPasswd, 3, 1)
    fc_b3 = Mid(sPasswd, 4, 1)
    fc_b2 = Mid(sPasswd, 5, 1)
    fn_b1 = Asc(fc_b1)
    fn_b2 = Asc(fc_b2)
    fn_b3 = Asc(fc_b3)
    fn_b4 = Asc(fc_b4)
    fn_b5 = Asc(fc_b5)
    fn_b1 = fn_b1 - 3
    fn_b2 = fn_b2 - 5
    fn_b3 = fn_b3 - 2
    fn_b4 = fn_b4 - 4
    fn_b5 = fn_b5 - 1
    fc_b1 = Chr(fn_b1)
    fc_b2 = Chr(fn_b2)
    fc_b3 = Chr(fn_b3)
    fc_b4 = Chr(fn_b4)
    fc_b5 = Chr(fn_b5)
    Desencripta = fc_b1 & fc_b2 & fc_b3 & fc_b4 & fc_b5
End Function

Function EliminaPalabras(Texto)
    Dim Eliminar, NuevoStr
	
	Eliminar = Array("select","drop",";","delete","insert","update","xp_")
	NuevoStr = Texto
	
	For i = 1 to UBound(Eliminar)
		NuevoStr = Replace(NuevoStr, Eliminar(i), "")
	Next
	
	EliminaPalabras = NuevoStr	
End Function

Function RutSinDV(RutIngresado)
	 Dim RutAux, RutAuxSinDV, i, c
	 
	 RutAux = Mid(RutIngresado,1,Len(RutIngresado) - 1)
	 
	 For i = 1 To Len(RutAux)
	    c = Mid(RutAux,i,1)
	    If c <> "." And c <> "-" Then
	        RutAuxSinDV = RutAuxSinDV & c
	    End If 	
	 Next
	 
	 RutSinDV = RutAuxSinDV
	 
End Function

Function GetNombreRamo(CodRamo)
  Dim StrSql
  Dim Rst

  StrSql = "Select Nombre from ra_ramo where codRamo = '" & CodRamo & "' "

  Set Rst = Session("Conn").Execute(StrSql)

  if not Rst.eof then
    GetNombreRamo = Rst("Nombre")
  else
    GetNombreRamo = "Sin Nombre"
  end if
End function

%>