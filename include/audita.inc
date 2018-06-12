<%

function RegAuditoria(Codcli, Opcion, Transaccion)
Dim StrSql
Dim Rst
    'StrSql = "insert into ra_audita (nivel,username,opcion,acceso_opcion,transaccion)" 
    'StrSql = StrSql & " values(3,'" & Trim(Codcli) & "','" & Trim(Opcion) & "',getdate() ,'" & Trim(Transaccion) & "',getdate() )"
	

	StrSql = "INSERT INTO RA_AUDITA (	USERNAME,	CODCLI,	OPCION,	FECHA,	TRANSACCION)" 
    StrSql = StrSql & " VALUES('" &Session("id_usuario")& "','" & Trim(Codcli) & "','" & Trim(Opcion) & "',getdate() ,'" & Trim(Transaccion) & "')"
    Session("Conn").execute StrSql, rdexecDirect
end function

%>