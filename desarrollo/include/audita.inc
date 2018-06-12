<%

function RegAuditoria(Codcli, Opcion, Transaccion)
Dim StrSql
Dim Rst

    StrSql = "insert into ra_audita (nivel,username,opcion,acceso_opcion,transaccion,horfec_trans)" 
    StrSql = StrSql & " values(3,'" & Trim(Codcli) & "','" & Trim(Opcion) & "',getdate() ,'" & Trim(Transaccion) & "',getdate() )"

    Conn.execute StrSql, rdexecDirect
end function

%>