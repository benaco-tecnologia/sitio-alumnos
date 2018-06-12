<%
function guarda_log(user, CODCLI, opcion,transaccion)
		
		'fecha_=now()
		set rs_a = Server.CreateObject("ADODB.recordset")
		'sql_a="insert into RA_AUDITA (USERNAME, OPCION, HORFEC_TRANS, ACCESO_OPCION, TRANSACCION, NIVEL, FECMOVTRANS) VALUES ('"& rut &"', '"& opcion &"', '"& fecha_ &"', '"& fecha_ &"','"& transaccion &"', '"& nivel &"', '"& fecha_&"')"
		'response.Write(sql_a)
		'response.End()



sql_a="INSERT INTO RA_AUDITA (USERNAME,CODCLI,OPCION,FECHA,TRANSACCION) VALUES ( '"& user &"','"& CODCLI &"', '"& opcion &"',GETDATE(),'"& transaccion &"') "


		'response.Write(sql_a)
		'response.End()	
		rs_a.Open sql_a, Session("Conn")
		
     
	
		guarda_log="S"
end function
%>
