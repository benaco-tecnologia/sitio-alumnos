<%
function guarda_log(rut, opcion,transaccion, nivel)
		
		fecha_=now()
		set rs_a = Server.CreateObject("ADODB.recordset")
		sql_a="insert into RA_AUDITA (USERNAME, OPCION, HORFEC_TRANS, ACCESO_OPCION, TRANSACCION, NIVEL, FECMOVTRANS) VALUES ('"& rut &"', '"& opcion &"', '"& fecha_ &"', '"& fecha_ &"','"& transaccion &"', '"& nivel &"', '"& fecha_&"')"
		'response.Write(sql_a)
		'response.End()
		
		rs_a.Open sql_a, conn
		
     
	
		guarda_log="S"
end function
%>
