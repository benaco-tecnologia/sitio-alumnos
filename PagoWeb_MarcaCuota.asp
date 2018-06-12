<%  response.buffer = false 
    Response.Expires = -1 
    
	
	Dim Sql
	Dim IdPagoWeb, Id_Cuota
	
	IDPagoWeb = Request("IDPAGOWEB")
	Id_Cuota = Request("ID_CUOTA")
	
	'incluye o excluye la cuota selecciona de la intención de Pago en que estamos trabajando.
	Sql = "Exec pr_Pago_Web_Marcar_Cuota @IdPagoWeb = " & IdPagoWeb & ", @Id_Cuota = " & (Id_Cuota)  
	'response.Write(Sql)
	'response.End()
	Session("Conn").Execute Sql
   	
	Dim Url
	'Url = "Pago-Web-Cuotas.asp?IDPAGOWEB=" + IdPagoWeb
	Url = "Pago-Web-Cuotas.asp"
	response.Redirect(Url)
       
%>