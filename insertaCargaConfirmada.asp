<!--#INCLUDE FILE="include/conexion.inc" -->
<%  
	if Session("CodCli") = "" then
	  Response.Redirect("saltoinicio.htm")
	  response.End()
	end if

	CodCli = Session("CodCli")
	Ano=session("ano")
	Periodo=session("periodo")
			
	StrSql = "sp_ConfirmacionInscripcionWebTRA '" & trim(Codcli) & "'," & Ano & "," & Periodo & ""					
	Set Rst = Session("Conn").Execute(StrSql)

	if not Rst.Eof then
		Confirma = Rst(0)			
	End if
		
	if Confirma =1 then
		str="update SIS_REG_INSCRIPCION set Acepto='SI',fechaAcepto=getdate(),tipo='PORTAL' where codcli='" & CodCli &"' and ano ='" & Ano &"' and periodo='" & Periodo & "'"
		Session("Conn").execute(str)
		response.redirect("resultadotr.asp")
	else
		response.redirect("resultadotr.asp")
	end if 
%>