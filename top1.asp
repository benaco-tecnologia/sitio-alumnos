<%
function CargarTop1()

	html="<table width='704' height='0' border='0' cellpadding='0' cellspacing='0'>"
	
	strPar = "select dbo.Fn_ValorParame('USAMAIL') as valor"
	set rsPar = Session("Conn").execute(strPar)
	
	if (rsPar("VALOR") = "S" or rsPar("VALOR") = "SI") then
	    html=html & "<td valign='top'><img name='A_r1_c4' src='imagenes/A-_r1_c4.jpg' width='939' height='52' border='0' alt=''></td>"
	else
		html=html & "<td valign='top'><img name='A_r1_c4' src='imagenes/A-_r1_c4.jpg' width='826' height='52' border='0' alt=''></td>"
	end if 
    html=html & "</table>"
	response.Write(html)
end function
%>
