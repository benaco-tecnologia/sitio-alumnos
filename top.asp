<%
function CargarTop()
	html="<table width='198' height='136' border='0' cellpadding='0' cellspacing='0' aling ='right'>"
	html=html & "<tr valing = 'top'>"
	html=html & "<td valign='top'><img name='A_r1_c1' src='imagenes/2_r1_c1.jpg' width='18'height='136' border='0' alt=''></td>"
	html=html & "<td valign='top' ><img name='A_r1_c2' src='imagenes/2_r1_c2.jpg' width='162' height='136' border='0' alt=''></td>"
	html=html & "<td valign='top'><img name='A_r1_c3' src='imagenes/2_r1_c3.jpg' width='18' height='136' border='0' alt=''></td>"
	html=html & "</tr>"
	html=html & "</table>"
	response.Write(html)
end function
%>