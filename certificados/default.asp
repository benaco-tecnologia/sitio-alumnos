<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	Response.Buffer = TRUE 
    Response.Expires = -1
	
	Response.Redirect("./validacion.asp" )
	' Response.Redirect("../)
	Response.Clear()
	Response.End()
%>