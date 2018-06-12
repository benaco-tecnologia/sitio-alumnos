<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	response.buffer = TRUE 
    Response.Expires = -1
	
	If Request.QueryString("usrcod") = "" or Request.ServerVariables("HTTP_REFERER") = "" Then
		Response.Clear()
		Response.End()
	Else
		Session("rut") = Request.QueryString("usrcod")
		Session("checado") = "1"
	End If
	
	Response.Redirect("./")
	Response.End()
%>