<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	Response.Buffer = TRUE 
    Response.Expires = -1
	
	If Request.QueryString("usrcod") = "" Then
		Response.Redirect("../v2/error.html")
		Response.Clear()
		Response.End()
	End If
	
	Response.Redirect("generar.asp")
	Response.Clear()
	Response.End()
%>