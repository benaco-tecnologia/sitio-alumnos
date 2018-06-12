<%
Server.ScriptTimeOut = 99999
Dim Logueado

Logueado = Session("Logueado")
If Len(Logueado) = 0 Or Logueado = Null Then Logueado = "0"

If Logueado = "1" Then 
   'Href = "adm-acad.asp"
   Href = "acceso.asp"
Else
   'Href = "adm-log.htm"
   Response.Redirect("alumnos.asp")
   
End If
%>
<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if
%>


<html>
<head>
<title>Inscripción de Asignaturas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<frameset cols="185,*" frameborder="NO" border="0" framespacing="0" rows="*" <% if Logueado = "1" then %>onUnload="javascript:window.open('VerifInscrip.asp')"<%end if%>> 
  <frame name="leftFrame" scrolling="NO" noresize src="f-izq.asp" frameborder="NO">
  <frame name="mainFrame" src="alumn-udd.html" frameborder="NO" marginwidth="0" marginheight="0" scrolling="yes" noresize>
</frameset>
<noframes> 
<body bgcolor="#FFFFFF" text="#000000">
</body>
</noframes> 
</html>
