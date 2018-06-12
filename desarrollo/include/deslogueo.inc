<%
If Len(Session("RutAlumn")) = 0 Or Session("RutAlumn") = "" Then
   Session("RutAlumn") = "0"	
%>
<script>window.top.location.href='alumn-udd.asp';</script>
<%
  Response.End()
End If
%>