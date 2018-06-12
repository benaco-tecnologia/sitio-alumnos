<%
Dim objMail
Set objMail = Server.CreateObject("CDONTS.NewMail")
objMail.from="mbartsch@netglobalis.net"
objMail.Subject = "TEST CDONTS"
objMail.To = "mikepre@hotmail.com"
objMail.BodyFormat = CdoBodyFormatHTML
objMail.MailFormat = CdoMailFormatMime
objMail.Body = "prueba usando cdonts con http://200.2.201.40/udd_prueba/mbr.asp"
objMail.send
Response.write("Sent1ba")
set objMail = nothing
%>
<P>
