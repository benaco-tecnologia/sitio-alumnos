<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<OBJECT name=Update width="32" height="33" id=Update RUNAT=server PROGID=ADODB.Recordset>
</OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=Actualiza name=Actualiza></OBJECT>
<%
Dim Rut, ClaveActual,NuevaClave, strSql, CambioClave,codapod,rutalumno,mystr,rst,clv,rclv
CambioClave = Request("CambioClave")
codapod=request("clave")
clave=request("sw")
CarreraAlumno=session("CarreraAlumno")
				
	mystr = "select dbo.Encrypt('"&Request("Clave")&"') as clave"
	Set rst = Session("Conn").execute(mystr)
	ClaveActual = Rst("clave")	
	strSql = "Select us_password from ca_usuarios where us_consuser = '" & Session("usrNW") & "'"
	Update.Open strSql, Session("Conn")

	If Trim(Update("us_password")) <> Trim(ClaveActual) Then
		Update.Close()
		Response.Redirect "atencion.asp?Error=2&Logueo=2" ' Clave invalida
		Response.End()
	else 
		clv = "select dbo.Encrypt('"&Request("NuevaClave")&"') as clave"
		Set rclv = Session("Conn").execute(clv)
		NuevaClave = rclv("clave") 
		
		strSql = "Update ca_usuarios set us_password = '" & NuevaClave & "',us_intentosfallidos = 0,us_feccambiopass = getdate() Where  us_consuser = '" & Session("usrNW") & "' and us_password = '"&Update("us_password") &"'"
		Actualiza.Open strSql,Session("Conn")
		
		'inicio nuevo actualiza tabla cambio password
		strICP = "INSERT INTO dbo.ca_usuariospasswords(id_usuario,up_feccambiopass,us_passwordold,us_passwordnew,id_personacambio) "
		strICP = strICP & "VALUES  ('"& session("cc_id_usuario") &"',GETDATE(),'"& Update("us_password") &"','" & NuevaClave & "','"& session("cc_id_persona") &"')"
		Session("Conn").Execute(strICP)
		'fin nuevo
		
		response.Redirect("CambioPWOk.asp")
		response.End()
		session("entrada")=""
	end if
	
	Update.Close()
%>
<!--#INCLUDE file="include/desconexion.inc" -->
<script>
function Uno()
{
alert("No Existe Codigo de Apoderado") 
}
</script>
