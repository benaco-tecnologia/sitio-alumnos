<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->

<%
if Session("CodCli") = "" then
  Response.Redirect("saltoinicio.htm")
end if
%>
<%
dim rut,clave,mensaje
rut=session("logrut")
clave=session("logclave")
%>
<html>
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<title>Portal de Alumnos en L&iacute;nea</title>
<body>
<table border="0" cellpadding="0" cellspacing="0"  align="left">
  <tr>
	<td>
      
      
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td colspan="2" valign="top" align="right"><% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
			    %></td>
			<td valign="top">
			<% CargarTop1()%><% SubMenu()%>
            
            
			<table width="100%" border="0" cellspacing="0" cellpadding="20" height="550" bgcolor="#FFFFFF">
			  <tr> 
				<td valign="top"><p><img src="Imagenes/titulos/T-correo-inst.gif" width="271" height="38"></p>
                
				  <table width="700" border="0" cellpadding="0" cellspacing="0">
                    
                    
                    
                    <tr>
            <td colspan="2" style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: normal; color: #000000">
                Los datos iniciales para acceder a su correo institucional son:
            </td>
        </tr>
        <tr><td colspan="2">&nbsp;</td></tr>
                <tr><td colspan="2">&nbsp;</td></tr>
         <tr>
            <td style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: normal; color: #000000" width="50">
                Usuario
            </td>
            <td align="left">
            <input name="TxtUsuario" type="text"  style="width:400px" class="casillas-form" id="TxtUsuario" value="<%=Session("mail_inst")%>" disabled="disabled">
            </td>
        </tr>
        <%
		strParame="SELECT dbo.Fn_ValorParame('MUESTRACLAVEMAILPORTALES')Parame"
		if BCL_ADO(strParame, rstParame) then
			MUESTRACLAVEMAILPORTALES=rstParame("Parame")
		else
			MUESTRACLAVEMAILPORTALES=""
		end if 
		
		if MUESTRACLAVEMAILPORTALES ="SI" then
		
		strClave="SELECT COALESCE(Clave_Mail_Inst,'')Clave FROM mt_client WHERE codcli='"& session("Rut") &"'"
		
		if BCL_ADO(strClave, rstClave) then
			Clave = rstClave("Clave")
		else
			Clave = ""
		end if 
		
		
		%>
         <tr>
            <td style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: normal; color: #000000">
                Clave&nbsp;Inicial
            </td>
            <td>
                            <input name="TxtClave" style="width:400px" type="text" class="casillas-form" id="TxtClave" value="<%=Clave%>" disabled="disabled">
            </td>
        </tr>
        <%end if%>
        
        <%
		strParame="SELECT coalesce(dbo.Fn_ValorParame('CORREOEXTERNOINSTITUCIONAL_PA'),'')Parame"
		if BCL_ADO(strParame, rstParame) then
			CORREO=rstParame("Parame")
		end if 
		
		if CORREO = "" then
			CORREO ="http://outlook.office365.com"
		end if 
		%>
         <tr><td colspan="2">&nbsp;</td></tr>
                 <tr><td colspan="2">&nbsp;</td></tr>
                         <tr><td colspan="2">&nbsp;</td></tr>
         <tr>
            <td colspan="2" style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: normal; color: #000000">
                Puedes acceder a tus correos haciendo click
                <a href="<%=CORREO%>" target="_blank">aqu&iacute;</a>
 o escribiendo la siguiente
            </td>
        </tr>
                <tr><td colspan="2">&nbsp;</td></tr>
         <tr>
            <td colspan="2" style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: normal; color: #000000">
                direcci&oacute;n en tu navegador: <%=CORREO%>
            </td>
        </tr>
        <tr>
            <td colspan="2" style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: normal; color: #000000">&nbsp;
            </td>
        </tr>
        <tr>
            <td colspan="2" style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: normal; color: #000000">&nbsp;
            </td>
        </tr>
        <tr>
            <td colspan="2" style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: normal; color: #000000">&nbsp;
            </td>
        </tr>
        <tr>
            <td colspan="2" style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: normal; color: #000000">&nbsp;
            </td>
        </tr>
		<%if CORREO ="http://outlook.office365.com" then%>
            <tr>
                <td  valign="top" align="left" colspan="2" style="font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: normal; color: #000000">
                    <a target="_blank" href="http://outlook.office365.com/"><IMG src="Imagenes/office365.png" border="0"  name="office" width="208" height="112"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <a target="_blank" href="http://moobs.azurewebsites.net"><IMG src="Imagenes/moodle.png" border="0"  name="moodle" width="208" height="112"></a>
                </td>
            </tr>
        <%end if%>              
                      </table>
                      
                      </td></tr></table></td></tr></table></body></html>