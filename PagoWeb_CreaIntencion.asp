<!--#INCLUDE file="analytics.asp" -->

<%
strParame="SELECT coalesce(dbo.Fn_ValorParame('USAANALYTICS'),'NO')Parame"
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
	USAANALYTICS=rstParame("Parame")
end if

if USAANALYTICS="SI" then
	analitycs() 
end if 
%>

<%
    'response.buffer = false 
    'Response.Expires = -1 

'RTF / 07.06.2011
'Da origen al proceso de pago x ServiPag
'Se parte creando un número de intención de pago el que es almacenado junto con los datos del RUT 
'luego se asociarán las cuota a cancelar y posteriormente se envía a Servipag.



if session("codcli") = null or session("Rut") = null then 
    response.Redirect("saltoinicio.htm")    
end if


dim Rut, miRst, Sql

Rut = Session("Rut")

'response.Write("1" + Rut + "*")


Set miRst = Nothing

Sql = "Exec Pr_PagoWeb_Crea_Intencion @Rut = " + Rut
'response.Write(Sql)

dim IDPAGOWEB 

Set miRst = Session("Conn").Execute(Sql)


Sql = "Select Convert(Varchar, Coalesce(@@Identity, 0)) as IDPAGOWEB "
Set miRst = Session("Conn").Execute(Sql)

   IDPAGOWEB = miRst("IDPAGOWEB")
   
Set miRst = Nothing

Sql = " INSERT INTO mt_pago_web (id, rut)  SELECT id, rut FROM mt_pago_web_corr WHERE id = " & IDPAGOWEB
Session("Conn").Execute Sql
    
        
   session("IDPAGOWEBBP") = IDPAGOWEB
    Dim Url
    'Url = "Pago-Web-Cuotas.asp?IDPAGOWEB=" + IDPAGOWEB
	Url = "Pago-Web-Cuotas.asp"
        
   response.Redirect(Url)


 %>
 <!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->