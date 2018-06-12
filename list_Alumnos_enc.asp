<%
    Response.Expires = -1
    Session.Timeout = 999
    Response.Buffer = false
    Server.ScriptTimeout = 99999
%>
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<% 
  	session("carrera_c")= request("carr_corto")
    session("carrera_l")= request("carr")
   	session("sede")=request("sede")
	
		select case 	session("sede")
				case "01":
						sede_="CONCEPCION"
				case "02":
						sede_="PUERTO MONTT"
				case "03":
						sede_="TALCAHUANO"		
				case "04":
						sede_="OSORNO"						
				case "05":
						sede_="VALDIVIA"						
				case "06":
						sede_="SANTIAGO"
			 end select
			 	
	mescla=session("sede")+ "" +session("carrera_c")
	
	
	
	dim nombre (10000)
	dim codcli (10000)
	dim rut (10000)
	dim ncarr (10000)
	if (mescla<>"") then 
		'strsql1= "select codcli,count (evaluado) as pesoT from ra_encuestados where codcli like '%" & mescla &"%'  and codcli in (select distinct codcli  from ra_encuestados ) group by codcli"
		strsql1= "select enc.codcli,count (enc.evaluado) as pesoT , alu.codcarpr, cli.paterno, cli.materno, cli.nombre, cli.codcli as rut,cli.dig  from ra_encuestados as enc, mt_alumno as alu, mt_client as cli where alu.codcli=enc.codcli and cli.codcli=alu.rut and enc.codcli like '%" & mescla &"%' and enc.codcli in (select distinct dos.codcli  from ra_encuestados as dos ) group by enc.codcli, alu.codcarpr, cli.paterno, cli.materno, cli.nombre,cli.codcli,cli.dig"
		'response.Write(strsql1)
		'response.End()
		i=0
		if bcl_ado(strsql1,rst2) then
			 Do While not rst2.eof 	     
			 	strsql= "select count (evaluado) as pesoS  from ra_encuestados with(nolock) where evaluado='S' and codcli ='"+ rst2("codcli")  +"'" 
				rst1=Session("Conn").execute(strsql)
				if (rst2("pesoT")=rst1("pesoS")) then 
					nombre(i)=rst2("nombre")+" "+rst2("paterno")+" "+rst2("materno")
					codcli(i)=rst2("codcli")
					rut(i)=rst2("rut")+"-"+rst2("dig")
					ncarr(i)=rst2("codcarpr")
					i=i+1
				end if 		
		
		         rst2.movenext
			 Loop    	
		 end if
	end if
 %>
<html>
<head>

<title>Evaluaci&oacute;n Docente</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="ed.css" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="0">
<!--#INCLUDE file="include/desconexion.inc" -->

<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td colspan="5" class="titulo2"> 
      <% response.Write(session("carrera_l") + " / " + sede_)%>
    </td>
  </tr>
  <tr> 
    <td colspan="5" height="1" bgcolor="#CCCCCC"></td>
  </tr>
  <tr align="right"> 
    <td colspan="5"> 
      <table width="89" height="27" border="0" cellpadding="0" cellspacing="0">
        <tr>
          <td align="center" background="/imag/imagenes/boton.jpg" ><a href="export.asp" class="tex"><b>Bajar 
            a Excel</b></a></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan="5">&nbsp; </td>
  </tr>
  <tr class="resaltar-pregunta"> 
    <td><strong>Nº</strong></td>
    <td><strong>NOMBRE</strong></td>
    <td><strong>RUT</strong></td>
    <!--td><strong>CARRERA</strong></td-->
    <td><strong>CÓDIGO DE ALUMNO</strong></td>
  </tr>
  <%
	if(i<>0) then 
		 for k=0 to i -1
		 
		 	
		 %>
  <tr class="tex"> 
    <td class="tex"> 
      <% response.Write(k+1)%>
    </td>
    <td class="tex"> 
      <% response.Write(nombre(k)) %>
    </td>
    <td class="tex"> 
      <% response.Write(rut(k)) %>
    </td>
    <!--td class="tex"> 
      <%' response.Write(ncarr(k)) %>
    </td-->
    <td class="tex"> 
      <% response.Write(codcli(k)) %>
    </td>
  </tr>
  <% 		  next
	else
		response.Write("<tr><td class='tex' >&nbsp;No hay encuestas evaluadas en esta carrera</td></tr>")
	end if
 %>
</table>
</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->