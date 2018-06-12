<%
    Response.Expires = -1
    Session.Timeout = 999
    Response.Buffer = false
    Server.ScriptTimeout = 99999
%>
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%    
'Obtengo numero de encuesta docente vigente		
sqlenc="select numero from RA_ENCUESTA Where CONVERT(DATE,getDate())>=FecIni And CONVERT(DATE,getDate()) <=FecFin AND TIPO_ENCUESTA=1"		
if bcl_ado(sqlenc,rstenc) then
	numeroEnc=valnulo(rstenc("numero"),num_) 
end if 

	  'Listado de asignaturas 	   	  
	  Strsql ="sp_listaRamosEncDoc '" & session("codcli") & "','" & session("Codsede") & "',"& session("Ano_Ed") &","& session("Periodo_Ed") &","& session("NumeroMaximoPrueba") &",'" & session("Cerrada") & "'" 
'response.Write(Strsql)
	  if bcl_ado(strsql,rs) then
	
			function ValidaRamos(Ramo,Secc)     
	
			  ValidaRamos=True
	
			  sql="Select * from ra_encuestados a,ra_encuesta b where a.Ano=" & session("Ano_Ed") & ""
			  'sql=sql & " and Periodo='" & session("Periodo_Ed") & "' and CodCli='" & session("CodCli") & "'"				  
			  sql=sql & "  and a.CodCli='" & session("CodCli") & "'"
			  sql=sql & " and a.CodRamo='" & Ramo & "' and a.CodSecc=" & valnulo(Secc,num_) & " and a.Evaluado='NO'"  
			  sql=sql & " AND a.encuesta = b.NUMERO AND a.ano = b.ano AND a.periodo = b.periodo AND b.TIPO_ENCUESTA = 1 and b.numero='"& numeroEnc &"'" 
			  'response.Write(sql)
			  if bcl_ado(sql,rst) then				  
				 ValidaRamos=False
			  end if    	  	  	  
		   end function
		   
		   
		  function ValidaEncuesta(Ramo,Secc,ano,periodo)     
	
			  ValidaEncuesta=True
	
			  sqlq="Select * from ra_encuestados a,ra_encuesta b where  a.Ano=" & ano & ""
			  sqlq=sqlq & " and a.Periodo='" & periodo & "'"				  
			  sqlq=sqlq & "  and a.CodCli='" & session("CodCli") & "'"
			  sqlq=sqlq & " and a.CodRamo='" & Ramo & "' and a.CodSecc=" & valnulo(Secc,num_) & " and a.Evaluado='NO'"  
			  sqlq=sqlq & " AND a.encuesta = b.NUMERO AND b.TIPO_ENCUESTA = 1 and b.numero='"& numeroEnc &"' "
			  if bcl_ado(sqlq,rstq) then
			  
				 ValidaEncuesta=false
			  end if    	  	  	  
		   end function
		end if
	
		sql="delete from ra_encuestados where CodCli='" & session("CodCli") & "'"
		sql=sql & "  and Evaluado='NO' and encuesta='"& numeroEnc &"'"  
		Session("Conn").execute(sql)	
		
	   '---NUEVO INSERTA A PARTIR DEL SELECT DE ARRIBA   		
		JSql ="sp_GrabaRamosEncDoc '" & session("codcli") & "','" & session("Codsede") & "',"& session("Ano_Ed") &","& session("Periodo_Ed") &","& session("NumeroMaximoPrueba") &",'" & session("Cerrada") & "'"
		'response.Write(JSql)
		Session("Conn").execute(JSql)      	
 %>
<html>
<head>
<title>Evaluaci&oacute;n Docente</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="ed.css" type="text/css">
<link href="css/tablas.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="0">
<span id="lblTituloAsignaturas" style="font-size: 25px"  class="text-menu">Asignaturas Inscritas</span> 
<table width="100%" border="0" cellspacing="1" cellpadding="0" bordercolor="#000000">
  <tr bgcolor="#000000" class="text-cabecera-celda"> 
    <td width="24%" height="20" background="Imagenes/fdo-cabecera-cel22.jpg" bgcolor="#FFFFFF"> 
      <p align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">C</font><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">&oacute;d.</font>        </p>
    </td>
    <td width="50%" height="20" background="Imagenes/fdo-cabecera-cel22.jpg" bgcolor="#FFFFFF"> 
      <p align="center"><font id="lblColumnAsign" face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Asignatura</font></p>
    </td>    
    <td width="16%" height="20" background="Imagenes/fdo-cabecera-cel22.jpg" bgcolor="#FFFFFF"> 
      <p align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">S</font><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">ecci&oacute;n</font></p>
    </td>
    <td width="10%" height="20" background="Imagenes/fdo-cabecera-cel22.jpg" bgcolor="#FFFFFF"> 
     <p align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1" color="#FFFFFF">Semestre</font></p>
    </td>
  </tr>
  <%if not rs.eof then  
       ColorFila="#F0F0DD"          
	  do while not rs.eof
	   if ValidaEncuesta(rs("RAMOEQUIV"),rs("codsecc"),rs("ANO"),rs("PERIODO")) then
	      ColorFila="#FF6600"   		     
	   end if
	   %>
  <tr>     
    <td width="24%" bgcolor="<%=ColorFila%>" height="20">&nbsp;<font face="Arial, Helvetica, sans-serif" size="1"><a href="javascript:PasarValores('<%=rs("RAMOEQUIV")%>','<%=rs("codsecc")%>','<%=rs("PERIODO")%>');"><%=rs("RAMOEQUIV")%> </a></font></td>
    <td width="50%" bgcolor="<%=ColorFila%>" height="20">&nbsp;<font face="Arial, Helvetica, sans-serif" size="1"><%=rs("ASIGNATURA")%></font></td>
    <td width="16%" bgcolor="<%=ColorFila%>" height="20">&nbsp;<font face="Arial, Helvetica, sans-serif" size="1"><%=rs("CODSECC")& " - " & rs("TIPOCURSO")%></font></td>
    <td width="10%" bgcolor="<%=ColorFila%>" height="20">&nbsp;<font face="Arial, Helvetica, sans-serif" size="1"><%=rs("ANO") & " - " & rs("PERIODO")%></font></td>
  </tr>
  <% 
     'if ColorFila="#F0F0DD" then
	 '   ColorFila= "##DBECF2"
	 'else
	   	ColorFila="#F0F0DD" 
     'end if
	  rs.movenext		
	   loop
	   else %>
  <tr> 
    <td width="24%" bgcolor="#DBECF2" height="20">&nbsp;</td>
    <td width="50%" bgcolor="#DBECF2" height="20">&nbsp;</td>
    <td width="16%" bgcolor="#DBECF2" height="20">&nbsp;</td>
    <td width="10%" bgcolor="#DBECF2" height="20">&nbsp;</td>
  </tr>
   <%end if%>
  <!--<tr> 
    <td width="15%" bgcolor="#F0F0DD" height="10">&nbsp;</td>
    <td width="50%" bgcolor="#F0F0DD" height="10">&nbsp;</td>
    <td width="15%" bgcolor="#F0F0DD" height="10">&nbsp;</td>
    <td width="20%" bgcolor="#F0F0DD" height="10">&nbsp;</td>
  </tr>
  <tr> 
    <td width="15%" bgcolor="#F7FFEE" height="10">&nbsp;</td>
    <td width="50%" bgcolor="#F7FFEE" height="10">&nbsp;</td>
    <td width="15%" bgcolor="#F7FFEE" height="10">&nbsp;</td>
    <td width="20%" bgcolor="#F7FFEE" height="10">&nbsp;</td>
  </tr>-->
</table>
<script>
function PasarValores(codramo,codsecc,periodo) {
parent.mainProf.location.href="profesores.asp?CodRamo=" + codramo + "&CodSecc=" + codsecc + "&Periodo=" + periodo + ""
}

</script>
<font id="lblMensaje" color="FF6600"> El color Naranjo, indica que ya se ha contestado la encuesta 
para ese ramo.</font>
</body>
<%ObjetosLocalizacion("ramos.asp")%>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
