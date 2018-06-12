<%
    Response.Expires = -1
    Session.Timeout = 999
    Response.Buffer = false
    Server.ScriptTimeout = 99999
%>

<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<%
   session("CodSecc")= Valnulo(request("CodSecc"),num_)
   session("CodRamo")=request("CodRamo")
   session("Periodo_pas")=request("Periodo")
   
  	strParame="SELECT coalesce(dbo.Fn_ValorParame('ENCUESTAPATODOSLOSDOCENTES'),'')Parame"
	set rstParame= Session("Conn").Execute(strParame)		
	if not rstParame.eof then
		ENCUESTAPATODOSLOSDOCENTES = rstParame("Parame")
	end if
		
    'Strsql="SELECT DISTINCT HOR.CODPROF, PRO.DV, PRO.AP_PATER, PRO.AP_MATER, PRO.NOMBRES,PRO.TIPOPROFE" 
    'Strsql=Strsql & " FROM RA_SECCIO SEC, RA_PROFES PRO, RA_HORPROF HOR " 
    'Strsql=Strsql & " WHERE SEC.ANO = " & valnulo(session("Ano_Ed"),NUM_) & ""
	'Strsql=Strsql & " AND  SEC.CODSECC = " & session("CodSecc") & "" 
    'Strsql=Strsql & " AND  SEC.CODRAMO ='" & session("CodRamo") & "' "
	'Strsql=Strsql & " AND  SEC.CODSEDE ='" & session("CodSede") & "' "
    'Strsql=Strsql & " AND  SEC.CERRADA = '" & session("Cerrada") & "'"
	'Strsql=Strsql & " AND  SEC.CODRAMO=HOR.CODRAMO" 
	'Strsql=Strsql & " AND  SEC.CODSECC=HOR.CODSECC" 
	'Strsql=Strsql & " AND  SEC.ANO=HOR.ANO" 
	'Strsql=Strsql & " AND  SEC.PERIODO=HOR.PERIODO" 
	'Strsql=Strsql & " AND  SEC.CODSEDE=HOR.CODSEDE" 
	'Strsql=Strsql & " AND  SEC.CODCARR=HOR.CODCARR" 
	'Strsql=Strsql & " AND  PRO.CODPROF=HOR.CODPROF" 
	'if ENCUESTAPATODOSLOSDOCENTES <> "SI" then
	'	Strsql=Strsql & " AND  SEC.CODPROF = PRO.CODPROF" 
	'end if 
	
	if ENCUESTAPATODOSLOSDOCENTES <> "SI" then
		filtradocente = 1
	else 
		filtradocente = 0
	end if 
	
	Strsql="SP_LISTAPROFES_ENCUESTA '"& session("codcli") &"','" & session("CodRamo") & "',"& valnulo(session("Ano_Ed"),NUM_)  &",null," & session("CodSecc") & ",'" & session("Cerrada") & "',"& filtradocente &" "
	
	'Si este curso está planificado
	'response.Write(Strsql)
	
    if BCL_ADO(Strsql,rs) then
	'set rs=Session("Conn").execute(StrSql)
		'response.Write(profe)		
		function ValidaRamos(Profe)     
			ValidaRamos=True
			sql="Select * from ra_encuestados where Ano=" & session("Ano_Ed") & ""
			if session("NumeroMaximoPrueba") <> 1 then
				sql=sql & " and Periodo='" & session("Periodo_pas") & "' "
			end if
			sql=sql & " and CodCli='" & session("CodCli") & "' and CodProf='" & Profe & "'"
			sql=sql & " and CodRamo='" &  session("CodRamo") & "' and CodSecc=" & session("CodSecc") & " and Evaluado='NO'"  
			
			'set rst=Session("Conn").execute(sql)		 
			'if not rst.eof then 
			if BCL_ADO(sql,rst) then
				ValidaRamos=False
			end if    	  	  	  
		end function
		
		
		
		function ValidaEncuestaProfesor(Profe)     
			ValidaEncuestaProfesor=False
			
			sql="Select * from ra_encuestados where Ano=" & session("Ano_Ed") & ""
			if session("NumeroMaximoPrueba") <> 1 then
				sql=sql & " and Periodo='" & session("Periodo_pas") & "' "
			end if
			sql=sql & " and CodCli='" & session("CodCli") & "' and CodProf='" & Profe & "'"
			sql=sql & " and CodRamo='" &  session("CodRamo") & "' and CodSecc=" & session("CodSecc") & " and Evaluado='SI'"
			sql=sql & " and encuesta=(select top 1 COALESCE(NUMERO,0) from RA_ENCUESTA WHERE convert(date,GETDATE()) BETWEEN fecini AND fecfin AND TIPO_ENCUESTA=1 ) "   

			'set rst=Session("Conn").execute(sql)		 
			'if not rst.eof then 
			if BCL_ADO(sql,rst) then
				ValidaEncuestaProfesor=True
			end if    	  	  	  
		end function
		
		function ExisteEncuesta()     
			
			'strsqlG = "SELECT tp.CODTIPOPREGUNTA, tp.DESCRIPCION FROM RA_TIPO_ENCUESTA te,RA_ENCUESTA e,RA_ENCUESTA_DETALLE ed,RA_TIPOPREGUNTA tp"
			'strsqlG = strsqlG & " WHERE te.codigo = 1 AND e.TIPO_ENCUESTA = te.Codigo"
			'strsqlG = strsqlG & " AND  e.Ano = " & session("Ano_Ed") & " AND e.PERIODO = " & session("Periodo_Ed") & " AND e.numero = ed.NRO_ENCUESTA AND ed.TIPO_PREGUNTA = tp.CODTIPOPREGUNTA"
			'strsqlG = strsqlG & " GROUP BY tp.CODTIPOPREGUNTA, tp.DESCRIPCION"
			
			sqlenc="select numero from RA_ENCUESTA Where CONVERT(DATE,getDate())>=FecIni And CONVERT(DATE,getDate())<=FecFin AND TIPO_ENCUESTA=1"	
			
			if bcl_ado(sqlenc,rstenc) then		
				session("NumeroEncuestaVigente")=valnulo(rstenc("numero"),num_)
				ExisteEncuesta=True		
			else
				ExisteEncuesta = False 
			end if
				  	  	  
		end function
		
	end if
  %>
<html>
<head>
<title>Evaluaci&oacute;n Docente</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="ed.css" type="text/css">
<link href="css/tablas.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#FFFFFF" text="#000000" topmargin="0"> 
<form name="TheForm" method="post" action="">
  <span id="lblTituloDocente" style="font-size: 25px"  class="text-menu">Profesores de la Asignatura</span> 
  <table width="75%" border="0" cellspacing="1" cellpadding="0" bordercolor="#000000">
  <tr bgcolor="#000000" class="text-cabecera-celda"> 
    
    <td width="70%" height="20" background="Imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda"> 
      <div align="center"><font color="#FFFFFF">Nombre</font></div>
    </td>
    <td width="5%" height="20" background="Imagenes/fdo-cabecera-cel22.jpg" class="text-cabecera-celda"> 
      <div align="center"><font color="#FFFFFF">Tipo</font></div>
    </td>
  </tr>
  <%
	
	Strsql="SP_LISTAPROFES_ENCUESTA '"& session("codcli") &"','" & session("CodRamo") & "',"& valnulo(session("Ano_Ed"),NUM_)  &"," & valnulo(session("Periodo_pas"),NUM_) & "," & session("CodSecc") & ",'" & session("Cerrada") & "',"& filtradocente &" "
	
   %>
   <%if BCL_ADO(Strsql,rs) then  
'	  	response.Write(strsql)
'		response.End
		ColorFila="#E7F0ED"          
	  	do while not rs.eof	  
	     	if ValidaEncuestaProfesor(rs("CODPROF")) then   
		    	ColorFila="#FF6600"
			end if
			if ExisteEncuesta() then
			Existe = True
			end if
	      	strNombreProfesor=EliminaPalabras(ValNulo(rs("ap_pater"), STR_) & " " & ValNulo(rs("ap_Mater"), STR_) & " " & ValNulo(rs("Nombres"), STR_))
		
			%>
			<tr bgcolor="<%=ColorFila%>"> 
			<td width="10%" height="20">&nbsp;<font face="Arial, Helvetica, sans-serif" size="1"><a href="javascript:PasarValores('<%=Encode(rs("CODPROF"))%>','<%=rs("TIPOPROFE")%>','<%=strNombreProfesor%>');"><%=strNombreProfesor%></a></font></td>
			<td width="25%" height="20">&nbsp;<font size="1" face="Arial, Helvetica, sans-serif"><%=rs("TIPOPROFE")%></font></td>
			</tr>
	  <% rs.movenext
		 if ColorFila="#E7F0ED" then
			ColorFila= "#EEF0FF"
		 else
			ColorFila="#E7F0ED" 
		 end if		
	 loop
  else %>
  <tr bgcolor="#EEF0FF">   
      <td width="10%" height="20" bgcolor="#EEF0FF">&nbsp;</td>
      <td width="5%" height="20" bgcolor="#EEF0FF">&nbsp;</td>
      <td width="60%" height="20" bgcolor="#EEF0FF">&nbsp;</td>
    <td width="25%" height="20">&nbsp;</td>
  </tr>
  <%end if%> 
  <!--<tr bgcolor="#E7F0ED"> 
    <td width="10%" height="10">&nbsp;</td>
    <td width="5%" height="10">&nbsp;</td>
    <td width="60%" height="10">&nbsp;</td>
    <td width="25%" height="10">&nbsp;</td>
  </tr>
  <tr bgcolor="#EEF0FF"> 
    <td width="10%" height="10">&nbsp;</td>
    <td width="5%" height="10">&nbsp;</td>
    <td width="60%" height="10">&nbsp;</td>
    <td width="25%" height="10">&nbsp;</td>
  </tr>-->
</table>
<script>
function PasarValores(cod,tipo,profe) {
   // parent.parent.bottomFrame1.location.href="list-preguntas.asp?CodProfe=" + cod + "&Tipo=" + tipo + "&Nom=" + profe + ""; 
     window.parent.parent.parent.mainFrame.location.href="ed2.asp?CodProfe=" + cod + "&Tipo=" + tipo + "&Nom=" + profe + ""; 
}

</script>
</form>
</body>
<%ObjetosLocalizacion("profesores.asp")%>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->