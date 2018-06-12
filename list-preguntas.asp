<%
    Response.Expires = -1
    Session.Timeout = 999
    Response.Buffer = false
    Server.ScriptTimeout = 99999
%>
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#include file="include/lib.asp" -->
<%
  dim Control 
  dim Valor
  dim strOpcionTexto
  dim rs
  dim rsG ,strsqlG,nroEnc,StrsglX

 ' StrsglX = "SELECT convert(varchar,numero) as numero FROM ra_encuesta WHERE TIPO_ENCUESTA = 1 AND ano = " & session("Ano_Ed") & " AND PERIODO = " & session("Periodo_pas") & "  "
 'if bcl_ado(StrsglX,rsX) then
 '       nroEnc=rsX("numero")
 '    else
 '       nroEnc="0"
 '    end if
 '  
 'Set rsBloques = getBloques(nroEnc)
 'Set rsPreguntas = getPreguntas(nroEnc, rsBloques("Orden"))
 
  strNumero=trim(request.Form("Minumero"))
  Accion=request.form("Accion")
  Micontrol="Combo" & strNumero
  'valor=trim(request.form("Micontrol"))
  strOpcion=trim(request.form("Micontrol"))
  strOpcionTexto= request.form("Mitexto")
  if trim(request.form("Mitexto")) = "" then 
    strOpcion=request.form("Micontrol")
  else  
     strOpcion= request.form("Mitexto")
  end if
  Dim numPreg
	'Total Pregunta Encuesta
     sql="Select count(*) from ra_preguntas"
     'set rst=conn.execute(sql)
     if bcl_ado(Sql,rst) then
        numPreg=valnulo(rst(0),num_)
     else
        numPreg=0
     end if

  if Accion= "Grabar" then

	 strOpcion=session("respuesta")
	 if strOpcion="" then
	    strOpcion=strOpcionTexto
	 end if 
      if strOpcion <> "" then
	     sql="Select * from ra_resultados where Ano='" & session("Ano_Ed") & "'"
		'JJJ sql=sql & " and Periodo='" & session("Periodo_Ed") & "' and CodCli='" & session("CodCli") & "'"
		 sql=sql & " and Periodo='" & session("Periodo_pas") & "' and CodCli='" & session("CodCli") & "'"
		 sql=sql & " and Rut='" & session("Rut") & "' and CodRamo='" & session("CodRamo") & "'"
		 sql=sql & " and CodSecc='" & session("CodSecc") & "' and CodProf='" & session("CodProf") & "'"
		 sql=sql & " and TipoDocente='" & session("TipoDoc") & "' and NroPregunta=" & strNumero & ""
		 'set rs = Conn.execute (sql)
		 
		 if bcl_ado(Sql,rs) then

               sql="Insert Into ra_Resultados (Ano,Periodo,CodCli,Rut,"
	           sql=sql & "CodRamo,CodSecc,CodProf,TipoDocente,NroPregunta,Respuesta) values("
	           'JJJ  sql=sql & "'" & session("Ano_Ed") & "','" & session("Periodo_Ed") & "' ,'" & session("CodCli") & "','" & session("Rut") & "',"
			   sql=sql & "'" & session("Ano_Ed") & "','" & session("Periodo_pas") & "' ,'" & session("CodCli") & "','" & session("Rut") & "',"
	           sql=sql & "'" & session("CodRamo") & "','" & session("CodSecc") & "' ,'" & session("CodProf") & "',"
 	           sql=sql & "'" & session("TipoDoc") & "'," & strNumero & " ," & valnulo(strOpcion,num_) & ")"         
	           'response.Write(sql)
	     else                 'Actualiza
		       sql="Update ra_resultados set Respuesta=" & valnulo(strOpcion,num_) & " where Ano='" & session("Ano_Ed") & "'"
		       'JJJ sql=sql & " and Periodo='" & session("Periodo_Ed") & "' and CodCli='" & session("CodCli") & "'"
			   sql=sql & " and Periodo='" & session("Periodo_pas") & "' and CodCli='" & session("CodCli") & "'"
		       sql=sql & " and Rut='" & session("Rut") & "' and CodRamo='" & session("CodRamo") & "'"
		       sql=sql & " and CodSecc='" & session("CodSecc") & "' and CodProf='" & session("CodProf") & "'"
		       sql=sql & " and TipoDocente='" & session("TipoDoc") & "' and NroPregunta=" & strNumero & ""		     
		 end if 		   
		 Session("Conn").execute Sql         
		 ' Si total de Respuesta es igual al numero de Preguntas , Actualiza a Profesor Encuestado
		 sql="Select Count(*) from ra_resultados where Ano='" & session("Ano_Ed") & "'"
		 'sql=sql & " and Periodo='" & session("Periodo_Ed") & "' and CodCli='" & session("CodCli") & "'"
		 sql=sql & " and Periodo='" & session("Periodo_pas") & "' and CodCli='" & session("CodCli") & "'"
		 sql=sql & " and Rut='" & session("Rut") & "' and CodRamo='" & session("CodRamo") & "'"
		 sql=sql & " and CodSecc='" & session("CodSecc") & "' and CodProf='" & session("CodProf") & "'"
		 sql=sql & " and TipoDocente='" & session("TipoDoc") & "'"
		'response.Write(sql)
	    'set rs = Conn.execute (sql)
         if bcl_ado(Sql,rs) then
            num_Respuestas=valnulo(rs(0),num_)	
			'response.Write(num_Respuestas & "<br>" & TotalRespuesta())					
			if num_Respuestas = TotalRespuesta() then
			    Strsql= " Update RA_ENCUESTADOS set Evaluado='SI'"
				Strsql=Strsql & " WHERE  ANO = " & valnulo(session("Ano_Ed"),NUM_) & ""
               'JJJ 31/01/2007
			  '  Strsql=Strsql & " AND  PERIODO = " & valnulo(session("Periodo_Ed"),NUM_) & " " 
			    Strsql=Strsql & " AND  PERIODO = " & valnulo(session("Periodo_pas"),NUM_) & " " 
                Strsql=Strsql & " AND  CODSECC = " & session("CodSecc") & "" 
                Strsql=Strsql & " AND  CODRAMO ='" & session("CodRamo") & "'"
                Strsql=Strsql & " AND  CODPROF ='" & session("CodProf") & "'"
                Strsql=Strsql & " AND CodCli='" & session("CodCli") & "'"
				'response.Write(Strsql)
				'response.End()
				
                'se hizo este camhio el 13 de marzo
		        Session("Conn").execute Strsql
			end if         
         end if
		 strPregunta=""

%>


<%
     end if				 
  end if
 
 function GetNombreAsig(Ramo)
 Dim StrSql
 Dim Rst

 'Strsql="select nombre,paterno,materno from mt_client where codcli='" & codcli & "'"
 sql="Select Nombre from ra_ramo where codramo='" & session("CodRamo") & "'"
 if bcl_ado(Sql,rst) then
   GetNombreAsig = Ucase(Rst("nombre"))
  else
    GetNombreAsig = "Sin Nombre"
  end if
  Rst.close() 
 end function
%>
<%
'**********************************************************************codigo Nuevo************************************

function BuscaRespuestaTextoLibre(numero)
Dim rstResp, sql
   sql="Select textolibre from RA_RESULTADOS"
   sql=sql & " where ano=" & session("Ano_Ed") & " and Periodo=" & session("Periodo_pas") & ""
   sql=sql & " and Codcli='" & session("CodCli") & "' and CodRamo='" & session("CodRamo") & "'"
   sql=sql & " and CodSecc= " & session("CodSecc") & " and CodProf ='" & session("CodProf") & "'"
   sql=sql & " and NroPregunta= " & numero & "" 
   sql=sql & " and encuesta = '" & session("NumeroEncuestaVigente") & "'"
   
   if bcl_ado(Sql,rstResp) then
      BuscaRespuestaTextoLibre= ValNulo(rstResp("textolibre"),Str_)
   else
      BuscaRespuestaTextoLibre= ""
   end if	  
end function  

function BuscaRespuesta(numero)
Dim rstResp, sql
   sql="Select Respuesta from RA_RESULTADOS"
   'JJJ sql=sql & " where ano=" & session("Ano_Ed") & " and Periodo=" & session("Periodo_Ed") & ""
   sql=sql & " where ano=" & session("Ano_Ed") & " and Periodo=" & session("Periodo_pas") & ""
   sql=sql & " and Codcli='" & session("CodCli") & "' and CodRamo='" & session("CodRamo") & "'"
   sql=sql & " and CodSecc= " & session("CodSecc") & " and CodProf ='" & session("CodProf") & "'"
   sql=sql & " and NroPregunta= " & numero & "" 
   sql=sql & " and encuesta = '" & session("NumeroEncuestaVigente") & "'"
   
   if bcl_ado(Sql,rstResp) then
      BuscaRespuesta= ValNulo(rstResp("Respuesta"),Str_)
   else
      BuscaRespuesta= ""
   end if	  
end function   

if Accion = "Grabar" then
     StrSql= "UPDATE RA_Encuestados SET OBSERVACION='" & EliminaPalabras(request("txtObs")) & "'"
	 Strsql=Strsql & " WHERE  ANO = " & valnulo(session("Ano_Ed"),NUM_) & ""
     'JJJ Strsql=Strsql & " AND  PERIODO = " & valnulo(session("Periodo_Ed"),NUM_) & " " 
	 Strsql=Strsql & " AND  PERIODO = " & valnulo(session("Periodo_pas"),NUM_) & " " 
     Strsql=Strsql & " AND  CODSECC = " & session("CodSecc") & "" 
     Strsql=Strsql & " AND  CODRAMO ='" & session("CodRamo") & "'"
     Strsql=Strsql & " AND  CODPROF ='" & session("CodProf") & "'"	
     Strsql=Strsql & " and CodCli='" & session("CodCli") & "'"
     'se hizo este camhio el 13 de marzo
	 Session("Conn").execute Strsql
	 Accion=""
end if

 ' Consulta Respuesta Existente
     StrSql= "SELECT Observacion FROM RA_Encuestados"
	 Strsql=Strsql & " WHERE  ANO = " & valnulo(session("Ano_Ed"),NUM_) & ""
     'JJJ Strsql=Strsql & " AND  PERIODO = " & valnulo(session("Periodo_Ed"),NUM_) & " " 
	 Strsql=Strsql & " AND  PERIODO = " & valnulo(session("Periodo_pas"),NUM_) & " " 
	 
     Strsql=Strsql & " AND  CODSECC = " & session("CodSecc") & "" 
     Strsql=Strsql & " AND  CODRAMO ='" & session("CodRamo") & "' "
	 Strsql=Strsql & " AND  CODPROF='" & session("CodProf") & "'"
	 Strsql=Strsql & " AND  CODCLI='" & session("CodCli") & "'"
	 Strsql=Strsql & " and encuesta=(select top 1 COALESCE(NUMERO,0) from RA_ENCUESTA WHERE convert(date,GETDATE()) BETWEEN fecini AND fecfin AND TIPO_ENCUESTA=1 ) " 
	 'se hizo este camhio el 13 de marzo
     'set rs=conn.execute(StrSql)
	'response.Write(Strsql)
	 if bcl_ado(Strsql,rs) then
	    strObs=ValNulo(rs("Observacion"),Str_)
	 else
	    strObs=""
	 end if
	 
	 'strsqlG = "SELECT CODTIPOPREGUNTA, DESCRIPCION FROM RA_ENCUESTA_DETALLE,RA_TIPOPREGUNTA WHERE NRO_ENCUESTA = '1'"
     'strsqlG = strsqlG & " AND CODTIPOPREGUNTA = TIPO_PREGUNTA GROUP BY  CODTIPOPREGUNTA, DESCRIPCION"
	 
		strsqlG = "SELECT tp.CODTIPOPREGUNTA, tp.DESCRIPCION FROM RA_TIPO_ENCUESTA te,RA_ENCUESTA e,RA_ENCUESTA_DETALLE ed,RA_TIPOPREGUNTA tp"
		strsqlG = strsqlG & " WHERE te.codigo = 1 AND e.TIPO_ENCUESTA = te.Codigo"
		strsqlG = strsqlG & " and e.numero='" & session("NumeroEncuestaVigente") & "' AND e.numero = ed.NRO_ENCUESTA AND ed.TIPO_PREGUNTA = tp.CODTIPOPREGUNTA"
		strsqlG = strsqlG & " GROUP BY tp.CODTIPOPREGUNTA, tp.DESCRIPCION"
		 
	 
	'  if not bcl_ado(strsqlG,rsTsT) then
	'    response.Redirect("ed.htm")		
	'  end if
	 
	 
	' 	StrSql= "SELECT a.numero,a.pregunta,b.codtipopregunta,b.descripcion "
    ' strsql=strsql & " FROM RA_PREGUNTAS a , ra_tipopregunta b  "
    ' strsql=strsql & " where a.COD_TIPOPREGUNTA=CODTIPOPREGUNTA  "
    ' strsql=strsql & " ORDER BY b.codtipopregunta,a.numero  "
	 'response.Write(strsql)
	 'response.end()
     'set rs=conn.execute(StrSql)
%>
<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="ed.css" type="text/css">
<script language="JavaScript">
<!--
  function Salir()
{ 
//window.location.href="salir.asp"
window.top.location.href="menu_encuestas.asp"
}
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}
function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}


function validacombo(control){
	
	for (i=1; i <= control; i++)
     {
		 
		//indice = document.getElementById("Combo"+i).selectedvalue;
		//indice = document.forms["TheForm"].getElementById("Combo"+i).selectedindex;
		//indice = TheForm.document.getElementById("Combo"+i).options[TheForm.document.getElementById("Combo"+i).selectedIndex].value;
		var combo = document.getElementById('Combo'+i);
		indice =combo.options[combo.options.selectedIndex].value;
		
		if( indice == null || indice == "NO" ) {
			alert('- Debe contestar toda la encuesta')
		    return false;
		}
	 }
	return true;
}



//-->
</script>
<link href="css/tablas.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" text="#000000" topmargin="0" onLoad="MM_preloadImages('ima/bot/limpiar-on.gif','ima/bot/enviar-on.gif')">
<form id="TheForm" name="TheForm" method="POST" action="Grabar.asp">
  <p> 
    <input type="hidden" name="Accion" value="<%=Accion%>">
    <input type="hidden" name="Minumero" value="<%=Minumero%>">
  </p>
  <table width="810" border="0" class="nom-profe">
    <tr> 
    
    
    <%
		strParame="SELECT coalesce(dbo.Fn_ValorParame('OCULTAGLOSAENCUESTAALUMNOS'),'NO')Parame"
		set rstParame= Session("Conn").Execute(strParame)		
		if not rstParame.eof then
			OCULTAGLOSAENCUESTAALUMNOS=rstParame("Parame")
		else
			OCULTAGLOSAENCUESTAALUMNOS="NO" 
		end if
		
		strParame="SELECT coalesce(dbo.Fn_ValorParame('GLOSAEDEXTRALINEA1'),'')Parame"
		set rstParame= Session("Conn").Execute(strParame)		
		if not rstParame.eof then
			GLOSAEDEXTRALINEA1=rstParame("Parame")
		else
			GLOSAEDEXTRALINEA1="" 
		end if
	%>
    <%if OCULTAGLOSAENCUESTAALUMNOS <> "SI" then%>
      <td class="tex" align="left"><p align="justify"><span class="Tit-celdas"><font id="lblTitCuestion" size="2" face="Arial, Helvetica, sans-serif">CUESTIONARIO 
          DE OPINION DEL ALUMNADO SOBRE LA DOCENCIA</font></span><font color="4a5da1" size="2" face="Arial, Helvetica, sans-serif"><br>
          </font><span id="lblTexto1" class="datos-informe">El presente cuestionario pretende conocer la opini&oacute;n de 
          los alumnos de esta asignatura acerca de la calidad del proceso de la 
          ense&ntilde;anza por parte del docente.</span> <br><span class="datos-informe">
          Es an&oacute;nimo, por favor responde con objetividad, sinceridad y 
          responsabilidad. Los resultados servir&aacute;n para mejorar la calidad 
          docente.</span><br><span class="datos-informe">
          POR FAVOR NO CONTESTES SI NO ASISTES HABITUALMENTE A CLASE O NO DISPONES 
          DE SUFICIENTE INFORMACION PARA OPINAR.<br>
          INSTRUCCIONES:  Usando el cuadro de respuestas al lado de cada pregunta 
           selecciona la opción que corresponde a la respuesta que más te identifica. 
           </span>
      </td>
    <%end if%>    
    </tr>
    
    <%
	MensajeGlosa=""
	strGlosa="sp_traeFwGlosasPortales 'list-preguntas.asp','PORTAL ALUMNO'"
	Set rsGlosa = Session("Conn").execute(strGlosa)
			
	if not rsGlosa.eof then       
		while not rsGlosa.eof
			MensajeGlosa =  MensajeGlosa + rsGlosa("glosa") + "</br>"
		rsGlosa.movenext		
		wend
	end if
  %>  		
  <%if MensajeGlosa<> ""then%>
	 <tr>
    	<td class="tex" align="left">
    	  <%=MensajeGlosa%>
      </td>
    </tr>
  <%else%>  
  	 <tr>
    	<td class="tex" align="left"><p align="justify"><span class="datos-informe">
    	  <%=GLOSAEDEXTRALINEA1%>
      </td>
    </tr>
  <%end if%>
     
  </table>
   <%

   if bcl_ado(strsqlG,rsG) then  
Control=1
	  while not rsG.EOF	  
	  
	    StrSql= "SELECT a.numero,ed.nro_linea,e.NUMERO AS encuesta,a.CODRESPUESTA,a.pregunta,CONVERT(VARCHAR,b.codtipopregunta) as codtipopregunta,b.descripcion,coalesce(TR.TextoLibre,'NO')TextoLibre "
     strsql=strsql & " FROM RA_PREGUNTAS a , ra_tipopregunta b ,RA_ENCUESTA E ,RA_ENCUESTA_DETALLE ED,RA_TIPORESPUESTA TR "
     strsql=strsql & " WHERE E.NUMERO = ED.NRO_ENCUESTA "
     strsql=strsql & " AND ED.TIPO_PREGUNTA = B.CODTIPOPREGUNTA "
     strsql=strsql & " AND ED.NRO_PREGUNTA = A.NUMERO "
	 strsql=strsql & " and TR.CODTIPORESPUESTA = A.CODRESPUESTA "
     strsql=strsql & " AND E.TIPO_ENCUESTA = 1 "
	 strsql=strsql & " and e.numero='" & session("NumeroEncuestaVigente") & "'  "
	 strsql=strsql & " AND b.codtipopregunta =  " & rsG("CODTIPOPREGUNTA") & ""
	 'strsql=strsql & " AND e.NUMERO = (select COALESCE(numero,0) from  dbo.RA_ENCUESTA WHERE GETDATE() BETWEEN fecini AND FECFIN AND TIPO_ENCUESTA=1) "
     strsql=strsql & " ORDER BY b.codtipopregunta,a.numero "
	 'VALIDAR QUE ENCUESTA DOCENTE SEA PARA PERIODO Y AÑO ACTUAL!!!
	  'response.Write(strsql)
	  'response.End()
	  
	   %>
	  
  <table width="825" border="0" cellspacing="1" cellpadding="0" bordercolor="#000000">
    <tr bgcolor="4a5da1" background="Imagenes/fdo-cabecera-cel.jpg">      
      <td width="2%" colspan="3" height="20" background="Imagenes/fdo-cabecera-cel.jpg" class="tex" align="left"> <div align="left"><span class="text-cabecera-celda"><font color="#FFFFFF"><%=rsG("descripcion")%></font></span></div>
    </tr>
    
     <%	 
	
if bcl_ado(Strsql,rs) then  


       ColorFila="#DBECF2" '"#E9F0E9"          
	   valordefecto=""

	  do while not rs.eof	     
		 if valordefecto=rs("codtipopregunta") then
		    registro1=""
			registro2=""
		 else
		    'registro1=rs("codtipopregunta")
			registro2= rs("descripcion")
		    valordefecto=rs("codtipopregunta")
		 end if	 
		 
		 
		  %>
    <tr bgcolor="<%=ColorFila%>"> 
        <td width="2%" height="10"><div align="center">&nbsp;<font face="Arial, Helvetica, sans-serif" size="2"><b><font size="1"><%=rs("nro_linea")%>)
          </font></b></font></div></td>
      <td width="60%" height="10">&nbsp;<font face="Arial, Helvetica, sans-serif" size="1"><b><%=rs("pregunta")%> 
        </b></font></td>
      <td height="38%"> <div align="center"></div>
        <%
Micontrol="Combo" & int(control)
Mitexto="Texto" & int(control)
Mivariable="P"  & int(control)
MiTipo = "Enc" & int(control)
Session("NroEncuesta") = rs("encuesta")
%>

<input type="hidden" name="<%=Mivariable%>" value="<%=rs("numero")%>">
<input type="hidden" name="<%=MiTipo%>" value="<%=rs("encuesta")%>">
    
<%if rs("TextoLibre") = "SI" then
	valorTextoLibre = BuscaRespuestaTextoLibre(rs("numero"))
%>
	<textarea maxlength="200" name="<%=Micontrol%>" id="<%=Micontrol%>" style="width:310px;font-size:12px; font-type:Arial"  rows="3" wrap="virtual"><%=valorTextoLibre%></textarea>
<%else%>
    <select name="<%=Micontrol%>" id="<%=Micontrol%>" style="width:100%">
    <option value="NO">===[ SELECCIONE UNA OPCION ]===</option>
	<%dim str 
      dim rst 
      dim Opcion
      str="select codrespuesta,respuesta from RA_RESPUESTA where codtiporespuesta = " & rs("CODRESPUESTA") & " "
      'set rst=Server.CreateObject ("ADODB.Recordset")
      'set rst=Conn.Execute (str)
        'if not rst.EOF then
         if bcl_ado(Str,rst) then
            Micontrol=""
            pricarr=true	
            while not rst.EOF
        '	if Micontrol="" and pricarr then
                    'sel=" selected"
        '			pricarr=false
        '			strvalor=rst("respuesta") 
        '		else
                    if trim(BuscaRespuesta(rs("numero")))=trim(rst("codrespuesta")) then
                        sel=" selected"
                    else					
                        sel=" "
                    end if
                    
        '		end if
                Micontrol= Micontrol & "<option value=""" & rst("codrespuesta") & """ " & sel & " >"
                Micontrol=	Micontrol & rst("respuesta") & "</option>" & vbcrlf
                rst.MoveNext 
            wend
        end if%>
              <%=Micontrol%> 
            </select> 
<%end if%>
        <input type="text" name="<%=Mitexto%>" style="display:none;"  maxlength="1" size="1" disabled="false" onChange="Javascript:Rango(this)" value=<%=BuscaRespuesta(rs("numero"))%> > 
        </div> </td>
    </tr>
    <% rs.movenext
     control=int(control) + 1
     if ColorFila="#FFF8F5" then
	    ColorFila= "#DBECF2" '"#E9F0E9"
	 else
	   	ColorFila="#DBECF2" '"#FFF8F5" 
     end if		
	   loop
	   else %>
    <tr bgcolor="#FFF8F5"> 
      <td width="3%" height="10">&nbsp;</td>
      <td width="23%" height="10">&nbsp;</td>
      <td width="2%" height="10">&nbsp;</td>
      <td width="48%" height="10">&nbsp;</td>
      <td height="10">&nbsp;</td>
    </tr>
    <%end if%>
    <!--  <tr bgcolor="#E9F0E9"> 
    <td width="10%" height="10">&nbsp;</td>
    <td width="75%" height="10">&nbsp;</td>
    <td height="10" width="15%">&nbsp;</td>
  </tr>
  <tr bgcolor="#FFF8F5"> 
    <td width="10%" bgcolor="#FFF8F5" height="10">&nbsp;</td>
    <td width="75%" height="10">&nbsp;</td>
    <td width="15%" height="10">&nbsp;</td>
  </tr>-->
  </table>
  </br>
   <%
   
    rsG.movenext
    wend 
	end if

	 %>
  <p><span class="Tit-celdas">Escriba una observaci&oacute;n en general 
    o en relaci&oacute;n a una respuesta en particular(m&aacute;ximo 500 caracteres)</span><span class="resaltar-pregunta"> : </span></p>
  <table width="81%" border="0" cellspacing="0" cellpadding="0" height="47">
    <tr> 
    <td rowspan="2" width="68%" valign="bottom"> 
      <textarea maxlength="500" name="txtObs" cols="80" rows="4"><%=strObs%></textarea>
    </td>
    <!--<td valign="bottom" height="20" width="32%"> 
      <div align="left" ><a href="javascript:Limpiar();" onMouseOver="MM_swapImage('Image1','','Imagenes/Botones/limpiar-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="Imagenes/BOtones/limpiar-of.gif" width="178" height="45" name="Image1" border="0"></a></div>
    </td> -->
  </tr>
  <tr> 
    <td valign="bottom" height="20" width="32%"> 
      <div align="left"></div>
    </td>
  </tr>
</table>
  <table width="613" border="0">
    <tr> 
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr> 
      <td width="104">  
      
      <a href="javascript:Mienvio()" onClick="return validacombo(<% response.Write(control-1) %>)" onMouseOver="MM_swapImage('Image2','','Imagenes/botones/enviar-on.gif',1)" onMouseOut="MM_swapImgRestore()"><img src="Imagenes/BOtones/enviar-of.gif" width="178" height="45" name="Image2" border="0"></a></td>
      <td width="499" class="titulo"><font color="#0000CC" size="2" face="Arial, Helvetica, sans-serif" class="text-normal-celdas">Haz 
        click en el bot&ograve;n <em><strong><font color="#CC0000">ENVIAR</font></strong></em> 
        para grabar todas tus respuestas para este docente.</font></td>
    </tr>
  </table>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
</form>
<p>&nbsp;</p>
<p>
  <script>

function PasarValores(num,preg)
{
parent.bottomFrame2.location.href="pregunta-opc.asp?NumPre=" + num + "&Pregunta=" + preg + ""
}

function Grabar()
{
//   document.TheForm.Accion.value="Grabar";
   document.TheForm.submit();
   }
 
 function Limpiar()
{
   document.TheForm.txtObs.value="";   
   }  
function ActualizarValor(numero) {
  {alert("Pregunta Nº " + numero)
  }  
  document.TheForm.Minumero.value=numero;
  document.TheForm.Accion.value="Actualizar";
  document.TheForm.submit();
}
function Actualizar() {
  document.TheForm.Accion.value="Actualizar";
  document.TheForm.submit();
}
</script>
  <script>
function Grabar2(valor) {
   if (valor=="") {
      alert("Debe Seleccionar una Pregunta de la Lista");
      return;
	 } 	  
	  
   if (document.TheForm.txtOpcion.value!=""){ 
			    if ((isNaN(document.TheForm.txtOpcion.value))==true)
		           {	alert("Valor debe ser Númerico(Entero)");
			            document.TheForm.txtOpcion.select();
			            document.TheForm.txtOpcion.focus();
			            return;
		        }else {
				if (document.TheForm.txtOpcion.value < 1 || document.TheForm.txtOpcion.value > 7) {
				    alert ("Valor fuera de rango");
					return;
				   }
				} 
	    }	
  document.TheForm.Accion.value="Grabar2";
  document.TheForm.submit();
}
function Intruc()
{ 
parent.parent.mainFrame.location.href="instrucciones.htm"
}

function Traspasa(Obj)
{
   i = Obj.name.substring(5,7);
   //Obj.name;
   //alert(i);
   for (j=0; j < document.TheForm.length; j++)
   { 
      //alert (document.TheForm.elements[j].name );
      if (document.TheForm.elements[j].name == "Texto" + i)
	  {
	     if (Obj.value == "OTRO")
		 {
	       document.TheForm.elements[j].value = "";
	       document.TheForm.elements[j].disabled= false;
		 }else
		 {
	       document.TheForm.elements[j].value = Obj.value;
	       document.TheForm.elements[j].disabled = true;
		 }
	  }
   }   
}

</script>
  <script>
function valueGetter(Control){   
//var msgWindow=window.open("");   
for (var i = 0; i < newWindow.document.TheForm.elements.length; i++);
 {      
 //msgWindow.document.write(newWindow.document.TheForm.elements[i].name + "<BR>");
 if (document.TheForm.elements[i].name==Control){
//     document.TheForm.Vcombo.value=document.TheForm.elements[i].value;
	 alert(document.TheForm.Vcombo.value);
     return;
 }
 }
 }
</script>
  <script>
function Rango(Obj)
 {
   i = Obj.name.substring(5,7);
   //Obj.name;
   //alert(i);
   for (j=0; j < document.TheForm.length; j++)
   { 
      //alert (document.TheForm.elements[j].name );
      if (document.TheForm.elements[j].name == "Texto" + i){
 //rutina
 if (document.TheForm.elements[j].value!=""){ 
			    if ((isNaN(document.TheForm.elements[j].value))==true)
		           {	alert("Valor debe ser Númerico(Entero)");
			            document.TheForm.elements[j].select();
			            document.TheForm.elements[j].focus();
			            return;
		        }else {
				if (document.TheForm.elements[j].value < 1 || document.TheForm.elements[j].value > 7) {
				    alert ("Valor fuera de rango");
					return;
				   }
				} 
	    }	
 //
	  }
   }   
}

</script>
  <script>
  
  function Mienvio(){
	
   document.TheForm.action="Grabar.asp";
 for (i=0;i<document.TheForm.length; i++)
 {
    document.TheForm.elements[i].disabled=false;
 }
//Alert ("test");
// document.TheForm.Action="Graba";
 document.TheForm.submit();
}

</script>
  <script>
function Asignar(Preg,Res)
{
   Preg ="Texto" + Preg;
     for (j=0; j < document.TheForm.length; j++){ 
      //alert (document.TheForm.elements[j].name );
      if (Preg==document.TheForm.elements[j].name){
	     document.TheForm.elements[j].value==Res;
	  }
   }   
}
</script>
</p>
</body>
<%ObjetosLocalizacion("list-preguntas.asp")%>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->