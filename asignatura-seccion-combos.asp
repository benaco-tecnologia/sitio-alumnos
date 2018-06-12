<%
  Response.Expires = -1 
  Response.Buffer = false
%>
<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="include/cupos.inc" -->
<!--#INCLUDE FILE="include/nombreramo.inc" -->

<script language="JavaScript">
			<!--
			function MM_preloadImages() { //v3.0
			  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
				var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
				if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
			}
			
			function MM_swapImgRestore() { //v3.0
			  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
			}
			
			function MM_findObj(n, d) { //v4.0
			  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
				d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
			  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
			  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
			  if(!x && document.getElementById) x=document.getElementById(n); return x;
			}
			
			function MM_swapImage() { //v3.0
			  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
			   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
			}
			//-->
			
			function marca(Obj, valor)
			{	
			   if (Obj.checked) {Obj.value = "false"} 
			   else {Obj.value = valor}
			  //Obj.checked = !Obj.checked;
			  Obj.blur()
			}
			
			function RevisaOpcion(Obj, Ramo)
			{ var j, k;
			  //alert(Obj.value);
			  for (j=0; j< document.forms.length; j++)
			  {
				 //if (document.forms.elements[i].type == "option")
				 for (k=0; k < document.forms[j].elements.length ;k++)
				 {
					if (document.forms[j].name != Obj.form.name)
					{
					  //alert(document.forms[j].name); 
					  if (document.forms[j].elements[k].type == "radio")
					  { 
						//alert(Ramo + " : " + document.forms[j].elements[k].value); //SEGUIR 
						//alert(Ramo + " : " + document.forms[j].elements[k].id); //SEGUIR  
						if ((Ramo == document.forms[j].elements[k].id) && (document.forms[j].elements[k].checked) )
						{ alert("Ramo : " + document.forms[j].elements[k].id + " inscrito con " + document.forms[j].name ); 
			//            { alert("Ramo : " + document.forms[j].elements[k].value + " inscrito con " + Obj.form.name ); 
						  Obj.checked = false;
						  return false;}
					  }
					}         
				 }
			  }
			  return true;
			}

			function GrabaSel(Obj, Curso, Ramo, Seccion, i, TipoCurso, CodSeccTeo)
			{ var j, k, Sec;
			 		 
			 //alert(Obj.name);	
			 //alert(Obj.checked); false
			 //lert(Obj.form.name); arq2009
			 
		 	
			  if (RevisaOpcion(Obj, Ramo))
			  {
				if ((Obj.checked) && (Obj.value == "false") )
				{
				  Obj.checked = false;
				  //RamoInscrito[i] = false;
				 
				 parent.Horario.location = "Inscribe.asp?I=S&C=" + Curso + "&R=" + Ramo + "&S=0&H=" + TipoCurso + "&D=" + CodSeccTeo;
				   
				}
				else {
									
			     parent.Horario.location = "Inscribe.asp?I=S&C=" + Curso + "&R=" + Ramo + "&S=" + Seccion + "&H=" + TipoCurso + "&D=" + CodSeccTeo;
				}

			  } 
			 
			 //si deseleccionar un teorico desmarca sus subsecciones 
			 var RamoP = Ramo +'P';
			 var RamoL = Ramo +'L';
			 
			 if (Obj.value == "false")
			 {	
			 	if (Obj.name != RamoL && Obj.name != RamoP)
				{
					opP = document.getElementsByName(RamoP);
					if (opP[0] != null)  // if nuevo para cuando no tiene este tipo de subseccion
						opP[0].checked = false;
					
					opL = document.getElementsByName(RamoL);
					if (opL[0] != null) 
						opL[0].checked = false;
				}
			 } 
			}
			
			function visualiza() {
			  var x=window.open ("tex-seccionhorario.htm","ResumenInscripción","width=300,height=300");
			 }
			
			
			function EscondeHor()
			{
			   //alert("paso");
			   //parent.location = "asignatura-seccion.asp?Ramos="
			   parent.Seleccion.resizeBy(1000, 1000);
			}
			
			</script>
			
			<script language="JavaScript">
			<!--
			function MM_openBrWindow(theURL,winName,features) { //v2.0
			  window.open(theURL,winName,features);
			}
			//-->
</script>

<% 

dim strNumerico
dim Nombrecarrera

RutCli=session("RutCli")

set rs=Server.CreateObject ("ADODB.Recordset")
	
	sql="select b.codcarr,b.nombre_l,estacad "
	sql=sql & " from mt_alumno a, mt_carrer b "
	sql=sql & " where a.codcarpr=b.codcarr"
	sql=sql & " and a.rut='" & Rutcli & "'"

set rs=Session("Conn").Execute (sql)
if not rs.EOF then
	cmbNumeros=""
	pricarr=true	
	while not rs.EOF
		if strcarrera ="" and pricarr then
			'sel=" selected"
			pricarr=false
			strcarrera=rs("codcarr") & " " & rs("Nombre_l") 
		else
			if strcarrera=rs("codcarr") then
				sel=" selected"
			else
				sel=""
			end if
		end if
		
		cmbNumeros= cmbNumeros & "<option value=""" & rs("codcarr") & """ " & sel & ">"
		cmbNumeros=	cmbNumeros & rs("CodCarr") & "</option>" & vbcrlf
		rs.MoveNext 

	wend
end if

if AccionVar="Enviar" then
  session("CarreraAlumno")=strcarrera
  Accionvar=""
  response.Redirect "ValidaClave.asp"
end if
if AccionVar="Seleccionar" then
	strNumeros =  1
	AccionVar=""
end if

strParame="SELECT coalesce(dbo.Fn_ValorParame('FILTRASECCIONINICIALTRPA'),'')Parame"
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
	FILTRASECCIONINICIALTRPA=rstParame("Parame")
end if

%>
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=profe name=profe></OBJECT>
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=seccion name=seccion></OBJECT>
<html>
<head>
<title>Portal de Alumnos en L&iacute;nea</title> 
<META HTTP-EQUIV="Content-Type" CONTENT="text/html">  
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<%
   
   Dim RamosSelec,FiltraJornada,xramo,aux,x,il,ip
   Dim PuedeBotar, TipoCurso, CodSeccTeo
   
   CodSede = Session("CodSede")
   Ano = Session("Ano")
   Periodo = Session("Periodo")   
   CodCarr = Session("CodCarr")
   CodCarr2 = Request("CodCarr2")
   CodPestud = Session("CodPestud")
   Codcli = Session("Codcli") 
   if Session("CodCli") = "" then
     Response.Redirect("saltoinicio.htm")
   end if
   Ramos = Request("Ramos")
   if Trim(Ramos) = "" then
      Ramos = Session("Ramos")
   else
      Session("Ramos") = Ramos
   end if

	RamosSelec =Split(Ramos, ";", -1, 1) 
	'response.Write(Ramos)
	'response.Write(ramoselec) 
	'response.end 
	
	Dim PrioridadRamos(30)
	Dim PrimeraVez, Rst, LastRstSeccion
	Dim ArrCupos(300, 2)
	Dim ArrayRamos, ArrayRamosEsp(300,9)
	Dim s_CodPestudAdic
	Dim s_CodPestudAlu
	Dim TipoRamoOtros
	Dim s_RamoFueraMallaRep
	
	PrimeraVez = true
	ConConvalidacion = PoseeConvalidaciones(Codcli)
	EstiloBachillerato = false
	 
	if Trim(Request("Niv")) <> "" then
		NivelChosen = Request("Niv")
	else
		NivelChosen = -1
	end if
%>

<script languaje="javascript">
		function VerificaRamosDebe()
		{
		var i;
		var DebesNoInscritos = false;
		
		  return true;
		  
		  for (i = 0; i <= <%=Ubound(RamosSelec, 1)%>; i++)
		  {
			if ((RamoDebe[i] == 0) && RamoInscrito[i]) { 
			   alert("inscrito");
			}
		  }   
		   if (!confirm("No ha inscrito los ramos Debe... No podrá inscribir"))
		   {window.location="asignatura-seccion.asp?Ramos=<%=Ramos%>";}
		}
		
		function Valida()
		{
			window.location="inscrip-asigna-combos-det.asp";
		}
--></script>

<body OnUnload="javascript:VerificaRamosDebe();">
<table border="0" cellpadding="0" cellspacing="0" align="left">
  <tr>
	<td>
		<table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
		  <tr>
			<td valign="top" >
			<% CargarTop1()%><% SubMenu()%>
			 <table width="762" border="0" cellspacing="15" cellpadding="0">
				  <tr> 
					<td width="436"><img src="Imagenes/titulos/T-rev_secciones_horarios.gif" width="400" height="38"></td>
					<td width="281">&nbsp;</td>
				  </tr>
				  <%
				  
				  'StrSql = "select ObligaRamosFueraMalla from mt_parame "
				  'StrSql = StrSql & " where ObligaRamosFueraMalla='S'"
				  
				  StrSql = "select coalesce(dbo.Fn_ValorParame('ObligaRamosFueraMalla'),'N') as ObligaRamosFueraMalla"
				  
				  if BCL_ADO(StrSql, rst) then
					ObligaRamosFueraMallaRep =rst("ObligaRamosFueraMalla") 
				  end if
				  'response.Write(StrSql)
				  'response.End()
				  
				  StrSql = "select filtrajornada from mt_carrer "
				  StrSql = StrSql & " where codcarr='" & trim(CodCarr) & "'"
				  StrSql = StrSql & " and filtrajornada = 'S'"
				  
				  if BCL_ADO(StrSql, rst) then
					FiltraJornada ="S"
					rst.close()
				  else
					FiltraJornada = "N"
				  end if
								
				  StrSql = "select codpestud from mt_alumno "
				  StrSql = StrSql & " where codcli='" & trim(Codcli) & "'"
				  
				  if BCL_ADO(StrSql, rst) then
					s_CodPestudAlu =rst("codpestud")
					rst.close()
				  else
					s_CodPestudAlu = ""
				  end if
				  
				   'Analizaremos si es una carrera de Bachillerato
				   StrSql = "Select count (Distinct s.CodCarr) "
				   StrSql = StrSql & " From ra_optivo o, ra_seccio s, mt_carrer c, ra_carga d, mt_alumno e, ra_curric f"
				   StrSql = StrSql & " Where o.CodSede = '" & CodSede & "' "
				   StrSql = StrSql & " And o.CodCarr = '" & CodCarr & "' "
				   StrSql = StrSql & " And o.CodSede = s.CodSede "
				   StrSql = StrSql & " And o.CodRamo = s.CodRamo "
				   StrSql = StrSql & " And s.Ano = " & Ano
				   StrSql = StrSql & " And s.Periodo = '" & Periodo & "' "
				   StrSql = StrSql & " And o.Ano = " & Ano
				   StrSql = StrSql & " And o.Periodo = '" & Periodo & "' "
				   If FiltraJornada = "S" Then
					StrSql = StrSql & " And s.Jornada = '" & Session("Jornada") & "' "
				   end if
				   StrSql = StrSql & " And s.codCarr = c.CodCarr "
				   StrSql = StrSql & " And d.codcli = e.codcli "
				   StrSql = StrSql & " And d.codramo = f.codramo " 
				   StrSql = StrSql & " And  e.codpestud = f.codpestud  "
				   StrSql = StrSql & " And  d.codcli='" & Session("codcli") & "' "
				   StrSql = StrSql & " And  f.tipo = 'OPTATIVO ESPECIAL'"
				   StrSql = StrSql & " And  o.tipo = 'OPTATIVO ESPECIAL'"
				   StrSql = StrSql & " And  d.INSCRITO <> 'S' "
				   'StrSql = StrSql & " and  s.tipocurso='T'" 
				   if BCL_ADO(StrSql, rst) then
					 CantidadRamos = rst(0)
					 rst.close()
				   else
					 CantidadRamos = 0
				   end if
				   
				   StrSql = "Select Distinct s.CodCarr, c.nombre_l "
				   StrSql = StrSql & " From ra_optivo o, ra_seccio s, mt_carrer c , ra_carga d, mt_alumno e, ra_curric f"
				   StrSql = StrSql & " Where o.CodSede = '" & CodSede & "' "
				   StrSql = StrSql & " And o.CodCarr = '" & CodCarr & "' "
				   StrSql = StrSql & " And o.CodSede = s.CodSede "
				   StrSql = StrSql & " And o.CodRamo = s.CodRamo "
				   'StrSql = StrSql & " And s.CodCarr = '" & CodCarr2 & "' "
				   StrSql = StrSql & " And s.Ano = " & Ano
				   StrSql = StrSql & " And s.Periodo = '" & Periodo & "' "
				   StrSql = StrSql & " And o.Ano = " & Ano
				   StrSql = StrSql & " And o.Periodo = '" & Periodo & "' "
				   StrSql = StrSql & " And s.codCarr = c.CodCarr "
				   If FiltraJornada = "S" Then
					StrSql = StrSql & " And s.Jornada = '" & Session("Jornada") & "' "
				   end if
				   StrSql = StrSql & " And (o.codOpt = '' or o.codOpt is null) "
				   StrSql = StrSql & " And d.codcli = e.codcli "
				   StrSql = StrSql & " And d.codramo = f.codramo " 
				   StrSql = StrSql & " And  e.codpestud = f.codpestud  "
				   StrSql = StrSql & " And  d.codcli='" & Session("codcli") & "' "
				   StrSql = StrSql & " And  f.tipo = 'OPTATIVO ESPECIAL'"
				   StrSql = StrSql & " And  o.tipo = 'OPTATIVO ESPECIAL'"
				   StrSql = StrSql & " And  d.INSCRITO <> 'S' "
				   'response.Write(StrSql)
			if BCL_ADO(Strsql, Rst) then
			  if CantidadRamos > 1 then
						   EstiloBachillerato = true       
						   If trim(CodCarr2) = "" then
							 CodCarr2 = Rst("CodCarr")
						   end if
					%>
						  <tr> 
									<td width = "436" height="40"> 
								    <% Response.Write("<font size='1' face='Arial, Helvetica, sans-serif'>")
												 for i = 0 to ubound(RamosSelec, 1) - 1
													Response.Write("<a href='asignatura-seccion.asp?P=N&Niv=" & i & "&CodCarr2=" & CodCarr2 & "'>" & RamosSelec(i) & "</a>&nbsp&nbsp ")
												 next
												 Response.Write("</font>")
										  %>									</td>
					
									<td height="40"> 
									  <div align="right"><a href="inscrip-asigna-combos-det.asp" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('continuar-on','','imagenes/continuar-on.gif',1);"></a>									  </div>
									  <form name=frmCarrera>
										<input type="hidden" name="Niv" value="<%=NivelChosen%>">
										<select name="CodCarr2" OnChange="javascript:submit()">
											  <% Do while not Rst.eof%>
													  <%   If trim(CodCarr2) = Rst("CodCarr") then %>
													  <option value=<%=Rst("CodCarr")%> selected><%=Rst("Nombre_l")%></option>
													  <%   else %>
													  <option value=<%=Rst("CodCarr")%>><%=Rst("Nombre_l")%></option>
													  <%   end if 
													 Rst.movenext
												   loop 
											   Rst.close()  
											 %>
										</select>
									  </form>									</td>
							</tr>
						<%
						   else
						 %>
							<td> 
							  <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font size="1"><!--a href="javascript:;" onMouseOver="MM_openBrWindow('tex-seccionhorario.htm','seccionhorario','width=500,height=500')"><img src="imag/botones/instrucciones.gif" width="122" height="23" border="0"></a--></font></b></font></p>							</td>
							<td> 
							  <div align="right">
							  <a href="javascript:Valida();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','Imagenes/botones/continuar-on.gif',1)"><img src="Imagenes/botones/continuar-of.gif" width="178" height="45" name="Image1" border="0"></a></div>							</td>
							</tr>
					 <%  
						   end if
			else
				%>
						<td height="21"> 
						  <p><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b><font size="1"><!--a href="javascript:;" onMouseOver="MM_openBrWindow('tex-seccionhorario.htm','seccionhorario','width=500,height=500')"><img src="imag/botones/instrucciones.gif" width="122" height="23" border="0"></a--></font></b></font></p>						</td>
						<td> 
						  <div align="right"><a href="javascript:Valida();" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','Imagenes/botones/continuar-on.gif',1)"><img   src="Imagenes/botones/continuar-of.gif" width="178" height="45" name="Image1" border="0"></a>						  </div>						</td>
					</tr>
				  <%          
			end if  
			%>
			</table>
			
			<%
			'dim xramos as string
			EstiloBachillerato = False
			if Nivelchosen < 0 then
			if EstiloBachillerato then
			Min = 0
			Max = 0
			else
			Min  = 0
			Max = ubound(RamosSelec, 1) - 1
			end if
			else   
			Min = NivelChosen
			Max = NivelChosen
			end if
			
			for i = Min to Max
			il=i
			ip=i	
					   PrioridadRamos(i) = PrioridadRamo(Codcli, RamosSelec(i), Ano, Periodo)
					%>
					<% ' Recorremos las secciones 
					 EsRepit = EsRepitente(RamosSelec(i), CodCli) 	
					 if EsRepit = true then
						EsRepit_P = 1
					 else
						EsRepit_P = 0
					 end if			 
					 Session("EsRepit") = EsRepitente(RamosSelec(i), CodCli)
					 TipoRamo = trim(GetTipoRamo(RamosSelec(i), CodPestud))     
					 Inconsistencia = false
					 'Se supone que el mismo ramo malla no puede estar en mas
					 'de una malla para el mismo alumno, en mallas otros ramos
					 s_CodPestudAdic = trim(GetPlanOtrosRamos(s_CodPestudAlu,RamosSelec(i)))
					 session("TipoRamoOtros") = GetTipoRamo(RamosSelec(i), s_CodPestudAdic)
					 'response.Write(session("TipoRamoOtros"))
					 s_RamoFueraMallaRep = ""
					 if ObligaRamosFueraMallaRep="S" then
						s_RamoFueraMallaRep = GetRamoFueraMallaRep(RamosSelec(i), CodCli,Ano,Periodo)
					 end if
					 
					 'response.write(RamosSelec(i))
					 RamosSelec(i)=trim(RamosSelec(i))
					 'if RamosSelec(i)="COCOFI0019" then
						  'response.write("Hola")
			'			  response.end					 	
					 'end if
					 select Case TipoRamo
					   case ""'otros ramos
								   StrSeccion = "Select distinct a.codramo,a.codsecc, b.nombre, a.cupo, e.CodProf, p.Ap_Pater, p.nombres, e.horario, e.tipocurso,e.codseccteo  " 
								   StrSeccion = StrSeccion & " from ra_optivo a,ra_ramo b,ra_seccio e, ra_profes p" 
								   StrSeccion = StrSeccion & " where a.codcarr  = '" & CodCarr & "' " 
								   StrSeccion = StrSeccion & " and a.codramo=b.codramo" 
								   StrSeccion = StrSeccion & " and a.ano=" & Ano & "" 
								   StrSeccion = StrSeccion & " and a.periodo=" & Periodo & "" 
								   StrSeccion = StrSeccion & " and a.codramo=e.codramo" 
								   StrSeccion = StrSeccion & " and a.codsecc=e.codsecc" 
									   If FiltraJornada = "S" Then
										   StrSeccion = StrSeccion & " And e.Jornada = '" & Session("Jornada") & "'"
									   end if                  
								   StrSeccion = StrSeccion & " and a.ano=e.ano" 
								   StrSeccion = StrSeccion & " and a.periodo=e.periodo" 
								   StrSeccion = StrSeccion & " and e.CodSede = '" & CodSede & "' "
								   if session("TipoRamoOtros") = "OBLIGATORIO" then 'Ingles (otros ramo con requisito)
											   StrSeccion = StrSeccion & " and a.codramo = '" & RamosSelec(i) & "' "
											   StrSeccion = StrSeccion & " and a.codopt = '' "
								   else 'Otros ramos sin requisito
										   StrSeccion = StrSeccion & " and a.codopt = '" & RamosSelec(i) & "' "  	
										   if trim(s_RamoFueraMallaRep)<>"" then
												   StrSeccion = StrSeccion & " and a.codramo = '" & trim(s_RamoFueraMallaRep) & "' "
										   end if
								   end if	 
								   StrSeccion = StrSeccion & " and e.CodProf *= p.CodProf "
								   StrSeccion = StrSeccion & " and a.tipo <> 'OPTATIVO ESPECIAL'"
								   
								   if FILTRASECCIONINICIALTRPA ="SI" then
								   		StrSeccion = StrSeccion & " and e.codsecc = dbo.FN_CODSECC_INICIAL_ALUMNO('"& session("codcli") &"') "
								   end if 
								   
								   StrSeccion = StrSeccion & " Order by  E.TIPOCURSO DESC, a.CodSecc "
								   EsEspecial = false
						case "OBLIGATORIO"
						
								 if EsPlanificadoXJornada(RamosSelec(i), CodSede, Ano, Periodo, CodCarr,FiltraJornada) then
										   'falta relacionarlo con la ra_optivo	
										   StrSeccion = "Select e.CodRamo, e.CodSecc, r.nombre, 0 as cupo, e.CodProf, p.Ap_Pater, p.Nombres, e.horario, e.tipocurso,e.codseccteo "
										   StrSeccion = StrSeccion & " from ra_seccio e, ra_profes p, ra_ramo r "
										   StrSeccion = StrSeccion & " Where e.CodRamo = '" & RamosSelec(i) & "' and e.CodSede = '" & CodSede & "' " 
										   StrSeccion = StrSeccion & " And e.Ano = " & ano & " And e.Periodo = " & Periodo & ""
										   StrSeccion = StrSeccion & " and e.CodProf *= p.CodProf "
										   StrSeccion = StrSeccion & " and e.CodCarr='" & trim(Codcarr) & "'" 
										   If FiltraJornada = "S" Then
											   StrSeccion = StrSeccion & " And (e.Jornada = '" & Session("Jornada") & "' OR e.Jornada2 = '" & Session("Jornada") & "')"
										   end if                  
										   StrSeccion = StrSeccion & " and e.codramo = r.codramo" 
										   
										   if FILTRASECCIONINICIALTRPA ="SI" then
												StrSeccion = StrSeccion & " and e.codsecc = dbo.FN_CODSECC_INICIAL_ALUMNO('"& session("codcli") &"') "
										   end if
								   
										   StrSeccion = StrSeccion & " Order by  E.TIPOCURSO DESC, e.CodSecc "
								 else
									   Equivalente = GetEquivalente(RamosSelec(i), CodSede, Ano, Periodo, CodCarr)
									   if Trim(Equivalente) <> "" then
									 '  			response.Write("Con Equivalente :" & Equivalente)
											   StrSeccion = "Select e.CodRamo, e.CodSecc, r.nombre, 0 as cupo, e.CodProf, p.Ap_Pater, p.Nombres, e.horario, e.tipocurso,e.codseccteo "
											   StrSeccion = StrSeccion & " from ra_seccio e, ra_profes p, ra_ramo r "
											   'ANTES
											   'StrSeccion = StrSeccion & " Where e.CodRamo = '" & Equivalente & "'"
											   
											   'Mejora 2.0,  SEGUIR !!
											   StrSeccion = StrSeccion & " Where e.CodRamo in (Select a.ramoequiv from ra_equiv a with (nolock), mt_carrer b with (nolock) Where a.codcarr = b.codcarr And b.Sede = '" & CodSede & "' And a.CodRamo = '" & RamosSelec(i) & "' and a.codcarr = '" & trim(CodCarr) & "')" 					    
											   StrSeccion = StrSeccion & " and e.CodSede = '" & CodSede & "' " 
											   StrSeccion = StrSeccion & " And e.Ano = " & ano & " And e.Periodo = " & Periodo
											   StrSeccion = StrSeccion & " and e.CodProf *= p.CodProf "
												   If FiltraJornada = "S" Then
													   StrSeccion = StrSeccion & " And (e.Jornada = '" & Session("Jornada") & "' or e.Jornada2 = '" & Session("Jornada") & "')"
												   end if                  
											   StrSeccion = StrSeccion & " and e.codramo = r.codramo" 
											   
											   if FILTRASECCIONINICIALTRPA ="SI" then
													StrSeccion = StrSeccion & " and e.codsecc = dbo.FN_CODSECC_INICIAL_ALUMNO('"& session("codcli") &"') "
											   end if
											   
											   StrSeccion = StrSeccion & " Order by  E.TIPOCURSO DESC, e.CodSecc "
			
									   else
									'   		response.Write("Sin Equivalencias :" & Equivalente)
											   StrSeccion = "Select distinct a.codramo,a.codsecc, b.nombre, a.cupo, e.CodProf, p.Ap_Pater, p.nombres, e.horario, e.tipocurso,e.codseccteo " 
											   StrSeccion = StrSeccion & " from ra_optivo a,ra_ramo b,ra_seccio e, ra_profes p" 
											   StrSeccion = StrSeccion & " where a.codcarr  = '" & CodCarr & "' " 
											   StrSeccion = StrSeccion & " and a.codramo=b.codramo" 
											   StrSeccion = StrSeccion & " and a.ano=" & Ano & "" 
											   StrSeccion = StrSeccion & " and a.periodo=" & Periodo & "" 
											   StrSeccion = StrSeccion & " and a.codramo=e.codramo" 
											   StrSeccion = StrSeccion & " and a.codsecc=e.codsecc" 
											   If FiltraJornada = "S" Then
												   StrSeccion = StrSeccion & " And (e.Jornada = '" & Session("Jornada") & "' or e.Jornada2 = '" & Session("Jornada") & "')"
											   end if                  
											   StrSeccion = StrSeccion & " and a.ano=e.ano" 
											   StrSeccion = StrSeccion & " and a.periodo=e.periodo" 
											   StrSeccion = StrSeccion & " and e.CodSede = '" & CodSede & "' "
											   StrSeccion = StrSeccion & " and a.CodSede = e.CodSede"
											   StrSeccion = StrSeccion & " And a.tipo <> 'OPTATIVO ESPECIAL'"				
											   StrSeccion = StrSeccion & " and a.codopt = '" & RamosSelec(i) & "' "
											   if trim(s_RamoFueraMallaRep)<>"" then
												   StrSeccion = StrSeccion & " and a.codramo = '" & trim(s_RamoFueraMallaRep) & "' "
											   end if
											   StrSeccion = StrSeccion & " and e.CodProf *= p.CodProf "
											   
											   if FILTRASECCIONINICIALTRPA ="SI" then
													StrSeccion = StrSeccion & " and e.codsecc = dbo.FN_CODSECC_INICIAL_ALUMNO('"& session("codcli") &"') "
											   end if
											   
											   StrSeccion = StrSeccion & " Order by  E.TIPOCURSO DESC, a.CodSecc "
											   'response.Write(StrSeccion)		 'SEGUIR !					  
										end if
								 end if 
								
								 EsEspecial = false
					   case "OPCIONAL"
									   StrSeccion = "Select distinct a.codramo,a.codsecc, b.nombre, a.cupo, e.CodProf, p.Ap_Pater, p.nombres, e.horario, e.tipocurso,e.codseccteo  " 
									   StrSeccion = StrSeccion & " from ra_optivo a,ra_ramo b,ra_seccio e, ra_profes p" 
									   StrSeccion = StrSeccion & " where a.codcarr  = '" & CodCarr & "' " 
									   StrSeccion = StrSeccion & " and a.codramo=b.codramo" 
									   StrSeccion = StrSeccion & " and a.ano=" & Ano & "" 
									   StrSeccion = StrSeccion & " and a.periodo=" & Periodo & "" 
									   StrSeccion = StrSeccion & " and a.codramo=e.codramo" 
									   StrSeccion = StrSeccion & " and a.codsecc=e.codsecc" 
											   If FiltraJornada = "S" Then
												   StrSeccion = StrSeccion & " And e.Jornada = '" & Session("Jornada") & "'"
											   end if                  
									   StrSeccion = StrSeccion & " and a.ano=e.ano" 
									   StrSeccion = StrSeccion & " and a.periodo=e.periodo" 
									   StrSeccion = StrSeccion & " and e.CodSede = '" & CodSede & "' "
									   StrSeccion = StrSeccion & " and a.CodSede = e.CodSede"
											   if trim(s_RamoFueraMallaRep)<>"" then
												   StrSeccion = StrSeccion & " and a.codramo = '" & trim(s_RamoFueraMallaRep) & "' "
											   end if
									   StrSeccion = StrSeccion & " and e.CodProf *= p.CodProf "
									   StrSeccion = StrSeccion & " and a.tipo <> 'OPTATIVO ESPECIAL'"
									   StrSeccion = StrSeccion & " and a.codopt = '" & RamosSelec(i) & "' "  
									   
									   if FILTRASECCIONINICIALTRPA ="SI" then
											StrSeccion = StrSeccion & " and e.codsecc = dbo.FN_CODSECC_INICIAL_ALUMNO('"& session("codcli") &"') "
									   end if
									   
									   StrSeccion = StrSeccion & " Order by  E.TIPOCURSO DESC, a.CodSecc "
									   EsEspecial = False
									   'response.write(StrSeccion)
					   case "OPTATIVO ESPECIAL"  
									   StrSeccion = "Select distinct a.codramo,a.codsecc, b.nombre, a.cupo, e.CodProf, p.Ap_Pater, p.nombres, e.horario, e.tipocurso,e.codseccteo " 
									   StrSeccion = StrSeccion & " from ra_optivo a,ra_ramo b,ra_seccio e, ra_profes p" 
									   StrSeccion = StrSeccion & " where a.codcarr  = '" & CodCarr & "' " 
									   StrSeccion = StrSeccion & " and a.codramo=b.codramo" 
									   StrSeccion = StrSeccion & " and a.ano=" & Ano & "" 
									   StrSeccion = StrSeccion & " and a.periodo=" & Periodo & "" 
									   StrSeccion = StrSeccion & " and a.codramo=e.codramo" 
									   StrSeccion = StrSeccion & " and a.codsecc=e.codsecc" 
									   StrSeccion = StrSeccion & " and a.ano=e.ano" 
									   If FiltraJornada = "S" Then
										   StrSeccion = StrSeccion & " And e.Jornada = '" & Session("Jornada") & "'"
									   end if                  
									   StrSeccion = StrSeccion & " and a.periodo=e.periodo" 
									   StrSeccion = StrSeccion & " and e.CodSede = '" & CodSede & "' "
									   StrSeccion = StrSeccion & " and a.CodSede = '" & CodSede & "' "
									   StrSeccion = StrSeccion & " and e.CodProf *= p.CodProf "
									   if trim(CodCarr2) <> "" then
										 StrSeccion = StrSeccion & " and e.CodCarr = '" & CodCarr2 & "' "
									   end if
									   StrSeccion = StrSeccion & " and a.tipo = 'OPTATIVO ESPECIAL'"
									   StrSeccion = StrSeccion & " and a.codopt = '" & RamosSelec(i) & "' "  
									   if trim(s_RamoFueraMallaRep)<>"" then
										   StrSeccion = StrSeccion & " and a.codramo = '" & trim(s_RamoFueraMallaRep) & "'"
									   end if
									   StrSeccion = StrSeccion & " and a.codramo not in (select a.codramo from ra_secuen a, ra_planesadic b"
									   StrSeccion = StrSeccion & " where a.codpestud=b.codpestudadic "
									   StrSeccion = StrSeccion & " and b.codpestud='" & trim(s_CodPestudAlu) & "'"
									   StrSeccion = StrSeccion & " and a.requisito not in "
									   StrSeccion = StrSeccion & " (select ramoequiv from ra_nota "
									   StrSeccion = StrSeccion & " where (estado='A' or estado='E' or estado='I') "
									   StrSeccion = StrSeccion & " and codcli='" & trim(Codcli) & "'))"
									   
									   if FILTRASECCIONINICIALTRPA ="SI" then
											StrSeccion = StrSeccion & " and e.codsecc = dbo.FN_CODSECC_INICIAL_ALUMNO('"& session("codcli") &"') "
									   end if
									   
									   StrSeccion = StrSeccion & " Order by  E.TIPOCURSO DESC, a.CodSecc "
									   EsEspecial = true
									   PrimeraVez = true 
									   'response.write(StrSeccion)
					end select 
					'response.Write(StrSeccion)
					'if RamosSelec(i)="COCOFI0019" then
						'response.write(StrSeccion)					 	
					'end if 

					 If EsEspecial then
								  
								  'response.end	
								if PrimeraVez then
										  if BCL_ADO(StrSeccion, RstSeccion) then
											  CodSeccTeo = valnulo(rstSeccion("codseccteo"), num_)
											  TipoCurso = valnulo(rstSeccion("tipocurso"), str_)
											  if TipoCurso = "T" then
													TextoTipoCurso ="Te&oacute;rico"
											  elseif TipoCurso = "P" then
													TextoTipoCurso ="Practico"
											  elseif TipoCurso = "L" then
													TextoTipoCurso ="Laboratorio"
											  end if
											  xramo=valnulo(rstSeccion("CodRamo"), str_)
											  PrimeraVez = false
											  ExistenRegistros = not RstSeccion.eof
											  ExistenRegistrosEs = ExistenRegistros
											  RevisaCupos = true
											  
										  end if	
								else
										  'Ejecuta query anterior	
										  
										  ExistenRegistros = ExistenRegistrosEs
										  RstSeccion.movefirst
										  xramo=valnulo(rstSeccion("CodRamo"), str_)
										  RevisaCupos = false
								end if
					 else
								if not inconsistencia then
								  if StrSeccion <> "" then
									  Set RstSeccion = Session("Conn").Execute (StrSeccion)
									  ExistenRegistros = not RstSeccion.eof
								  end if
								else 
								  ExistenRegistros = false
								end if
					 end if
					 if clng(Session("AnoAlumno"))  = clng(Session("Ano")) then 
						' Es de primer ano, solo debe inscribir los optativos
								if TipoRamo <> "OBLIGATORIO" then
										paso = true
								else
										if EsAntiguo then
												paso = true
										else                      
												paso = false
										end if
								end if
					 else
								paso = true
					 end if
						 if inconsistencia then
						   paso = false
						 end if
					 if ExistenRegistros then
							  TipoCurso = valnulo(rstSeccion("tipocurso"), str_)
							  xramo=valnulo(rstSeccion("Codramo"),str_)												  
								PosRamos = 0
								if EsEspecial then
								  if PrimeraVez then
									ArrayRamos = RstSeccion.GetRows 
									 CodSeccTeo = valnulo(rstSeccion("codseccteo"), num_)
								  else 
										ArrayRamos = RstSeccion.GetRows 
								  end if
								 else
										ArrayRamos = RstSeccion.GetRows 
								 end if
				
				if TipoCurso="T"  then	 
			%>
					<form  id=form1<%=RamosSelec(i)%>  name=<%=RamosSelec(i)%>>
					  <table width="282" border="0" cellspacing="15" cellpadding="0">
					<tr valign="top"> 
					  
				  <td colspan="2" height="83"><table width="750" cellspacing="1" cellpadding="0" height="106" border="0" bordercolor="#FFFFFF">
					  <tr bgcolor="83a3d0"> 
						<td height="22" colspan="8" background="Imagenes/fdo-cabecera-cel.jpg"> <div align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2" name="lblAsig1">Asignatura: </font></b></font><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2">
							<% if TipoRamo<>"" then
													response.Write(RamosSelec(i) + " " + GetNombreRamo(RamosSelec(i)))
												else
													response.Write(RamosSelec(i) & " " & replace(GetNombreRamo(RamosSelec(i)),"FORMACION GENERAL","CURSO OPTATIVO ADICIONAL"))
												end if%>
							</font></b></font></div></td>
					  </tr>
					  <tr background="Imagenes/fdo-cabecera-cel22.jpg" height="30"> 
						<td width="80" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Lista 
						Curso</font></b></font></div></td>
						<td width="181" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1" name="lblColumnaProfesor">Profesor</font></b></font></div></td>
						<td width="75" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">S</font><font size="1">ecci&oacute;n</font></b></font></div> <div align="center"></div></td>
						<td width="141" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1" name="lblColumnaRamo">Ramo</font></b></font></div></td>
						<td width="34" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1"> 
						Cupo</font></b></font></div></td>
						<td width="72" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Inscribir</font></b></font></div></td>
						<td width="101" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1" name="lblColumnaTipoRamo">Tipo Asignatura </font></b></font></div></td>
						<td width="57" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1"></font></b></font></div></td>
					  </tr>
					  <%             
									 x=0
									 do while posRamos <= Ubound(ArrayRamos, 2) and TipoCurso="T" 'and x=0
									 
											 TextoTipoCurso ="Te&oacute;rico"
											 session("OtrosRamos") = "N"						                   
											 TipoCupo = "CT"
												 Select Case TipoRamo
												 Case ""
													  if session("TipoRamoOtros") = "OBLIGATORIO" then
														  session("OtrosRamos") = "S"						  
													  end if 
													  Cupos = SacaCupoAntiguo (ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, CodCarr, CodSede, EsRepit, TipoCupo,RamosSelec(i))
													  if session("RamoObliPlanificadoOp") = "N" then
														  Inscritos = CantidadInscritos ( ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), CodSede, Ano, Periodo, TipoCupo, TipoCurso )	
													  else
														  Inscritos = CantidadInscritosOptativos(RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), CodSede, Ano, Periodo, CodCarr, TipoCupo, TipoCurso) 							  	  	
													  end if	
												  Case "OBLIGATORIO"
												  
												  		'strCuposX = "SacaCupoAntiguo @Ramo='"& ArrayRamos(0, posRamos) &"',@Seccion="& ArrayRamos(1, posRamos) &",@Ano="& Ano &",@Periodo="& Periodo &",@Codcarr='"& CodCarr &"',@G_SEDE='"& CodSede &"',@EsRepit= "& EsRepit_P &",@Es_CarreraAlumno=1,@RamoMalla='"& RamosSelec(i) &"',@TipoInscripcion=20,@TipoCupo='"& TipoCupo &"',@lc_CodCarrSecc='',@OtrosRamos='',@CupoCero='',@RamoObliPlanificadoOp='' "
														
														'set rs=Session("Conn").Execute (sql)														
														'if BCL_ADO(strCuposX, rstCuposX) then
															'Cupos = valnulo(rstCuposX(0), num_)
														'end if 	 
														Cupos = SacaCupoAntiguo (ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, CodCarr, CodSede, EsRepit, TipoCupo,RamosSelec(i))					
																						 
														 if session("RamoObliPlanificadoOp") = "N" then
															  Inscritos = CantidadInscritos ( ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), CodSede, Ano, Periodo, TipoCupo, TipoCurso )
														 else
															  Inscritos = CantidadInscritosOptativos(RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), CodSede, Ano, Periodo, CodCarr, TipoCupo, TipoCurso)															  
														 end if	
														 'response.Write(" Ramo: " & ArrayRamos(0, posRamos) )
														 'response.Write(" CupoCero: " & session("CupoCero") )
														 'response.Write(" RamoObliPlanificadoOp: " & session("RamoObliPlanificadoOp") )
														 'response.Write( "TipoCupo: " & TipoCupo)
														 'response.Write(" Cupos: " & Cupos )
														 'response.Write(" Inscritos: " & Inscritos )
														 
												   Case "OPCIONAL"
														  Cupos = SacaCupoOptativo(RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, CodCarr,CodSede) 
														  Inscritos = CantidadInscritosOptativos(RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), CodSede, Ano, Periodo, CodCarr, TipoCupo, TipoCurso)
												   Case "OPTATIVO ESPECIAL"
														  Inscritos=0
														  if revisaCupos then
															Cupos = SacaCupoOptativoEspecial(ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, CodCarr,RamosSelec(i)) 
															Inscritos = CantidadInscritosOptativoEspecial(ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), CodSede, Ano, Periodo, CodCarr, TipoCupo, TipoCurso)
															ArrCupos(posRamos, 1) = Cupos
															ArrCupos(posRamos, 2) = Inscritos                          
														  else
															Cupos = ArrCupos(posRamos, 1) 
															Inscritos = ArrCupos(posRamos, 2) 
														  end if
												 End Select
													  %>
                      <%
					strParame="SELECT dbo.Fn_ValorParame('VALIDAHORARIOTRPORTAL')Parame"
					if BCL_ADO(strParame, rstParame) then
						ValidaHorario=rstParame("Parame")
					else
						ValidaHorario=""
					end if 
					  
					if ValidaHorario="SI" then  
					
						strHorari="select top 1 1 from dbo.RA_HORARI WHERE codramo='"&  ArrayRamos(0, posRamos) &"' AND ano='"& Ano &"' AND periodo='"& Periodo &"' AND CODSECC="& ArrayRamos(1, posRamos) &" and codsede ='"& session("codsede") &"'"
						
						if BCL_ADO(strHorari, rstHorari) then
							TieneHorario="SI"
						else
							TieneHorario="NO"
						end if 
					else
						TieneHorario="SI" 
					end if 
					  
					  if TieneHorario="SI" then%>
					  <tr bgcolor="#DBECF2" > 
						<td width="80" height="20" align="center" bgcolor="#DBECF2"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="detalle-asignatura.asp?R=<%=ArrayRamos(0, posRamos)%>&S=<%=ArrayRamos(1, posRamos)%>&P=<%=ArrayRamos(5, posRamos) + " " + ArrayRamos(6, posRamos)%> " target="_blank"><%=ArrayRamos(0, posRamos)%></a></font></div></td>
						<td width="181" height="16" align="center"> <div align="center"><font face="Verdana" size="1"><%=ArrayRamos(5, posRamos) + " " + ArrayRamos(6, posRamos)%></font></div></td>
						<td height="16" align="center" bgcolor="#DBECF2"> <div align="center"><font face="Verdana" size="1"><%=ArrayRamos(1, posRamos)%></font></div> <div align="center"></div></td>
						<td width="141" height="16" align="center"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"></font><font face="Verdana" size="0"><%=ArrayRamos(2, posRamos)%></font></div></td>
						<td width="34" height="16" align="center"> <div align="center"><font face="Verdana" size="1"> 
							<% 
																						If Inscritos >= Cupos then
																									if EstaPreInscrito(Codcli, RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, TipoCurso) then
																										PuedeBotar = True 					
																									else
																										PuedeBotar = False 					
																									end if
																							Response.Write("No")
																						else
																								  Response.Write("Si")
																								  PuedeBotar = True 					
																						end if
																					  %>
						</font></div></td>
						<td width="72" height="16"> <div align="center">
							<p><font face="Verdana" size="1"> 
							  <%
																									 EstadoReal = EstadoRamoReal(CodCli, ArrayRamos(0, posRamos),TipoCurso)
																									If (EstadoReal = "A") or (EstadoReal = "E") or (EstadoReal = "I") then
																									   Cursado = True
																									 else
																									   Cursado = False
																									 end if 
																							  %>
							  <% 
							  strPre="select top 1 1 from ra_Carga where codcli ='"& session("codcli") &"' and ano = "& ano &" and periodo = "& periodo &" and codramo = '"& ArrayRamos(0, posRamos) &"' and preinscrito = 'S'"
								if BCL_ADO(strPre, rstPre) then
									BloqueaxPreinscrito="DISABLED"
								else
									BloqueaxPreinscrito=""
								end if 
							   
							  if EstaPreInscrito(Codcli, RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, TipoCurso) then %>
							  </font><font face="Verdana" size="1">
									<input type="radio" id="<%=ArrayRamos(0, posRamos)%>" name="<%=RamosSelec(iP)%>" value="false" OnClick="javascript:GrabaSel(this,'<%=RamosSelec(i)%>', '<%=ArrayRamos(0, posRamos)%>', '<%=ArrayRamos(1, posRamos)%>',<%=i%>,'<%=mid(TextoTipoCurso ,1,1)%>','<%=ArrayRamos(9, posRamos)%>');javascript:marca(this, '<%=trim(ArrayRamos(0, posRamos))%>')" CHECKED <%=BloqueaxPreinscrito%>>
							  <% else %>
									<input type="radio" id="<%=ArrayRamos(0, posRamos)%>" name="<%=RamosSelec(iP)%>" value="<%=ArrayRamos(0, posRamos)%>"  OnClick="javascript:GrabaSel(this, '<%=RamosSelec(i)%>', '<%=ArrayRamos(0, posRamos)%>', '<%=ArrayRamos(1, posRamos)%>',<%=i%>,'<%=mid(TextoTipoCurso ,1,1)%>','<%=ArrayRamos(9, posRamos)%>');javascript:marca(this, '<%=trim(ArrayRamos(0, posRamos))%>')" <%If  (Inscritos >= Cupos) or Cursado then Response.Write("DISABLED") else Response.Write(BloqueaxPreinscrito)End if%> >
							  
							  <% end if %>
							  </font> </p>
						</div></td>
						<td width="101" ><center> <p><font face="Verdana" size="1"><%=TextoTipoCurso%></font><font face="Verdana" size="1"></font></p></center></td>
						<td width="57" valign="top"><center><p><font face="Verdana" size="1"></font><font face="Verdana" size="1">
                        <%
						strParameC="SP_TRAE_COMBOS_RAMOS '"& session("Codcli") &"','"& ArrayRamos(0, posRamos) &"',"& ArrayRamos(1, posRamos) &","& Session("Ano") &", "& Session("Periodo") &" "
						if BCL_ADO(strParameC, rstParameC) then
							combosRamos=rstParameC("combosRamos")
						end if 
						response.Write(combosRamos)
						%>
                        </font></p></td>
					  </tr>
                      <%else  'sin horario !!%>
                      <tr bgcolor="#99D7CD" > 
						<td width="80" height="20" align="center" bgcolor="#99D7CD"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="detalle-asignatura.asp?R=<%=ArrayRamos(0, posRamos)%>&S=<%=ArrayRamos(1, posRamos)%>&P=<%=ArrayRamos(5, posRamos) + " " + ArrayRamos(6, posRamos)%> " target="_blank"><%=ArrayRamos(0, posRamos)%></a></font></div></td>
						<td width="181" height="16" align="center"> <div align="center"><font face="Verdana" size="1"><%=ArrayRamos(5, posRamos) + " " + ArrayRamos(6, posRamos)%></font></div></td>
						<td height="16" align="center" bgcolor="#99D7CD"> <div align="center"><font face="Verdana" size="1"><%=ArrayRamos(1, posRamos)%></font></div> <div align="center"></div></td>
						<td width="141" height="16" align="center"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"></font><font face="Verdana" size="0"><%=ArrayRamos(2, posRamos)%></font></div></td>
						<td width="34" height="16" align="center"> <div align="center"><font face="Verdana" size="1"> 
							<% 
																						If Inscritos >= Cupos then
																									if EstaPreInscrito(Codcli, RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, TipoCurso) then
																										PuedeBotar = True 					
																									else
																										PuedeBotar = False 					
																									end if
																							Response.Write("No")
																						else
																								  Response.Write("Si")
																								  PuedeBotar = True 					
																						end if
																					  %>
						</font></div></td>
						<td width="72" height="16"> <div align="center">
							<p><font face="Verdana" size="1"> 
							  <%
																									 EstadoReal = EstadoRamoReal(CodCli, ArrayRamos(0, posRamos),TipoCurso)
																									If (EstadoReal = "A") or (EstadoReal = "E") or (EstadoReal = "I") then
																									   Cursado = True
																									 else
																									   Cursado = False
																									 end if 
																							  %>
							  
									<input id="<%=ArrayRamos(0, posRamos)%>" type="radio" name="<%=RamosSelec(iP)%>" value="<%=ArrayRamos(0, posRamos)%>" DISABLED >
							 
							  </font> </p>
						</div></td>
						<td width="101" ><center> <p><font face="Verdana" size="1"><%=TextoTipoCurso%></font><font face="Verdana" size="1"></font></p></center></td>
						<td width="57" valign="top" align="center"> <p><font face="Verdana" size="1">SIN HORARIO</font><font face="Verdana" size="1"></font></p></td>
					  </tr>
                      <%end if 'fin sin horario%>
					  <%		posRamos = posRamos + 1
								if posramos <=   Ubound(ArrayRamos, 2) then
									TipoCurso = ArrayRamos(8, posRamos)
								end if
								
								loop
						  %>
					</table></td>
			  </tr>
					</table>
			  </form>
													   <%    
				END IF									   
							 
							 if  TipoCurso="P"  then
%>

				<form  id=form1<%=RamosSelec(i)%>  name=<%=RamosSelec(i)%>>
			<table width="707" border="0" cellspacing="15" cellpadding="0">
			<tr valign="top"> 
			  <td width="707" height="83" colspan="2"> 
				  <table width="750" cellspacing="0" cellpadding="0" height="106" border="1" bordercolor="#FFFFFF">
					  <tr bgcolor="83a3d0"> 
								<td height="22" colspan="8"> <div align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2" name="lblAsig2">Asignatura: </font></b></font><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2">
									<% if TipoRamo<>"" then
											response.Write(RamosSelec(i) + " " + GetNombreRamo(RamosSelec(i)))
										else
											response.Write(RamosSelec(i) & " " & replace(GetNombreRamo(RamosSelec(i)),"FORMACION GENERAL","CURSO OPTATIVO ADICIONAL"))
										end if%></font></b></font></div>								</td>
					  </tr>
					  <tr background="Imagenes/fdo-cabecera-cel22.jpg"> 
								<td width="97" height="30" background="Imagenes/fdo-cabecera-cel22.jpg" >
									   <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Lista Curso</font></b></font></div>						</td>
								<td width="203" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
									   <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Profesor</font></b></font></div>						</td>
								<td width="60" height="30" background="Imagenes/fdo-cabecera-cel22.jpg">
										<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">S</font><font size="1">ecci&oacute;n</font></b></font></div>
										<div align="center"></div></td>
								<td width="130" height="30" background="Imagenes/fdo-cabecera-cel22.jpg">
										 <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Ramo</font></b></font></div>						</td>
								<td width="32" height="30" background="Imagenes/fdo-cabecera-cel22.jpg">
										 <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1"> 
											Cupo</font></b></font></div>						</td>
								<td width="41" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
										  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Inscribir</font></b></font></div>						</td>
								<td width="75" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Tipo 
										  Asignatura </font></b></font>					</td>
								
						<td width="94" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif" >
						  <center>
							<b>Corresponde Seccion Teorica </b>
			</center></font> </td>
					  </tr>
			  <% 	x=0
					 do while posRamos <= Ubound(ArrayRamos, 2) and TipoCurso="P" 'AND X=0
					 TextoTipoCurso ="Práctico"
					 session("OtrosRamos") = "N"						                   
					 TipoCupo = "CT"
					 Select Case TipoRamo
						 Case ""
						  if session("TipoRamoOtros") = "OBLIGATORIO" then
							  session("OtrosRamos") = "S"						  
						  end if
						  Cupos = SacaCupoAntiguo (ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, CodCarr, CodSede, EsRepit, TipoCupo,RamosSelec(i))
						  if session("RamoObliPlanificadoOp") = "N" then
							  Inscritos = CantidadInscritos ( ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), CodSede, Ano, Periodo, TipoCupo, TipoCurso )
						  else
							  Inscritos = CantidadInscritosOptativos(RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), CodSede, Ano, Periodo, CodCarr, TipoCupo, TipoCurso) 							  	  	
						  end if	
								Case "OBLIGATORIO"
										 Cupos = SacaCupoAntiguo (ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, CodCarr, CodSede, EsRepit, TipoCupo,RamosSelec(i))
										  if session("RamoObliPlanificadoOp") = "N" then
											  Inscritos = CantidadInscritos ( ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), CodSede, Ano, Periodo, TipoCupo, TipoCurso )
										  else
											  Inscritos = CantidadInscritosOptativos(RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), CodSede, Ano, Periodo, CodCarr, TipoCupo, TipoCurso)
										  end if	
							   Case "OPCIONAL"
										  Cupos = SacaCupoOptativo(RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, CodCarr,CodSede) 
										  Inscritos = CantidadInscritosOptativos(RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), CodSede, Ano, Periodo, CodCarr, TipoCupo, TipoCurso)
							   Case "OPTATIVO ESPECIAL"
										  Inscritos=0
										  if revisaCupos then
											Cupos = SacaCupoOptativoEspecial(ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, CodCarr,RamosSelec(i)) 
											Inscritos = CantidadInscritosOptativoEspecial(ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), CodSede, Ano, Periodo, CodCarr, TipoCupo, TipoCurso)
											ArrCupos(posRamos, 1) = Cupos
											ArrCupos(posRamos, 2) = Inscritos                          
										  else
											Cupos = ArrCupos(posRamos, 1) 
											Inscritos = ArrCupos(posRamos, 2) 
										  end if
								 End Select
								%>
                                <%
							strParame="SELECT dbo.Fn_ValorParame('VALIDAHORARIOTRPORTAL')Parame"
							if BCL_ADO(strParame, rstParame) then
								ValidaHorario=rstParame("Parame")
							else
								ValidaHorario=""
							end if 
							  
							if ValidaHorario="SI" then  
							
								strHorari="select top 1 1 from dbo.RA_HORARI WHERE codramo='"&  ArrayRamos(0, posRamos) &"' AND ano='"& Ano &"' AND periodo='"& Periodo &"' AND CODSECC="& ArrayRamos(1, posRamos) &""
								
								if BCL_ADO(strHorari, rstHorari) then
									TieneHorario="SI"
								else
									TieneHorario="NO"
								end if 
							else
								TieneHorario="SI" 
							end if 
					
					  if TieneHorario="SI" then%>
							   <tr bgcolor="#DBECF2"> 
												<td width="97" height="20" align="center">
													   <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="detalle-asignatura.asp?R=<%=ArrayRamos(0, posRamos)%>&S=<%=ArrayRamos(1, posRamos)%> " target="_blank"><%=ArrayRamos(0, posRamos)%></a></font></div>								 </td>
												<td width="203" height="20" align="center"> 
													   <div align="center"><font face="Verdana" size="1"><%=ArrayRamos(5, posRamos) + " " + ArrayRamos(6, posRamos)%></font></div>								 </td>
												<td height="20" align="center">
													   <div align="center"><font face="Verdana" size="1"><%=ArrayRamos(1, posRamos)%></font></div>								 
													   <div align="center"></div></td>
												<td width="130" height="20" align="center">
													   <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"></font><font face="Verdana" size="0"><%=ArrayRamos(2, posRamos)%></font></div>								 </td>
												<td width="32" height="20" align="center"> <div align="center"><font face="Verdana" size="1"> 
														<% 
																If Inscritos >= Cupos then
																			if EstaPreInscrito(Codcli, RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, TipoCurso) then
																				PuedeBotar = True 					
																			else
																				PuedeBotar = False 					
																			end if
																	Response.Write("No")
																else
																		  Response.Write("Si")
																		  PuedeBotar = True 					
																end if
															  %></font></div>								 </td>
												<td width="41" height="20" > <div align="center"><font face="Verdana" size="1"> 
																	<%
																			 EstadoReal = EstadoRamoReal(CodCli, ArrayRamos(0, posRamos),"T")
																			If (EstadoReal = "A") or (EstadoReal = "E") or (EstadoReal = "I") then
																			   Cursado = True
																			 else
																			   Cursado = False
																			 end if 
																	  %>
																<%  
																	 if EstaPreInscrito(Codcli, RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, TipoCurso) then 
																	 %>
																					<input type="radio" name="<%=RamosSelec(iP)&"P"%>" value="<%=ArrayRamos(0, posRamos)%>" OnFocus="javascript:marca(this, '<%=trim(ArrayRamos(0, posRamos))%>')" OnClick="javascript:GrabaSel(this,'<%=RamosSelec(i)%>', '<%=ArrayRamos(0, posRamos)%>', '<%=ArrayRamos(1, posRamos)%>',<%=i%>,'<%=mid(TextoTipoCurso ,1,1)%>','<%=ArrayRamos(9, posRamos)%>')" CHECKED <%If ((Inscritos >= Cupos) or (cursado)) and Not PuedeBotar then Response.Write("DISABLED") End if%> >
																<% else %>
																					<input type="radio" name="<%=RamosSelec(iP)&"P"%>" value="<%=ArrayRamos(0, posRamos)%>" OnFocus="javascript:marca(this, '<%=trim(ArrayRamos(0, posRamos))%>')" OnClick="javascript:GrabaSel(this, '<%=RamosSelec(i)%>', '<%=ArrayRamos(0, posRamos)%>', '<%=ArrayRamos(1, posRamos)%>',<%=i%>,'<%=mid(TextoTipoCurso ,1,1)%>','<%=ArrayRamos(9, posRamos)%>')" <%If  (Inscritos >= Cupos) or Cursado then Response.Write("DISABLED") End if%> >
																<% end if %></font> </div>								 </td>
											<td width="75" height="20" >
											 <div align="center"><font face="Verdana" size="1"><%=TextoTipoCurso %></font></div>								 </td>
											
											<td width="94" height="20" >
													 <p><font face="Verdana" size="1" color="#000000" ><center><%=ArrayRamos(9, posRamos)%></center></font><font face="Verdana" size="1"></font></p>								 </td>
					  </tr>
                      <%else  'sin horario !!%>
                      <tr bgcolor="#99D7CD" > 
						<td width="80" height="20" align="center" bgcolor="#99D7CD"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="detalle-asignatura.asp?R=<%=ArrayRamos(0, posRamos)%>&S=<%=ArrayRamos(1, posRamos)%>&P=<%=ArrayRamos(5, posRamos) + " " + ArrayRamos(6, posRamos)%> " target="_blank"><%=ArrayRamos(0, posRamos)%></a></font></div></td>
						<td width="181" height="16" align="center"> <div align="center"><font face="Verdana" size="1"><%=ArrayRamos(5, posRamos) + " " + ArrayRamos(6, posRamos)%></font></div></td>
						<td height="16" align="center" bgcolor="#99D7CD"> <div align="center"><font face="Verdana" size="1"><%=ArrayRamos(1, posRamos)%></font></div> <div align="center"></div></td>
						<td width="141" height="16" align="center"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"></font><font face="Verdana" size="0"><%=ArrayRamos(2, posRamos)%></font></div></td>
						<td width="34" height="16" align="center"> <div align="center"><font face="Verdana" size="1"> 
							<% 
																						If Inscritos >= Cupos then
																									if EstaPreInscrito(Codcli, RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, TipoCurso) then
																										PuedeBotar = True 					
																									else
																										PuedeBotar = False 					
																									end if
																							Response.Write("No")
																						else
																								  Response.Write("Si")
																								  PuedeBotar = True 					
																						end if
																					  %>
						</font></div></td>
						<td width="72" height="16"> <div align="center">
							<p><font face="Verdana" size="1"> 
							  <%
																									 EstadoReal = EstadoRamoReal(CodCli, ArrayRamos(0, posRamos),"T")
																									If (EstadoReal = "A") or (EstadoReal = "E") or (EstadoReal = "I") then
																									   Cursado = True
																									 else
																									   Cursado = False
																									 end if 
																							  %>
							  
									<input type="radio" name="<%=RamosSelec(iP)%>" value="<%=ArrayRamos(0, posRamos)%>" DISABLED >
							 
							  </font> </p>
						</div></td>
						<td width="101" ><center> <p><font face="Verdana" size="1"><%=TextoTipoCurso%></font><font face="Verdana" size="1"></font></p></center></td>
                        <td width="94" height="20" >
													 <p><font face="Verdana" size="1" color="#000000" ><center><%=ArrayRamos(9, posRamos)%></center></font><font face="Verdana" size="1"></font></p>								 </td>
					  </tr>
                      <%end if 'fin sin horario%>
                      
                      
						  <%
						 posRamos = posRamos + 1
						if posramos <=   Ubound(ArrayRamos, 2) then
							TipoCurso = ArrayRamos(8, posRamos)
						end if
						loop
				  %>
				</table>
				</td>
			  </tr>
			</table>
			</form>
				   <%    
			 END IF
			 
			 if  TipoCurso="L"  then	 	
			   %>
			<form  id=form1<%=RamosSelec(i)%>  name=<%=RamosSelec(i)%>>
				<table width="750" border="0" cellspacing="15" cellpadding="0">
					<tr valign="top"> 
			  <td width="707" height="83" colspan="2"> 
				  <table width="750" cellspacing="0" cellpadding="0" height="106" border="1" bordercolor="#FFFFFF">
					  <tr bgcolor="83a3d0"> 
								<td height="22" colspan="8"> <div align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2" name="lblAsig3">Asignatura: </font></b></font><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2">
									<% if TipoRamo<>"" then
											response.Write(RamosSelec(i) + " " + GetNombreRamo(RamosSelec(i)))
										else
											response.Write(RamosSelec(i) & " " & replace(GetNombreRamo(RamosSelec(i)),"FORMACION GENERAL","CURSO OPTATIVO ADICIONAL"))
										end if%></font></b></font></div>								</td>
					  </tr>
					  <tr background="Imagenes/fdo-cabecera-cel22.jpg"> 
								<td width="97" height="30" background="Imagenes/fdo-cabecera-cel22.jpg" >
									   <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Lista Curso</font></b></font></div>						</td>
								<td width="203" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
									   <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Profesor</font></b></font></div>						</td>
								<td width="60" height="30" background="Imagenes/fdo-cabecera-cel22.jpg">
										<div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">S</font><font size="1">ecci&oacute;n</font></b></font></div>
										<div align="center"></div></td>
								<td width="130" height="30" background="Imagenes/fdo-cabecera-cel22.jpg">
										 <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Ramo</font></b></font></div>						</td>
								<td width="32" height="30" background="Imagenes/fdo-cabecera-cel22.jpg">
										 <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1"> 
											Cupo</font></b></font></div>						</td>
								<td width="41" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"> 
										  <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Inscribir</font></b></font></div>						</td>
								<td width="75" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1">Tipo 
										  Asignatura </font></b></font>					</td>
								
						<td width="94" height="30" background="Imagenes/fdo-cabecera-cel22.jpg"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif" >
						  <center>
							<b>Corresponde Seccion Teorica </b>
			</center></font> </td>
					  </tr>
			  <% 	x=0
					 do while posRamos <= Ubound(ArrayRamos, 2) and TipoCurso="L" 'AND X=0
					 TextoTipoCurso ="Laboratorio"
					 session("OtrosRamos") = "N"						                   
					 TipoCupo = "CT"
					 Select Case TipoRamo
						 Case ""
						  if session("TipoRamoOtros") = "OBLIGATORIO" then
							  session("OtrosRamos") = "S"						  
						  end if
						  Cupos = SacaCupoAntiguo (ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, CodCarr, CodSede, EsRepit, TipoCupo,RamosSelec(i))
						  if session("RamoObliPlanificadoOp") = "N" then
							  Inscritos = CantidadInscritos ( ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), CodSede, Ano, Periodo, TipoCupo, TipoCurso )
						  else
							  Inscritos = CantidadInscritosOptativos(RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), CodSede, Ano, Periodo, CodCarr, TipoCupo, TipoCurso) 							  	  	
						  end if	
								Case "OBLIGATORIO"
										 Cupos = SacaCupoAntiguo (ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, CodCarr, CodSede, EsRepit, TipoCupo,RamosSelec(i))
										  if session("RamoObliPlanificadoOp") = "N" then
											  Inscritos = CantidadInscritos ( ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), CodSede, Ano, Periodo, TipoCupo, TipoCurso )
										  else
											  Inscritos = CantidadInscritosOptativos(RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), CodSede, Ano, Periodo, CodCarr, TipoCupo, TipoCurso)
										  end if	
							   Case "OPCIONAL"
										  Cupos = SacaCupoOptativo(RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, CodCarr,CodSede) 
										  Inscritos = CantidadInscritosOptativos(RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), CodSede, Ano, Periodo, CodCarr, TipoCupo, TipoCurso)
							   Case "OPTATIVO ESPECIAL"
										  Inscritos=0
										  if revisaCupos then
											Cupos = SacaCupoOptativoEspecial(ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, CodCarr,RamosSelec(i)) 
											Inscritos = CantidadInscritosOptativoEspecial(ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), CodSede, Ano, Periodo, CodCarr, TipoCupo, TipoCurso)
											ArrCupos(posRamos, 1) = Cupos
											ArrCupos(posRamos, 2) = Inscritos                          
										  else
											Cupos = ArrCupos(posRamos, 1) 
											Inscritos = ArrCupos(posRamos, 2) 
										  end if
								 End Select
								%>
                                <%strParame="SELECT dbo.Fn_ValorParame('VALIDAHORARIOTRPORTAL')Parame"
							if BCL_ADO(strParame, rstParame) then
								ValidaHorario=rstParame("Parame")
							else
								ValidaHorario=""
							end if 
							  
							if ValidaHorario="SI" then  
							
								strHorari="select top 1 1 from dbo.RA_HORARI WHERE codramo='"&  ArrayRamos(0, posRamos) &"' AND ano='"& Ano &"' AND periodo='"& Periodo &"' AND CODSECC="& ArrayRamos(1, posRamos) &""
								
								if BCL_ADO(strHorari, rstHorari) then
									TieneHorario="SI"
								else
									TieneHorario="NO"
								end if 
							else
								TieneHorario="SI" 
							end if 
					 
					  if TieneHorario="SI" then%>
							   <tr bgcolor="#DBECF2"> 
												<td width="97" height="20" align="center">
													   <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="detalle-asignatura.asp?R=<%=ArrayRamos(0, posRamos)%>&S=<%=ArrayRamos(1, posRamos)%> " target="_blank"><%=ArrayRamos(0, posRamos)%></a></font></div>								 </td>
												<td width="203" height="20" align="center"> 
													   <div align="center"><font face="Verdana" size="1"><%=ArrayRamos(5, posRamos) + " " + ArrayRamos(6, posRamos)%></font></div>								 </td>
												<td height="20" align="center">
													   <div align="center"><font face="Verdana" size="1"><%=ArrayRamos(1, posRamos)%></font></div>								 
													   <div align="center"></div></td>
												<td width="130" height="20" align="center">
													   <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"></font><font face="Verdana" size="0"><%=ArrayRamos(2, posRamos)%></font></div>								 </td>
												<td width="32" height="20" align="center"> <div align="center"><font face="Verdana" size="1"> 
														<% 
																If Inscritos >= Cupos then
																			if EstaPreInscrito(Codcli, RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, TipoCurso) then
																				PuedeBotar = True 					
																			else
																				PuedeBotar = False 					
																			end if
																	Response.Write("No")
																else
																		  Response.Write("Si")
																		  PuedeBotar = True 					
																end if
															  %></font></div>								 </td>
												<td width="41" height="20" > <div align="center"><font face="Verdana" size="1"> 
																	<%
																			 EstadoReal = EstadoRamoReal(CodCli, ArrayRamos(0, posRamos),"T")
																			If (EstadoReal = "A") or (EstadoReal = "E") or (EstadoReal = "I") then
																			   Cursado = True 
																			 else
																			   Cursado = False
																			 end if 
																	  %>
																<%  
																
																	 if EstaPreInscrito(Codcli, RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, TipoCurso) then 
																	 %>
																					<input type="radio" name="<%=RamosSelec(iP)&"L"%>" value="<%=ArrayRamos(0, posRamos)%>" OnFocus="javascript:marca(this, '<%=trim(ArrayRamos(0, posRamos))%>')" OnClick="javascript:GrabaSel(this,'<%=RamosSelec(i)%>', '<%=ArrayRamos(0, posRamos)%>', '<%=ArrayRamos(1, posRamos)%>',<%=i%>,'<%=mid(TextoTipoCurso ,1,1)%>','<%=ArrayRamos(9, posRamos)%>')" CHECKED <%If ((Inscritos >= Cupos) or (cursado)) and Not PuedeBotar then Response.Write("DISABLED") End if%> >
																<% else %>
																					<input type="radio" name="<%=RamosSelec(iP)&"L"%>" value="<%=ArrayRamos(0, posRamos)%>" OnFocus="javascript:marca(this, '<%=trim(ArrayRamos(0, posRamos))%>')" OnClick="javascript:GrabaSel(this, '<%=RamosSelec(i)%>', '<%=ArrayRamos(0, posRamos)%>', '<%=ArrayRamos(1, posRamos)%>',<%=i%>,'<%=mid(TextoTipoCurso ,1,1)%>','<%=ArrayRamos(9, posRamos)%>')" <%If  (Inscritos >= Cupos) or Cursado then Response.Write("DISABLED") End if%> >
																<% end if %></font> </div>								 </td>
											<td width="75" height="20" >
											 <div align="center"><font face="Verdana" size="1"><%=TextoTipoCurso %></font></div>								 </td>
											
											<td width="94" height="20" >
													 <p><font face="Verdana" size="1" color="#000000" ><center><%=ArrayRamos(9, posRamos)%></center></font><font face="Verdana" size="1"></font></p>								 </td>
					  </tr>
                      <%else  'sin horario !!%>
                      <tr bgcolor="#99D7CD" > 
						<td width="80" height="20" align="center" bgcolor="#99D7CD"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="detalle-asignatura.asp?R=<%=ArrayRamos(0, posRamos)%>&S=<%=ArrayRamos(1, posRamos)%>&P=<%=ArrayRamos(5, posRamos) + " " + ArrayRamos(6, posRamos)%> " target="_blank"><%=ArrayRamos(0, posRamos)%></a></font></div></td>
						<td width="181" height="16" align="center"> <div align="center"><font face="Verdana" size="1"><%=ArrayRamos(5, posRamos) + " " + ArrayRamos(6, posRamos)%></font></div></td>
						<td height="16" align="center" bgcolor="#99D7CD"> <div align="center"><font face="Verdana" size="1"><%=ArrayRamos(1, posRamos)%></font></div> <div align="center"></div></td>
						<td width="141" height="16" align="center"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"></font><font face="Verdana" size="0"><%=ArrayRamos(2, posRamos)%></font></div></td>
						<td width="34" height="16" align="center"> <div align="center"><font face="Verdana" size="1"> 
							<% 
																						If Inscritos >= Cupos then
																									if EstaPreInscrito(Codcli, RamosSelec(i), ArrayRamos(0, posRamos), ArrayRamos(1, posRamos), Ano, Periodo, TipoCurso) then
																										PuedeBotar = True 					
																									else
																										PuedeBotar = False 					
																									end if
																							Response.Write("No")
																						else
																								  Response.Write("Si")
																								  PuedeBotar = True 					
																						end if
																					  %>
						</font></div></td>
						<td width="72" height="16"> <div align="center">
							<p><font face="Verdana" size="1"> 
							  <%
																									 EstadoReal = EstadoRamoReal(CodCli, ArrayRamos(0, posRamos),"T")
																									If (EstadoReal = "A") or (EstadoReal = "E") or (EstadoReal = "I") then
																									   Cursado = True
																									 else
																									   Cursado = False
																									 end if 
																							  %>
							  
									<input type="radio" name="<%=RamosSelec(iP)%>" value="<%=ArrayRamos(0, posRamos)%>" DISABLED >
							 
							  </font> </p>
						</div></td>
						<td width="101" ><center> <p><font face="Verdana" size="1"><%=TextoTipoCurso%></font><font face="Verdana" size="1"></font></p></center></td>
                        <td width="94" height="20" >
													 <p><font face="Verdana" size="1" color="#000000" ><center><%=ArrayRamos(9, posRamos)%></center></font><font face="Verdana" size="1"></font></p>								 </td>
					  </tr>
                      <%end if 'fin sin horario%>
                      
                      
						  <%
						 posRamos = posRamos + 1
						if posramos <=   Ubound(ArrayRamos, 2) then
							TipoCurso = ArrayRamos(8, posRamos)
						end if
						loop
				  %>
				</table>
				</td>
			  </tr>
			  </table>
			</form>
   <%   end if
   	END IF
							%>
<%
  next 
%>
			</td>
		  </tr>
		</table>
	</td>
  </tr>
</table>

</body>
<%ObjetosLocalizacion("asignatura-seccion.asp")%>
</html>

<script>
function Actualizar() {
  document.TheForm.Accion.value="Seleccionar";
  document.TheForm.submit();
}
function Enviar() {
     document.TheForm.Accion.value="Enviar";
     document.TheForm.submit();
}
</script>



<!--#INCLUDE file="include/desconexion.inc" -->
