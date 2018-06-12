<%
  Response.Buffer = false
  Response.Expires = -1 
  ' Un gran parametro que indica si funciona como solicitud
  SoloPrueba = REQUEST("P")
  If Trim(SoloPrueba = "") then
     SoloPrueba = "N"
  end if
%>
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#INCLUDE FILE="include/cupos.inc" -->
<!--#INCLUDE FILE="include/nombreramo.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<script language="JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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
</script>
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=profe name=profe></OBJECT> 
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=seccion name=seccion></OBJECT> 
<OBJECT RUNAT=server PROGID=ADODB.Recordset id=RS name=RS></OBJECT> 

<html> <head> 
<title>Portal de Alumnos en L&iacute;nea</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="css/tex-normales.css" type="text/css">
<link rel="stylesheet" href="css/tablas.css" type="text/css">
<%
   Dim RamosSelec
   dim var
   var=request("P")
   CodSede = Session("CodSede")
   Ano = Session("Ano") & ""
   Periodo = Session("Periodo") & ""
   CodCarr = Session("CodCarr")
   CodCarr2 = Request("CodCarr2")
   CodPestud = Session("CodPestud")
   Codcli = Session("Codcli") 
   Jornada = Request("J")
  ' Response.Write("Es " &  NivelChosen)
   If trim("CodCli") = "" then
     Response.Write("Debe comenzar de nuevo")
     Response.end
   end if
   ' Debemos separar los ramos
  
  'Dim PrioridadRamos(30)
  Dim RamosSeccion
  Dim PrimeraVez, LastRstSeccion
  Dim ArrVacio(7,0)
  Dim s_RamoFueraMallaRep  
  
  PrimeraVez = true
  ' Aqui guardar en arreglo con el fin de no volver a buscar
  Dim ArrCupos(300, 5)
  
  StrSql = "select ObligaRamosFueraMalla from mt_parame"
  StrSql = StrSql & " where ObligaRamosFueraMalla='S'"
  
  if BCL_ADO(StrSql, rst) then
  	ObligaRamosFueraMallaRep ="S"
	rst.close()
  else
	ObligaRamosFueraMallaRep = "N"
  end if
  	  
  StrSql = "select Max(convert(numeric, nivel )) as Max, Min(convert(numeric, nivel )) as Min "
  StrSql = StrSql & " from ra_curric c "
  StrSql = StrSql & " Where CodPestud = '" & CodPestud & "' "
  StrSql = StrSql & " and not (exists (select codramo from ra_nota where Codcli = '" & Codcli & "' "
  StrSql = StrSql & " and codramo = c.codramo and estado in ('A', 'E', 'I', ''))) "
  StrSql = StrSql & " and not exists ((select codramo from ra_carga where Codcli = '" & Codcli & "' "
  StrSql = StrSql & " and codramo = c.codramo and inscrito = 'S')) "
  
  Dim RstMinMax
'  Response.Write(StrSql)
 ' response.End
  if BCL_ADO(StrSql, RstMinMax) then
    Min = RstMinMax("Min")
    Max = RstMinMax("Max")
    RstMinMax.close()
  else
    Min = 1
    Max = 12
  end if

   if Trim(Request("Niv")) <> "" then
     NivelChosen = Request("Niv")
   else
      NivelChosen = Min
      'Response.Write("Es cuatro")
        ' Analizar  el nivel que conviene mostrar
   end if
  
%>

<script languaje="javascript">
function visualiza() {
  //alert(P)
  var x=window.open ("tex-inscripcionespecial.htm","Solicitud_Especial","width=400,height=600,resizable=yes,toolbar=yes");
 }

function marca(Obj, valor)
{
   if (Obj.checked) {Obj.value = "false"} 
   else {Obj.value = valor}
  //Obj.checked = !Obj.checked;
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
            //alert(Ramo + " : " + document.forms[j].elements[k].value);
            if ((Ramo == document.forms[j].elements[k].value) && (document.forms[j].elements[k].checked) )
            { alert("Ramo : " + document.forms[j].elements[k].value + " inscrito con " + document.forms[j].name ); 
              Obj.checked = false;
              return false;}
          }
        }         
     }
  }
  return true;
}


function GrabaTexto(Obj)
{
   var k, Ramo, Seccion, Curso, Texto;
   
   for (k = 0; k < Obj.form.elements.length; k++)
   {
      if (Obj.form.elements[k].type == "radio")
      {
         if (Obj.form.elements[k].checked)
         { 
            Ramo = Obj.form.elements[k].value;
            Curso = Obj.form.name;
            Seccion = "0";
            //alert("Curso = " + Curso);
            parent.Horario.location = "GrabaSol-Especial-Nivel.asp?A=I&R=" + Ramo + "&S=" + Seccion + "&T=" + Obj.value + "&C=" + Curso;            
         }
      } 
   }
}

function GrabaSel(Obj, Curso, Ramo, Seccion, CodCarr)
{ var j, k;

   if ((Obj.checked) && (Obj.value == "false") )
   {
      Obj.checked = false;
      //RamoInscrito[i] = false;
      <% if SoloPrueba = "N" then %>
      parent.Horario.location = "GrabaSol-Especial-Nivel.asp?A=D&R=" + Ramo + "&S=" + Seccion + "&C=" + Curso + "&CodCarr=" + CodCarr
      <% else %>
     parent.Horario.location = "PruebaRamo.asp?A=D&R=" + Ramo + "&S=" + Seccion;
      <% end if %>
   }   
   else {
   if ((Obj.checked) && (Obj.value != "false") )
     {
      <% if SoloPrueba = "N" then %>
       parent.Horario.location = "GrabaSol-Especial-Nivel.asp?A=I&R=" + Ramo + "&S=" + Seccion + "&T=" + Obj.form.elements["T" + Curso].value + "&C=" + Curso + "&CodCarr=" + CodCarr
      <% else %>
     parent.Horario.location = "PruebaRamo.asp?A=I&R=" + Ramo + "&S=" + Seccion;
      <% end if %>
     }
   else
     {
      <% if SoloPrueba = "N" then %>
       parent.Horario.location = "GrabaSol-Especial-Nivel.asp?A=I&R=" + Ramo + "&S=" + Seccion + "&T=" + Obj.form.elements["T" + Curso].value + "&C=" + Curso + "&CodCarr=" + CodCarr
      <% else %>
     parent.Horario.location = "PruebaRamo.asp?A=I&R=" + Ramo + "&S=" + Seccion;
      <% end if %>
     }
   }
}

function EscondeHor()
{
   //alert("paso");
   //parent.location = "asignatura-seccion.asp?Ramos="
   parent.Seleccion.resizeBy(1000, 1000);
}

function ConfirmaSoli()
{var j, k, unoalmenos;
<%
  'Session("InscripcionDebeRealizada") = "S"
  if SoloPrueba = "N" then
%>
  // veremos si al menos hay uno inscrito
  unoalmenos = false; 
  for (j=0; j< document.forms.length; j++)
  {
     //if (document.forms.elements[i].type == "option")
     for (k=0; k < document.forms[j].elements.length ;k++)
     {
          //alert(document.forms[j].name); 
          if (document.forms[j].elements[k].type == "radio")
          { 
             if (document.forms[j].elements[k].checked)
             { unoalmenos = true}
          }
     }
  }
  
  if (!unoalmenos) { alert("Debe inscribir al menos uno") }     
  if (unoalmenos) {     
    if (confirm("Ud solicitará las asignaturas de forma definitiva"))
    { 
       window.location =  "confirmacion-solici-especial-nivel.asp" 
    }
 } 
<%
 else
%>
  window.location =  "inscrip-asigna.asp";
<% 
 end if
%>  
}

function Nada()
{
}

function llamanivel(nivel)
{
    window.location = "solicitud-seccion-especial-nivel.asp?P=<%=var%>&Niv=" + nivel;
}
function RefrescarJornada(Jornada)
{
    window.location = "solicitud-seccion-especial-nivel.asp?P=<%=var%>&Niv=" + nivel + "&J=" + Jornada;
}

function Actualizar() {
  document.TheForm.submit();
}

</script>


<body  bgcolor="#FFFFFF" text="#000000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" > 
<% 'if (not InscritoNormal()) and (SoloPrueba = "N") then %>
  <script>
    // alert("No ha inscrito aún vía solicitud Normal ...");
  </script>   
<%  'end if%>
<table border="0" cellpadding="0" cellspacing="0"  align="left">
  <tr>
    <td><table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
      <tr>
         <td height="30" align="left" valign="top"><% CargarTop1()%><% SubMenu()%>
		    <table width="750" border="0" cellspacing="15" cellpadding="0" height="43">
                <tr>
                  <%if var="S" then %>
                  <td width="336"><img name="fder_r2_c1" src="Imagenes/titulos/T-otras_asignaturas.gif" width="400" height="38" border="0"></td>
                  <%else%>
                  <td width="336"><img name="fder_r2_c1" src="Imagenes/titulos/T-solic_espec_inscrip.gif" width="471" height="38" border="0"></td>
                  <%end if%>
                  <%if var="S" then %>
                  <td width="248">&nbsp;</td>
                  <%else%>
                  <td width="248">&nbsp;</td>
                  <%end if%>
                </tr>
                <tr>
                  <%if var="S" then %>
                  <td width="336" valign="bottom" height="34" class="tex-normales"><font color="#CCCCCC">&nbsp;</font><a href="javascript:;" ></a> Asignaturas Jornada
                    <select name="SelJornada" onclick="javascript:RefrescarJornada(this.value);" onChange="javascript:Actualizar();">
                        <option value="D" selected>Diurna</option>
                        <option value="V">Vespertina</option>
                    </select>
                      <br>
                      <%else%>
                  <td width="336" valign="bottom" height="34" class="tex-normales"><%end if%>
                      <%Response.write("Seleccione Nivel") 
        for i = ValNulo((Min), num_) to ValNulo((Max), num_) %>
                      <a href="javascript:llamanivel(<%=i%>)"><%=i%></a> &nbsp;
                    <% next %>                  </td>
                  <td width="248" height="34"><%if var="S" then %>
                      <div align="right"><a href="javascript:ConfirmaSoli()" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','Imagenes/botones/A-_r12_c2_f2.jpg',1)"><img src="Imagenes/botones/A-_r12_c2.jpg" width="162" height="21" border="0" name="Image1"></a></div>
                    <% else %>
                      <div align="right"><a href="javascript:ConfirmaSoli()" onMouseOut="MM_swapImgRestore()" onMouseOver="MM_swapImage('Image1','','Imagenes/botones/confirmar-solic-espec-on.gif',1)"><img src="Imagenes/botones/confirmar-solic-espec-of.gif" width="246" height="45" border="0" name="Image1"></a>					  </div>
                    <% end if %>                  </td>
                  <td height="34">&nbsp;</td>
                </tr>
                <%

   StrSql = "Select count (Distinct s.CodCarr) "
   StrSql = StrSql & " From ra_optivo o, ra_seccio s, mt_carrer c "
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
   StrSql = StrSql & " And o.tipo = 'OPTATIVO ESPECIAL'"
   StrSql = StrSql & " and s.tipocurso='T'" 
   StrSql = StrSql & " And s.Jornada= '" & Jornada & "' "
   strsql = strsql & " and o.codramo='xxx'"
   if BCL_ADO(StrSql, rst2) then
     CantidadRamos = rst2(0)
     rst2.close()
   else
     CantidadRamos = 0
   end if

   ' Analizaremos si es una carrera de Bachillerato
   StrSql = "Select Distinct s.CodCarr, c.nombre_l "
   StrSql = StrSql & " From ra_optivo o, ra_seccio s, mt_carrer c "
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
   StrSql = StrSql & " And s.Jornada= '" & Jornada & "' "
   StrSql = StrSql & " And o.tipo = 'OPTATIVO ESPECIAL'"
   StrSql = StrSql & " and s.tipocurso='T'" 
   strsql = strsql & " and o.codramo='xxx'"

   'Response.Write(StrSql) 
%>
                <%

   if BCL_ADO(Strsql, Rst2) then
    if CantidadRamos > 1 then
     If trim(CodCarr2) = "" then
        CodCarr2 = Rst2("CodCarr")
     end if
%>
                <tr>
                  <td class="tex-normales" valign="middle" height="23" align="right"><div align="right"></div></td>
                  <td colspan=2 height="23" align="left" valign="middle"><form name=frmCarrera>
                      <select name="CodCarr2" onChange="javascript:submit()">
                        <% Do while not Rst2.eof%>
                        <%   If trim(CodCarr2) = Rst2("CodCarr") then %>
                        <option value=<%=Rst2("CodCarr")%>><%=Rst2("Nombre_l")%></option>
                        <%   else %>
                        <option value=<%=Rst2("CodCarr")%> selected><%=Rst2("Nombre_l")%></option>
                        <%   end if 
             Rst2.movenext
           loop 
           Rst2.close()  
         %>
                      </select>
                  </form></td>
                </tr>
                <%          
   end if
  end if 
%>
              </table>
              
          <%

'RS.Open  " select * from ra_OtrosRamos where codcli = '" + Codcli + "' ", Session("Conn")
RS.Open  " OTROSRAMOS '" + Codcli + "' , " & NivelChosen, Session("Conn")

response.Write("<table width='750'border='0' cellspacing='15' cellpadding='0'>")

if not RS.EOF then
  ArregloOtrosRamos = RS.GetRows
  Matching = true
else
  Matching = false
end if

'Do while not RS.EOF
if Matching then
  posOtros = 0
  do while posOtros <= Ubound(ArregloOtrosRamos, 2)
 '  for i = 0 to ubound(RamosSelec, 1) - 1
    if (clng(ArregloOtrosRamos( 2, posOtros)) = clng(NivelChosen)) then
      'PrioridadRamos(i) = PrioridadRamo(Codcli, ArregloOtrosRamos(0, posOtros), Ano, Periodo)
%>
              <% ' Recorremos las secciones 
                 'Response.Write("Ramo = " & RamosSelec(i) & " Tipo = " & GetTipoRamo(RamosSelec(i), CodPestud) & " CodPestud = " & CodPestud )
                 'Response.End
                 EsRepit = EsRepitente(ArregloOtrosRamos(0, posOtros), CodCli) 
                 TipoRamo = GetTipoRamo(ArregloOtrosRamos(0, posOtros), CodPestud)      
                 session("OtrosRamos") = "N"
                 'Response.Write("Tipo Ramo = " & TipoRamo)
                 'Response.End 
				 
				 s_RamoFueraMallaRep = ""
				 if ObligaRamosFueraMallaRep="S" then
				 	s_RamoFueraMallaRep = GetRamoFueraMallaRep(ArregloOtrosRamos(0, posOtros), CodCli,Ano,Periodo)
				 end if
				 
                 select Case TipoRamo
                   case "OBLIGATORIO"
                     if EsPlanificado(ArregloOtrosRamos(0, posOtros), CodSede, Ano, Periodo,codcarr) then
                       StrSeccion = "Select e.CodRamo, e.CodSecc, r.nombre, 0 as cupo, e.CodProf, p.Ap_Pater, p.Nombres, e.horario "
                       StrSeccion = StrSeccion & " from ra_seccio e, ra_profes p, ra_ramo r "
                       StrSeccion = StrSeccion & " Where e.CodRamo = '" & ArregloOtrosRamos(0, posOtros) & "' and e.CodSede = '" & CodSede & "' " 
                       StrSeccion = StrSeccion & " And e.Ano = " & ano & " And e.Periodo = " & Periodo
                       StrSeccion = StrSeccion & " and e.CodProf *= p.CodProf "
                       StrSeccion = StrSeccion & " and e.codramo = r.codramo" 
					   StrSeccion = StrSeccion & " and e.tipocurso='T'" 
					   StrSeccion = StrSeccion & " and e.CodCarr='" & trim(Codcarr) & "'" 
                       StrSeccion = StrSeccion & " Order by convert(numeric, CodSecc) "
                     else
                       Equivalente = GetEquivalente(ArregloOtrosRamos(0, posOtros), CodSede, Ano, Periodo, CodCarr)
                       if Trim(Equivalente) <> "" then
                          StrSeccion = "Select e.CodRamo, e.CodSecc, r.nombre, 0 as cupo, e.CodProf, p.Ap_Pater, p.Nombres, e.horario "
                          'StrSeccion = "Select CodRamo, CodSecc, '', '', CodProf 
                          StrSeccion = StrSeccion & " from ra_seccio e, ra_profes p, ra_ramo r "
                          StrSeccion = StrSeccion & " Where e.CodRamo = '" & Equivalente & "' and e.CodSede = '" & CodSede & "' " 
                          StrSeccion = StrSeccion & " And e.Ano = " & ano & " And e.Periodo = " & Periodo
                          StrSeccion = StrSeccion & " and e.CodProf *= p.CodProf "
                          StrSeccion = StrSeccion & " and e.codramo = r.codramo" 
                          StrSeccion = StrSeccion & " and e.tipocurso='T'" 
                          StrSeccion = StrSeccion & " Order by convert(numeric, CodSecc) "
                       else
						   StrSeccion = "Select distinct a.codramo,a.codsecc, b.nombre, a.cupo, e.CodProf, p.Ap_Pater, p.nombres, e.horario  " 
						   StrSeccion = StrSeccion & " from ra_optivo a,ra_ramo b,ra_seccio e, ra_profes p" 
						   StrSeccion = StrSeccion & " where a.codcarr  = '" & CodCarr & "' " 
						   StrSeccion = StrSeccion & " and a.codramo=b.codramo" 
						   StrSeccion = StrSeccion & " and a.ano=" & Ano & "" 
						   StrSeccion = StrSeccion & " and a.periodo=" & Periodo & "" 
						   StrSeccion = StrSeccion & " and a.codramo=e.codramo" 
						   StrSeccion = StrSeccion & " and a.codsecc=e.codsecc" 
						   StrSeccion = StrSeccion & " and a.ano=e.ano" 
						   StrSeccion = StrSeccion & " and a.periodo=e.periodo" 
						   StrSeccion = StrSeccion & " and e.CodSede = '" & CodSede & "' "
						   StrSeccion = StrSeccion & " and a.codopt = '" & ArregloOtrosRamos(0, posOtros) & "' "
   	   					   if trim(s_RamoFueraMallaRep)<>"" then
		                       StrSeccion = StrSeccion & " and a.codramo = '" & trim(s_RamoFueraMallaRep) & "' "
						   end if
						   StrSeccion = StrSeccion & " and e.CodProf *= p.CodProf "
					       StrSeccion = StrSeccion & " and e.tipocurso='T'" 
     					   StrSeccion = StrSeccion & " And a.tipo <> 'OPTATIVO ESPECIAL'"
							if ArregloOtrosRamos(0, posOtros)="EDUC0012" then
				'				response.write(StrSeccion)
							end if
						   EsEspecial = false
                       end if
                     end if 
                     EsEspecial = false
                   case "OPCIONAL"
                       StrSeccion = "Select distinct a.codramo,a.codsecc, b.nombre, a.cupo, e.CodProf, p.Ap_Pater, p.nombres, e.horario  " 
                       StrSeccion = StrSeccion & " from ra_optivo a,ra_ramo b,ra_seccio e, ra_profes p" 
                       StrSeccion = StrSeccion & " where a.codcarr  = '" & CodCarr & "' " 
                       StrSeccion = StrSeccion & " and a.codramo=b.codramo" 
                       StrSeccion = StrSeccion & " and a.ano=" & Ano & "" 
                       StrSeccion = StrSeccion & " and a.periodo=" & Periodo & "" 
                       StrSeccion = StrSeccion & " and a.codramo=e.codramo" 
                       StrSeccion = StrSeccion & " and a.codsecc=e.codsecc" 
                       StrSeccion = StrSeccion & " and a.ano=e.ano" 
                       StrSeccion = StrSeccion & " and a.periodo=e.periodo" 
                       StrSeccion = StrSeccion & " and e.CodSede = '" & CodSede & "' "
                       StrSeccion = StrSeccion & " and a.codopt = '" & ArregloOtrosRamos(0, posOtros) & "' "
   					   if trim(s_RamoFueraMallaRep)<>"" then
	                       StrSeccion = StrSeccion & " and a.codramo = '" & trim(s_RamoFueraMallaRep) & "' "
					   end if
                       StrSeccion = StrSeccion & " and e.CodProf *= p.CodProf "
                       StrSeccion = StrSeccion & " and e.tipocurso='T'" 
					   StrSeccion = StrSeccion & " And a.tipo <> 'OPTATIVO ESPECIAL'"

                       EsEspecial = false
                       'Response.Write(StrSeccion)
                     
                   case "OPTATIVO ESPECIAL"  
                       StrSeccion = "Select distinct a.codramo,a.codsecc, b.nombre, a.cupo, e.CodProf, p.Ap_Pater, p.nombres, e.horario " 
                       StrSeccion = StrSeccion & " from ra_optivo a,ra_ramo b,ra_seccio e, ra_profes p" 
                       StrSeccion = StrSeccion & " where a.codcarr  = '" & CodCarr & "' " 
                       StrSeccion = StrSeccion & " and a.codramo=b.codramo" 
                       StrSeccion = StrSeccion & " and a.ano=" & Ano & "" 
                       StrSeccion = StrSeccion & " and a.periodo=" & Periodo & "" 
                       StrSeccion = StrSeccion & " and a.codramo=e.codramo" 
                       StrSeccion = StrSeccion & " and a.codsecc=e.codsecc" 
                       StrSeccion = StrSeccion & " and a.ano=e.ano" 
                       StrSeccion = StrSeccion & " and a.periodo=e.periodo" 
                       StrSeccion = StrSeccion & " and e.CodSede = '" & CodSede & "' "
                       StrSeccion = StrSeccion & " and a.CodSede = '" & CodSede & "' "
                       StrSeccion = StrSeccion & " and e.CodProf *= p.CodProf "
                       if trim(CodCarr2) <> "" then
                         StrSeccion = StrSeccion & " and e.CodCarr = '" & CodCarr2 & "' "
                       end if
					   if trim(s_RamoFueraMallaRep)<>"" then
	                       StrSeccion = StrSeccion & " and a.codramo = '" & trim(s_RamoFueraMallaRep) & "'"
					   end if
                       StrSeccion = StrSeccion & " and e.tipocurso='T'" 
     				   StrSeccion = StrSeccion & " and a.tipo = 'OPTATIVO ESPECIAL'"

                       EsEspecial = True
                       'Response.Write(StrSeccion)
                       'Response.end 
                 end select
                 
                 If EsEspecial then
                    if PrimeraVez then
                      'Set RstSeccion = Session("Conn").Execute (StrSeccion)
                      'RamosSeccion = RstSeccion.GetRows()
                      'Set LastRstSeccion = RstSeccion
                      PrimeraVez = false
                      RevisaCupos = true
                      'Response.Write("Leido")
                    else
                      RevisaCupos = false
                      'Set RstSeccion = Session("Conn").Execute (StrSeccion)
                      'Set RstSeccion = LastRstSeccion
                      'RstSeccion.movefirst
                      'Response.Write("de nuevo <br>" )
                    end if
                 else
                    'Set RstSeccion = Session("Conn").Execute (StrSeccion)
                    'if RstSeccion.eof then 
                    '   RamosSeccion = ArrVacio
                    'else
                    '   RamosSeccion = RstSeccion.GetRows()
                    'end if
                 end if
                'Response.Write(StrSeccion & "<br>") 
                'Response.Write("Antes Query <br>")
				
                if trim(StrSeccion) <> "" then
%>
              <!--#INCLUDE FILE="include/reconexion.inc" -->
             
			  <%
'				  if BCL_ADO(StrSeccion, RstSeccion) then
					
'                  end if	
					  	
                  'Set RstSeccion = Session("Conn").Execute (StrSeccion)
				  
                end if
	   				 
				if trim(StrSeccion) <> "" then
                  if BCL_ADO(StrSeccion, RstSeccion) then			
%>
              <form  id=<%=ArregloOtrosRamos(0, posOtros)%>  name=<%=ArregloOtrosRamos(0, posOtros)%> action="javascript:Nada()">
			  <tr valign="top"> 
			    <td><table width="750" cellspacing="1" cellpadding="0" height="19" border="0" bordercolor="#FFFFFF">
                  <tr>
                    <td height="18" colspan="7" background="Imagenes/fdo-cabecera-cel2.jpg"><div align="left" class="text-cabecera-celda"><font c ><b><font size="2">Asignatura</font></b></font><font c ><b><font size="2">:
                      <% =ArregloOtrosRamos(0, posOtros) + " " + ArregloOtrosRamos(1, posOtros)%>
                    </font></b></font></div></td>
                  </tr>
                  <% If SoloPrueba = "N" then %>
                  <tr bgcolor="83a3d0">
                    <td height="10" colspan="1" width="10%" background="imagenes/fdo-cabecera-cel22.jpg"><div align="left"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="2">Motivo: </font></b></font></div></td>
                    <td height="10" colspan="6"><textarea name="<%="T" & ArregloOtrosRamos(0, posOtros)%>" cols="82" onBlur="javascript:GrabaTexto(this)" rows="1" wrap="PHYSICAL"></textarea>
                    </td>
                  </tr>
                  <% end if%>
                  <tr class="text-cabecera-celda" background="imagenes/fdo-cabecera-cel22.jpg" >
                    <td width="10%" height="30" background="imagenes/fdo-cabecera-cel22.jpg" ><div align="center"> <font size="1" face="Geneva, Arial, Helvetica, san-serif" color="#FFFFFF">Lista</font> <b><font size="1" face="Geneva, Arial, Helvetica, san-serif"> <br>
                                <font color="#FFFFFF"> Curso</font></font></b></div></td>
                    <td width="10%" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1" face="Geneva, Arial, Helvetica, san-serif">profesor</font></b></font></div></td>
                    <td width="5%" height="30" valign="middle" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1" face="Geneva, Arial, Helvetica, san-serif">sec.</font></b></font></div></td>
                    <td width="30%" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1" face="Geneva, Arial, Helvetica, san-serif">horario</font></b></font></div></td>
                    <td width="30%" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1" face="Geneva, Arial, Helvetica, san-serif">Ramo</font></b></font></div></td>
                    <td width="5%" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" color="#FFFFFF"><b><font size="1"> <font face="Geneva, Arial, Helvetica, san-serif">Cupo</font></font></b></font></div></td>
                    <td width="20%" height="30" background="imagenes/fdo-cabecera-cel22.jpg"><div align="center"><font face="Geneva, Arial, Helvetica, san-serif" size="1" color="#FFFFFF"><b>Solicitar</b></font></div></td>
                  </tr>
                  <%                 
                 'RSSeccion.Movelast
                 'RSSeccion.Movefirst
                 'pos = 1
                 'do while (pos <= Ubound(RamosSeccion, 2) )
                 '    Response.Write("Ramo " & RamosSeccion(0, pos) & "<br>")
                 '    Response.Write("Seccion " & RamosSeccion(1, pos) & "<br>")
                 '    pos = pos + 1
                 'loop             
                 pos = 0     
                 RamosSeccion = RstSeccion.getrows
                 'do while not RstSeccion.eof
                 do while (pos <= Ubound(RamosSeccion, 2) )                
                     'Response.Write("Ramo " & RamosSeccion(0, pos) & "<br>")
                     'Response.Write("Seccion " & RamosSeccion(1, pos) & "<br>")
                     'Response.Write("Seccion " & RamosSeccion(4, pos) & "<br>")
					 'session("OtrosRamos") = "N"
                     TipoCupo = "CT"
                     Select Case TipoRamo
                       Case "OBLIGATORIO"
                          'Cupos = SacaCupoAntiguo ( RstSeccion("CodRamo"), RstSeccion("CodSecc"), Ano, Periodo, CodCarr, CodSede, EsRepit, TipoCupo )
                          'Inscritos = CantidadInscritos ( RstSeccion("CodRamo"), RstSeccion("CodSecc"), CodSede, Ano, Periodo, TipoCupo )
                          
						  Cupos = SacaCupoAntiguo ( RamosSeccion(0, pos), RamosSeccion(1, pos), Ano, Periodo, CodCarr, CodSede, EsRepit, TipoCupo,ArregloOtrosRamos(0, posOtros))
						  
						  if session("RamoObliPlanificadoOp") = "N" then
	                      		Inscritos = CantidadInscritos ( RamosSeccion(0, pos), RamosSeccion(1, pos), CodSede, Ano, Periodo, TipoCupo, "T")
						  else
                                Inscritos = CantidadInscritosOptativos(ArregloOtrosRamos(0, posOtros), RamosSeccion(0, pos), RamosSeccion(1, pos), CodSede, Ano, Periodo,Codcarr, TipoCupo, "T")						  		
						  end if 
					   	'response.Write("Cupos :" & Cupos & " Inscritos: " & Inscritos )
				  
                       Case "OPCIONAL"
                          Cupos = SacaCupoOptativo(ArregloOtrosRamos(0, posOtros), RamosSeccion(0, pos), RamosSeccion(1, pos), Ano, Periodo, CodCarr,CodSede) 
                          Inscritos = CantidadInscritosOptativos(ArregloOtrosRamos(0, posOtros), RamosSeccion(0, pos), RamosSeccion(1, pos), CodSede, Ano, Periodo,Codcarr, TipoCupo, "T")
                       Case "OPTATIVO ESPECIAL"
                          Inscritos=0
                          if RevisaCupos then
                            Cupos = SacaCupoOptativoEspecial(RamosSeccion(0, pos), RamosSeccion(1, pos), Ano, Periodo, CodCarr,ArregloOtrosRamos(0, posOtros)) 
                            Inscritos = CantidadInscritosOptativoEspecial(RamosSeccion(0, pos), RamosSeccion(1, pos), CodSede, Ano, Periodo,Codcarr, TipoCupo, "T")
                            ArrCupos(pos, 1) = Cupos
                            ArrCupos(pos, 2) = Inscritos
                            ArrCupos(pos, 3) = RamosSeccion(0, pos)
                            ArrCupos(pos, 4) = RamosSeccion(1, pos)
                          else
                            'Response.Write("Posicion = " & RstSeccion.AbsolutePosition)
                            'Response.End
                            Cupos = ArrCupos(pos, 1)
                            Inscritos = ArrCupos(pos, 2)
                          end if
                     End Select

                  Response.Write("<tr bgcolor='#DBECF2'>" & chr(13))
                  Response.Write("<td width='10%' height='16' align='center' bgcolor='#DBECF2'>")
                  'Response.Write("<div align='center'><font size='1' face='Verdana, Arial, Helvetica, sans-serif'><a href='detalle-asignatura.asp?R=" & RstSeccion("CodRamo") & "&S=" & RstSeccion("CodSecc") & "' target='_blank'>" & RstSeccion("CodRamo") & "</a></font></div></td>")
                  Response.Write("<div align='center'><font size='1' face='Verdana, Arial, Helvetica, sans-serif'><a href='detalle-asignatura.asp?R=" & RamosSeccion(0, pos)& "&P=" & RamosSeccion(5, pos) + " " + RamosSeccion(6, pos) & "&S=" & RamosSeccion(1, pos) & "' target='_blank'>" & RamosSeccion(0, pos) & "</a></font></div></td>")

                  Response.Write("<td width='40%' height='16' align='center'>") 
                  'Response.Write("<div align='center'><font face='Verdana' size='1'>" & GetNombreProfe(RamosSeccion(4, pos)) & "</font></div></td>")
                  'Response.Write("<div align='center'><font face='Verdana' size='1'>" & RstSeccion("Ap_Pater") & " " & rstSeccion("Nombres") &  "</font></div></td>")
                  Response.Write("<div align='center'><font face='Verdana' size='1'>" & RamosSeccion(5, pos) & " " & RamosSeccion(6, pos) &  "</font></div></td>")
                  Response.Write("<td width='5%' height='16' align='center'> ")
                  Response.Write(" <div align='center'><font face='Verdana' size='1'>" & RamosSeccion(1, pos) & "</font></div></td>")
                  Response.Write("<td width='30%' height='16' align='center'>") 
                  'Response.Write(" <div align='center'><font face='Verdana' size='1'>" & GetHorario(RamosSeccion(0, pos), CodSede, RamosSeccion(1, pos), Ano, Periodo) & "</font></div></td>" )
                  Response.Write(" <div align='center'><font face='Verdana' size='1'>" & RamosSeccion(7, pos) & "</font></div></td>" )
                  Response.Write("<td width='30%' height='16' align='center'> ")
                  Response.Write(" <div align='center'><font face='Verdana' size='0'>" & RamosSeccion(2, pos) & "</font></div></td> ")
                  Response.Write(" <div align='center'><font face='Verdana' size='0'>" &  "</font></div></td> ")
                  Response.Write("<td width='5%' height='16' align='center'> ")
                  Response.Write(" <div align='center'><font face='Verdana' size='1'> ")
                  'Response.Write("C=" & Cupos & ",I=" & Inscritos)
                  if Cupos > Inscritos then
                     Response.Write("No")
                  else
                     Response.Write("Si")
                  end if
                  Response.Write("</font></div></td>")

                  Response.Write("<td width='20%' height='16' bgcolor='#DBECF2' valign='center'> ")
                  Response.Write("  <div align='center'><font face='Verdana' size='1'> " & chr(13))

                  'EstadoReal = EstadoRamoReal(CodCli, RamosSeccion(0, pos))
                         
                  If (EstadoReal = "A") or (EstadoReal = "E") or (EstadoReal = "I") then
                       Cursado = True
                  else
                       Cursado = False
                  end if
				  
                  if EstaPreInscrito(Codcli, ArregloOtrosRamos(0, posOtros), RamosSeccion(0, pos), "", Ano, Periodo, "T") then 
                     Response.Write(" <input type='radio' name=" & ArregloOtrosRamos(0, posOtros) & " value=" & RamosSeccion(0, pos) & _
                                    " OnFocus=" & chr(34) & "javascript:marca(this, '" & trim(RamosSeccion(0, pos)) &"')" & Chr(34) & _
                                    " OnClick=" & chr(34) & "javascript:GrabaSel(this, '" & ArregloOtrosRamos(0, posOtros) & "', '" & RamosSeccion(0, pos) & "', '" & RamosSeccion(1, pos) & "','" & CodCarr & "')" & chr(34) & " DISABLED " )
                     'If (Inscritos >= Cupos) or (cursado) then 
                     '  Response.Write("DISABLED") 
                     'End if
                     Response.Write(" > " & chr(13))
                  else 
                     Response.Write(" <input type='radio' name=" & ArregloOtrosRamos(0, posOtros) & " value=" & RamosSeccion(0, pos) & _
                                    " OnFocus=" & chr(34) & "javascript:marca(this, '" & trim(RamosSeccion(0, pos)) &"')" & Chr(34) & _
                                    " OnClick=" & chr(34) & "javascript:GrabaSel(this, '" & ArregloOtrosRamos(0, posOtros) & "', '" & RamosSeccion(0, pos) & "', '" & RamosSeccion(1, pos) & "','" & CodCarr & "')" & chr(34) & " " )
                    ' Response.Write(" <input type='radio' name=" & RS("CodRamo") & " value=" & RstSeccion("CodRamo") & _
                    '                " OnFocus=" & chr(34) & "javascript:marca(this, '" & trim(RstSeccion("CodRamo")) &"')" & Chr(34) & _
                    '                " OnClick=" & chr(34) & "javascript:GrabaSel(this, '" & RS("CodRamo") & "', '" & RstSeccion("CodRamo") & "', '" & RstSeccion("CodSecc") & "','" & CodCarr & "')" & chr(34) )
                     If (Inscritos >= Cupos) or (cursado) then 
                       'Response.Write("DISABLED") 
                     End if
                     Response.Write(" > " & chr(13))
                  end if
                  Response.Write(" </font> </div></td>" & chr(13))
                 pos = pos + 1
                 'RstSeccion.Movenext
                 loop
               end if
               Response.Write(" </table>")
               Response.Write("</form>")
 		       'Response.Write("<br>")
			 
    'RS.movenext
      end if
    end if
    posOtros = posOtros + 1
  loop 

 end if 

%>
				
                </table></td>
				</tr>
				</table>
              </form>
		  </td>
		  </tr>
		</table>
	</td>
  </tr>
</table>
</body>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->