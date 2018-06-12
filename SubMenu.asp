<!--#INCLUDE FILE="include/funciones.inc" -->
<%
Function BCL_ADO(strTMP, RstTemp) 
    Set RstTemp = Session("Conn").Execute(strTMP)	
    BCL_ADO = Not (RstTemp.BOF And RstTemp.EOF)
	
End Function 

function SubMenu()
%>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</script>
<script languaje="javascript">
function Enviar3(pag,valor,validaencuesta) 
{
	if (validaencuesta=='SI') 
	{	
		if (valor==0) { 
			alert("Para acceder a esta opci\u00f3n debe responder sus Encuestas Pendientes");
			return;
		}
		top.location.href = pag;
	}
	else
	{
		top.location.href = pag;
	}
}
</script>
<style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
-->
</style>
<%
strParame = "select dbo.Fn_ValorParame('BLOQUEAPAENCUESTAS')Parame"
set rstParame = Session("Conn").execute(strParame)
if not rstParame.EOF then
	BLOQUEAPAENCUESTAS=rstParame("Parame")
else
	BLOQUEAPAENCUESTAS=""
end if 

%>
<%session("TomadeRamo")=TomadeRamo(session("Codcarr"),session("Codcli"),session("ano_ed"),session("periodo_ed"))%>
<table border="0" cellpadding="0" cellspacing="0">
    <tr>
   <td><a  target="_top" href="javascript:Enviar3('<%="menu_tomaderamos.asp"%>','<%=session("TomadeRamo")%>','<%=BLOQUEAPAENCUESTAS%>');" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('_r2_c1','','images/-_r2_c1_f2.jpg',1);"><img name="_r2_c1" src="images/-_r2_c1.jpg" width="131" height="28" border="0" alt=""></a></td>
   <td><a  target="_top" href="menu_encuestas.asp" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('_r2_c2','','images/-_r2_c2_f2.jpg',1);"><img name="_r2_c2" src="images/-_r2_c2.jpg" width="113" height="28" border="0" alt=""></a></td>
   <td><a  target="_top" href="javascript:Enviar3('<%="menu_material.asp"%>','<%=session("TomadeRamo")%>','<%=BLOQUEAPAENCUESTAS%>');" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('_r2_c3','','images/-_r2_c3_f2.jpg',1);"><img name="_r2_c3" src="images/-_r2_c3.jpg" width="113" height="28" border="0" alt=""></a></td>
   <td><a  target="_top" href="javascript:Enviar3('<%="menu_consultas.asp"%>','<%=session("TomadeRamo")%>','<%=BLOQUEAPAENCUESTAS%>');" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('_r2_c4','','images/-_r2_c4_f2.jpg',1);"><img name="_r2_c4" src="images/-_r2_c4.jpg" width="113" height="28" border="0" alt=""></a></td>
   <td><a target="_top" href="menu_finanzas.asp" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('_r2_c5','','images/-_r2_c5_f2.jpg',1);"><img name="_r2_c5" src="images/-_r2_c5.jpg" width="113" height="28" border="0" alt=""></a></td>
   <td><a target="_top" href="javascript:Enviar3('<%="menu_perfil.asp"%>','<%=session("TomadeRamo")%>','<%=BLOQUEAPAENCUESTAS%>');" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('_r2_c6','','images/-_r2_c6_f2.jpg',1);"><img name="_r2_c6" src="images/-_r2_c6.jpg" width="113" height="28" border="0" alt=""></a></td>
   <td><a target="_top" href="javascript:Enviar3('<%="menu_ayuda.asp"%>','<%=session("TomadeRamo")%>','<%=BLOQUEAPAENCUESTAS%>');" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('_r2_c7','','images/-_r2_c7_f2.jpg',1);"><img name="_r2_c7" src="images/-_r2_c7.jpg" width="113" height="28" border="0" alt=""></a></td>
   <td>
   <% 
   strPar = "select dbo.Fn_ValorParame('USAMAIL') as valor"
	set rsPar = Session("Conn").execute(strPar)
	if (not rsPar.EOF) then
	if (rsPar("VALOR") = "S" or rsPar("VALOR") = "SI") then
   If Session("mail_inst") <> empty Then %>
	 
    <a target="_top" href="javascript:Enviar3('<%="mail.asp"%>','<%=session("TomadeRamo")%>','<%=BLOQUEAPAENCUESTAS%>');" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('_r2_c9','','images/-_r2_c9_f2.jpg',1);"><img name="_r2_c9" src="images/-_r2_c9.jpg" width="113" height="28" border="0" alt=""></a> 
     
    <% else %>
    
       <a target="_top" href="javascript:Enviar3('<%="nomail.asp"%>','<%=session("TomadeRamo")%>','<%=BLOQUEAPAENCUESTAS%>');" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('_r2_c9','','images/-_r2_c9_f2.jpg',1);"><img name="_r2_c9" src="images/-_r2_c9.jpg" width="113" height="28" border="0" alt=""></a> 
      
<% end if 
end if
	  	 end if
%>
   
   </td>
   <td><img name="_r2_c8" src="images/-_r2_c8.jpg" width="17" height="28" border="0" alt=""></td>
   <td><img src="images/spacer.gif" width="1" height="28" border="0" alt=""></td>
  </tr>
  <td valign="top"n align="left"><img name="A_r2_c11" src="imagenes/esquina_sup_der.jpg" width="13" height="13" border="0" alt=""></td>
  
</table>
<%
strParameCO="SELECT coalesce(dbo.Fn_ValorParame('ACTIVABOTONCERTIFONLINE'),'')Parame,coalesce(dbo.Fn_ValorParame('ACTIVABOTONSOLICITUDES'),'')Parame2	"
set rstParameCO= Session("Conn").Execute(strParameCO)		
if not rstParameCO.eof then
		ACTIVABOTONCERTIFONLINE = rstParameCO("Parame")
		ACTIVABOTONSOLICITUDES = rstParameCO("Parame2")
	else
		ACTIVABOTONCERTIFONLINE="" 
		ACTIVABOTONSOLICITUDES =""
end if
%>
<table border="0" cellpadding="0" cellspacing="0">
    <tr> 
        <td width="473" valign="top"><p>&nbsp;</p></td>
        <td width="473" valign="top"><p>&nbsp;</p></td>
        
        <%if ACTIVABOTONCERTIFONLINE = "SI" then
		
			dim URLCERTIF
			Opcion0="721"
			habilitada0=EstaHabilitadaNW (Opcion0)
			url0=GetPermisoNW (Opcion0)
			if url0 = "N" then
				Permiso0 = 0
			else
				Permiso0 = 1
			end if
					
			if trim(habilitada0)="S"  then
			 if Permiso0=1 then
				URLCERTIF="RedirectNet.asp?PaginaRedirect="&url0
			  else
				URLCERTIF="MensajesBloqueos.asp" 
			  end if
			 else
				  URLCERTIF= "MensajeBloqueoHabilita.asp"
			end if
		
		 %>
            <td width="204">
                <div align="right"><a href="<%=URLCERTIF%>"><img src="Imagenes/botones/certifonline.jpg" width="113" height="50"
                        name="ImageIsline" border="0"></a>
                </div>
            </td>
        <%end if%>        
        
        <%if ACTIVABOTONSOLICITUDES = "SI" then
			
			strSol ="SP_TRAE_DATOS_SOL_PA '"& session("codcli") &"'"
			if bcl_ado(strSol,Rstsol)  then
				 SolRut = Rstsol("RUT")
				 SolCodcarr = Rstsol("CODCARR")
				 SolNombre = Rstsol("NOMBRE")
				 SolCarrera = Rstsol("CARRERA")
				 SolPaterno = Rstsol("PATERNO")
				 SolMaterno = Rstsol("MATERNO")
				 SolMail = Rstsol("MAIL")
			end if 
		
			urlSol = "http://www2.iplacex.cl/form/index.php?"
			urlSol = urlSol & "rut="& SolRut &" "
			urlSol = urlSol & "codCar="& SolCodcarr &"&"
			urlSol = urlSol & "nombre="& SolNombre &"&"
			urlSol = urlSol & "carrera="& SolCarrera &"&"
			urlSol = urlSol & "apPaterno="& SolPaterno &"&"
			urlSol = urlSol & "apMaterno="& SolMaterno &"&"
			urlSol = urlSol & "mail="& SolMail &""
			
			%>
			<td width="133">
				<div align="right"><a href="<%=urlSol%>"><img src="Imagenes/botones/solicitudesenlinea.jpg" width="113" height="50"
					name="ImageIsline" border="0"></a>
				</div>
			</td>
        <%end if%>  
    
    </tr>
</table>



<%
end function
%>
