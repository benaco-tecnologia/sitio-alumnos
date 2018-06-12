<%
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
<table border="0" cellpadding="0" cellspacing="0">

<%session("TomadeRamo")=TomadeRamo(session("Codcarr"),session("Codcli"),session("ano_ed"),session("periodo_ed"))%>
  <tr>
   <td><a target="_top" href="#" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('A_r2_c4','','imagenes/menues1_r2_c1_f2.jpg',1);"><img name="A_r2_c4" src="imagenes/menues1_r2_c1.jpg" width="131" height="28" border="0" alt=""></a></td>
   <td><a target="_top" href="#" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('A_r2_c6','','imagenes/menues1_r2_c2_f2.jpg',1);"><img name="A_r2_c6" src="imagenes/menues1_r2_c2.jpg" width="113" height="28" border="0" alt=""></a></td>
   <td><a target="_top" href="#" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('A_r2_c7','','imagenes/menues1_r2_c3_f2.jpg',1);"><img name="A_r2_c7" src="imagenes/menues1_r2_c3.jpg" width="113" height="28" border="0" alt=""></a></td>
   <td><a target="_top" href="#" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('A_r2_c8','','imagenes/menues1_r2_c4_f2.jpg',1);"><img name="A_r2_c8" src="imagenes/menues1_r2_c4.jpg" width="113" height="28" border="0" alt=""></a></td>
    <td><a target="_top" href="#" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('_r2_c5','','images/-_r2_c5_f2.jpg',1);"><img name="_r2_c5" src="images/-_r2_c5.jpg" width="113" height="28" border="0" alt=""></a></td>
   <td><a target="_top" href="#" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('A_r2_c9','','imagenes/menues1_r2_c5_f2.jpg',1);"><img name="A_r2_c9" src="imagenes/menues1_r2_c5.jpg" width="113" height="28" border="0" alt=""></a></td>
   <td><a target="_top" href="#" onMouseOut="MM_swapImgRestore();" onMouseOver="MM_swapImage('A_r2_c10','','imagenes/menues1_r2_c6_f2.jpg',1);"><img name="A_r2_c10" src="imagenes/menues1_r2_c6.jpg" width="112" height="28" border="0" alt=""></a></td>
   
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
   
   <td><img src="imagenes/spacer.gif" width="1" height="28" border="0" alt=""></td>
  </tr>
  <td valign="top"n align="left"><img name="A_r2_c11" src="imagenes/esquina_sup_der.jpg" width="13" height="13" border="0" alt=""></td>
</table>
<%
end function
%>
