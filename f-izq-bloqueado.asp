<%
function CargarMenu()
%>

<SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript"  SRC="include/validarut.js"></SCRIPT>
<script language="JavaScript">
<!--
function Ingresar()
{
	document.login.action="ValidaClave.asp";
	document.login.submit();
}
<!--
function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
 var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
   var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
   if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

//-->

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}
//-->
</script><style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
}
-->
</style>

<link href="css/tablas.css" rel="stylesheet" type="text/css">

<table border="0" cellpadding="0" cellspacing="0" align="right">
  <tr>
    <td colspan="3" valign="top">
	<table width="198" border="0" cellspacing="0" cellpadding="0" height="559">
        <tr valign="top">
          <td width="18" height="449"><img src="imagenes/-_r4_c1.jpg" name="_r4_c1" width="18" height="273" border="0"></td>
          <td width="162" height="449"><table width="162" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td><a href=MAILTO:czuniga@ipleones.cl target="_blank")"  onMouseOut="MM_swapImgRestore()"  onMouseOver="MM_swapImage('_r9_c2','','imagenes/Botones/-_r9_c2_f2.jpg',1);" ><img name="_r9_c2" src="imagenes/botones/-_r9_c2.jpg" width="162" height="27" border="0"></a></td>
              </tr>
              <tr>
                <td background="imagenes/-_r10_c2.jpg" valign="bottom" height="139"><table border="0" cellspacing="0" cellpadding="0" width="130" align="center" height="79">
                      <form name="login" method="post" onSubmit="javascript:return ValidaDatos();" >
                      <input type="hidden" name="rut" value="">
                      <input type="hidden" name="pin" value="">
					  <tr>
                        <td height="30" width="130"><div align="center">
                          <input name="logrut" type="text" class="casillas-form" id="logrut" size="13" onChange="javascript:checkRutField(this.value);" value=""  maxlength="12" disabled>
</div></td>
                      </tr>
                      <tr>
                        <td height="10" width="130"><div align="center"></div></td>
                      </tr>
                      <tr>
                        <td height="39" width="130"><div align="center">
                            <input name="logclave" type="password" class="casillas-form" id="logclave" size="13" maxlength="8" disabled >
                        </div></td>
                      </tr>
                    </form>
                </table></td>
              </tr>
              <tr>
                <td align="center">&nbsp;</td>
              </tr>
              <tr>
                <td valign="top"><img name="_r12_c2" src="imagenes/-_r12_c2.jpg" width="162" height="23" border="0"></td>
              </tr>
			<td valign="top"><img src="Imagenes/A-_r13_c2.jpg" width="162" height="149" ></td>
          </table>
          </td>
          <td height="449" width="18" valign="top"><img name="_r4_c3" src="imagenes/-_r4_c3.jpg" width="18" height="273" border="0"></td>
        </tr>
    </table></td>
  </tr>
</table>
<img src="imagenes/spacer.gif" width="1" height="28" border="0">
<%
end function
%>
