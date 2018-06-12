<table width='100%'  border='0' cellspacing='0' cellpadding='0'>
      <tr>
        <td><table style='width:100%' cellpadding='0' cellspacing='0'>
            <tr>
              <td><img src='addon/imagenes/box_header_left.png' alt='' width='6' height='6' align='absbottom' /></td>
              <td width='100%' background='addon/imagenes/box_header_bg.png'></td>
              <td><img src='addon/imagenes/box_header_right.gif' alt='' width='10' height='6' align='absbottom' /></td>
            </tr>
        </table></td>
      </tr>
      <tr>
        <td><table align='left' cellpadding='0' cellspacing='0' style='width:100%'>
            <tr>
              <td width='6' height='24' background='addon/imagenes/box_lado_left.png'><img src='addon/imagenes/box_lado_left.png' width='6' height='2' /></td>
              <td align='left' valign='top' background='addon/imagenes/filete_v.png' bgcolor='#C4CBD2'><table width='100%'  border='0' cellpadding='0' cellspacing='0'>
                  <tr>
                    <td><table width="100%" border="0" cellpadding="3" cellspacing="5">
                        <tr>
                          <td width="10"><img src="../../addon/imagenes/kshilupi2.gif" width="8" height="9"></td>
                          <td width="1190" class="normalname"><%=boxScroll_MsgTitulo %></td>
                        </tr>
                    </table></td>
                  </tr>
                </table>
                  <img src='addon/imagenes/horizontal_line.png' width='100%' height='8' /></td>
              <td width='10' background='addon/imagenes/box_lado_right.png'><img src='addon/imagenes/box_lado_right.png' width='10' height='8' /></td>
            </tr>
        </table></td>
      </tr>
      <tr>
        <td><table cellpadding='0' cellspacing='0' bgcolor='#C4CBD2' style='width:100%'>
            <tr>
              <td width='6' background='addon/imagenes/box_lado_left.png'><img src='addon/imagenes/box_lado_left.png' width='6' height='2' /></td>
              <td width='100%' bgcolor='#ffffff'><table width=100% border=0 align=center cellpadding=3 cellspacing=2 bgcolor='#C4CBD2'>
                  <tbody>
                    <tr>
                      <td align="left" valign="top" bordercolor='#CC0000' bgcolor='#E2E6EA'><script language="JavaScript1.2"><!--
// ancho
var marqueewidth=100
// alto
var marqueeheight=200
// velocidad
var speed=3
// contenido del scroll
var marqueecontents='<%=Scroll_txt %>'
if (document.all)
	document.write('<marquee direction="up" scrollAmount='+speed+' style="width:'+marqueewidth+';height:'+marqueeheight+'">'+marqueecontents+'</marquee>')
	
function regenerate(){
	window.location.reload()
}
function regenerate2(){
	if (document.layers){
	setTimeout("window.onresize=regenerate",450)
	intializemarquee()
	}
}
function intializemarquee(){
	document.cmarquee01.document.cmarquee02.document.write(marqueecontents)
	document.cmarquee01.document.cmarquee02.document.close()
	thelength=document.cmarquee01.document.cmarquee02.document.height
	scrollit()
}
function scrollit(){
	if (document.cmarquee01.document.cmarquee02.top>=thelength*(-1)){
	document.cmarquee01.document.cmarquee02.top-=speed
	setTimeout("scrollit()",100)
	} else {
			document.cmarquee01.document.cmarquee02.top=marqueeheight
			scrollit()
		}
}
window.onload=regenerate2
//--></script>
                      </td>
                    </tr>
                  </tbody>
              </table></td>
              <td width='10' background='addon/imagenes/box_lado_right.png'><img src='addon/imagenes/box_lado_right.png' width='10' height='2' /></td>
            </tr>
        </table></td>
      </tr>
      <tr>
        <td><table style='width:100%' cellpadding='0' cellspacing='0'>
            <tr>
              <td><img src='addon/imagenes/box_footer_left.gif' width='6' height='10' /></td>
              <td width='100%' background='addon/imagenes/box_footer_bg.png'></td>
              <td><img src='addon/imagenes/box_footer_right.png'  width='10' height='10' /></td>
            </tr>
        </table></td>
      </tr>
    </table>