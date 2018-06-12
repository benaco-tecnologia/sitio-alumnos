<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	Response.Buffer = TRUE 
    Response.Expires = -1
%>
<!--#INCLUDE VIRTUAL="/certificados/conexion.inc"  -->
<!--#INCLUDE VIRTUAL="/include/funciones.inc" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" xml:lang="es-es" lang="es-es" dir="ltr">
<head>
  <base href="http://www.ecas.cl/" />
  <meta http-equiv="content-type" content="text/html; charset=utf-8" />
  <meta name="robots" content="index, follow" />
  <meta name="keywords" content="ECAS, ecas, ECAS, Contadores de Chile, Contadores de Santiago, auditoria" />
  <meta name="description" content="Escuela de Contadores auditores de Santiago" />
  <meta name="generator" content="Daniel Fuentes Busco @ tuky.cl" />
  <title>Validación de Certificados</title>

  <link href="http://www.ecas.cl/templates/jsn_epic_pro/favicon.ico" rel="shortcut icon" type="image/x-icon" />
  <script type="text/javascript" src="http://www.ecas.cl/media/system/js/mootools.js"></script>
  <!--[if gte IE 6]><link href="http://www.ecas.cl/components/com_chronocontact/themes/default/css/style1-ie6.css" rel="stylesheet" type="text/css" /><![endif]-->
                <!--[if gte IE 7]><link href="http://www.ecas.cl/components/com_chronocontact/themes/default/css/style1-ie7.css" rel="stylesheet" type="text/css" /><![endif]-->
                <!--[if !IE]> <--><link href="http://www.ecas.cl/components/com_chronocontact/themes/default/css/style1.css" rel="stylesheet" type="text/css" /><!--> <![endif]-->
                        <script type="text/javascript">
			var CF_LV_Type = 'default';			</script>
            <link rel="stylesheet" href="http://www.ecas.cl/components/com_chronocontact/css/calendar2.css" type="text/css" />
            <link href="http://www.ecas.cl/components/com_chronocontact/css/tooltip.css" rel="stylesheet" type="text/css" />

            <script type="text/javascript" src="http://www.ecas.cl/components/com_chronocontact/js/calendar2.js"></script>
            <script src="http://www.ecas.cl/components/com_chronocontact/js/livevalidation_standalone.js" type="text/javascript"></script>
            <link href="http://www.ecas.cl/components/com_chronocontact/css/consolidated_common.css" rel="stylesheet" type="text/css" />
			<script src="http://www.ecas.cl/components/com_chronocontact/js/customclasses.js" type="text/javascript"></script>
            <script type="text/javascript">
	Element.extend({
		getInputByName2 : function(nome) {
			el = this.getFormElements().filterByAttribute('name','=',nome)
			return (el)?(el.length)?el:el:false;
		}
	});
	window.addEvent('domready', function() {
																									});
</script>			
        			                        <script type="text/javascript">	
                window.addEvent('domready', function() {
                                });
            </script>
					<style type="text/css">
			span.cf_alert {
				background:#FFD5D5 url(http://www.ecas.cl/components/com_chronocontact/css/images/alert.png) no-repeat scroll 10px 50%;
				border:1px solid #FFACAD;
				color:#CF3738;
				display:block;
				margin:15px 0pt;
				padding:8px 10px 8px 36px;
			}
			
			.info, .success, .warning, .error, .validation {
				border: 1px solid;
				margin: 10px 10px;
				margin-left: 30px;
				padding:15px 10px 15px 50px;
				background-repeat: no-repeat;
				background-position: 10px center;
			}
			
			.success {
				color: #4F8A10;
				background-color: #DFF2BF;
				background-image:url('http://alumnos.ecas.cl/certificados/v2/success.png');
			}
			.warning {
				color: #9F6000;
				background-color: #FEEFB3;
				background-image: url('http://alumnos.ecas.cl/certificados/v2/warning.png');
			}
			.error {
				color: #D8000C;
				background-color: #FFBABA;
				background-image: url('http://alumnos.ecas.cl/certificados/v2/error.png');
			}
		</style>	
		
				
                <script src="http://www.ecas.cl/components/com_chronocontact/js/jsvalidation2.js" type="text/javascript"></script>

        	<script type='text/javascript'>
				var fieldsarray = new Array();
				var fieldsarray_count = 0;window.addEvent('domready', function() {
				elementExtend();setValidation("ChronoContact_Certificados", 1, 0, 0);});</script>	
        	<script type="text/javascript">
	elementExtend();
	window.addEvent('domready', function() {	
		});
</script>

<link rel="shortcut icon" href="http://www.ecas.cl/images/favicon.ico" />
<link rel="stylesheet" href="http://www.ecas.cl/templates/system/css/system.css" type="text/css" />
<link rel="stylesheet" href="http://www.ecas.cl/templates/system/css/general.css" type="text/css" />
<link href="http://www.ecas.cl/templates/jsn_epic_pro/css/template.css" rel="stylesheet" type="text/css" media="screen" />
<link href="http://www.ecas.cl/templates/jsn_epic_pro/css/template_grey.css" rel="stylesheet" type="text/css" media="screen" /><link href="http://www.ecas.cl/templates/jsn_epic_pro/css/jsn_iconlinks.css" rel="stylesheet" type="text/css" media="screen" /><link href="http://www.ecas.cl/templates/jsn_epic_pro/ext/cb/style.css" rel="stylesheet" type="text/css" media="screen" /><link href="http://www.ecas.cl/templates/jsn_epic_pro/ext/cb/style_grey.css" rel="stylesheet" type="text/css" media="screen" /><link href="http://www.ecas.cl/templates/jsn_epic_pro/ext/docman/style.css" rel="stylesheet" type="text/css" media="screen" /><link href="http://www.ecas.cl/templates/jsn_epic_pro/ext/docman/style_grey.css" rel="stylesheet" type="text/css" media="screen" /><link href="http://www.ecas.cl/templates/jsn_epic_pro/ext/vm/style.css" rel="stylesheet" type="text/css" media="screen" /><link href="http://www.ecas.cl/templates/jsn_epic_pro/ext/jevents/style.css" rel="stylesheet" type="text/css" media="screen" /><link href="http://www.ecas.cl/templates/jsn_epic_pro/ext/jevents/style_grey.css" rel="stylesheet" type="text/css" media="screen" /><link href="http://www.ecas.cl/templates/jsn_epic_pro/ext/rsg2/style.css" rel="stylesheet" type="text/css" media="screen" /><style type="text/css">
body {
	margin: 0;
	padding: 0;	
}
	#jsn-page {
		width: 680px;
	}
	
	#jsn-header {
		height: 68px;
	}
	
	#jsn-pinset {
		right: 86px;
	}
	
	#jsn-puser9 {
		float: left;
		width: 19%;
	}
	#jsn-pheader {
		float: left;
		width: 100%;
	}
	#jsn-puser8 {
		float: right;
		width: 21%;
	}
	
	#jsn-content_inner1 {
		background: transparent url(http://www.ecas.cl/templates/jsn_epic_pro/images/bg/leftside19-bg-full.png) repeat-y 19% top;
		padding: 0;
	}
	#jsn-maincontent_inner {
		padding-left: 0;
	}
	
	#jsn-leftsidecontent {
		float: left;
		width: 19%;
	}
	#jsn-maincontent {
		float: left;
		width: 80.94%;
	}
	#jsn-rightsidecontent {
		float: right;
		width: 21%;
	}
	
			ul.menu-icon li.order1 a:link,
			ul.menu-icon li.order1 a:visited {
				background-image: url("http://www.ecas.cl/templates/jsn_epic_pro/images/icon-module-home.png");
			}
			
			ul.menu-icon li.order2 a:link,
			ul.menu-icon li.order2 a:visited {
				background-image: url("http://www.ecas.cl/templates/jsn_epic_pro/images/icon-module-info.png");
			}
			
			ul.menu-icon li.order3 a:link,
			ul.menu-icon li.order3 a:visited {
				background-image: url("http://www.ecas.cl/templates/jsn_epic_pro/images/icon-module-image.png");
			}
			
			ul.menu-icon li.order4 a:link,
			ul.menu-icon li.order4 a:visited {
				background-image: url("http://www.ecas.cl/templates/jsn_epic_pro/images/icon-module-download.png");
			}
			
			ul.menu-icon li.order5 a:link,
			ul.menu-icon li.order5 a:visited {
				background-image: url("http://www.ecas.cl/templates/jsn_epic_pro/images/icon-module-mail.png");
			}
			
			ul.menu-icon li.order6 a:link,
			ul.menu-icon li.order6 a:visited {
				background-image: url("http://www.ecas.cl/templates/jsn_epic_pro/images/icon-module-comment.png");
			}
			
	#jsn-master {
		font-size: 80%;
		font-family: "Cambria", Helvetica, sans-serif;
	}
	
	h1, h2, h3, h4, h5, h6,
	ul.menu-suckerfish a,
	.componentheading, .contentheading {
		font-family: Georgia, serif !important;
	}
	</style><script type="text/javascript" src="http://www.ecas.cl/templates/jsn_epic_pro/js/jsn_script.js"></script>
	<script type="text/javascript">
		var defaultFontSize = 80;
	</script>
	<script type="text/javascript" src="http://www.ecas.cl/templates/jsn_epic_pro/js/jsn_epic.js"></script>

	<!--[if lte IE 6]>
<link href="http://www.ecas.cl/templates/jsn_epic_pro/css/jsn_fixie6.css" rel="stylesheet" type="text/css" />
<style type="text/css">
	img {  behavior: url(http://www.ecas.cl/templates/jsn_epic_pro/js/iepngfix.htc); }
</style>
<![endif]-->
<!--[if lte IE 7]>
<script type="text/javascript" src="http://www.ecas.cl/templates/jsn_epic_pro/js/suckerfish.js"></script>
<![endif]-->
<!--[if IE 7]>
<link href="http://www.ecas.cl/templates/jsn_epic_pro/css/jsn_fixie7.css" rel="stylesheet" type="text/css" />
<![endif]-->
</head>
<body id="jsn-master">
	<div id="jsn-page">
<div id="jsn-header">
  <div id="jsn-logo"><a href="http://www.ecas.cl/" title="Escuela de Contadores Auditores de Santiago"><img src="http://www.ecas.cl/templates/jsn_epic_pro/images/logo.png" width="422" height="68" alt="Escuela de Contadores Auditores de Santiago" /></a></div>
</div>

		<div id="jsn-body">
		  <div id="jsn-content"><div id="jsn-content_inner1"><div id="jsn-content_inner2">
								<div id="jsn-leftsidecontent" class="jsn-column">
					<div id="jsn-pleft">		<div class="module">

			<div>
				<div>
					<div>
											<p><img src="http://www.ecas.cl/images/stories/conceptual/conceptual02.jpg" border="0" width="148" height="434" /></p>					</div>
				</div>
			</div>
		</div>
	</div>

				</div>
								<div id="jsn-maincontent" class="jsn-column"><div id="jsn-maincontent_inner">
										
															<div id="jsn-mainbody">

<% 
	If Request.Form("btnValidar") <> "" Then
		Set rs = Server.CreateObject("ADODB.RecordSet")
		StrSql = "select top(1) rut, fecha, institucion, validez from certificados_generados where codigo_verificacion = '" & Request.Form("codigo_verificacion") & "' order by id_generado desc;"
		rs.Open strSql,Conn
		
		if not rs.EOF then
			fecha_ = DateAdd("d", rs.Fields(3), rs.Fields(1))
			fecha_ = day(fecha_) & "/" &month(fecha_)&  "/" &year(fecha_)
			if (DateDiff("d",DateAdd("d", rs.Fields(3), rs.Fields(1)), Now()) < 0) Then 
				sql = "INSERT INTO certificados_validados (resultado, codigo_verificacion) VALUES ('success', '"&Request.Form("codigo_verificacion")&"'); select TOP 1 rut from certificados_generados;"
				'Set rs = Server.CreateObject("ADODB.RecordSet")
				'rs.Open sql,Conn
				%>
		
        	<div class="success">El código ingresado corresponde a un certificado válido hasta el <%=fecha_%>.<br />
  <a href="http://www.ecasvirtual.cl/certificados/Certificado Ecas (Cod. <%Response.Write(UCase(Request.Form("codigo_verificacion")))%>).pdf" target="_blank">Haga clic acá para ver el certificado.</a></div>
<%
			Else
				sql = "INSERT INTO certificados_validados (resultado, codigo_verificacion) VALUES ('warning', '"&Request.Form("codigo_verificacion")&"'); select TOP 1 rut from certificados_generados;"
				'Set rs = Server.CreateObject("ADODB.RecordSet")
				'rs.Open sql,Conn
				%>
        	<div class="warning">El código ingresado corresponde a un certiticado vencido el <%=fecha_%>.</div>
        <%
			End If
		Else
			sql = "INSERT INTO certificados_validados (resultado, codigo_verificacion) VALUES ('error', '"&Request.Form("codigo_verificacion")&"'); select TOP 1 rut from certificados_generados;"
			'Set rs = Server.CreateObject("ADODB.RecordSet")
			'rs.Open sql,Conn
		%>
       <div class="error">El código ingresado no corresponde a ningún certificado existente.</div> 
        <%
		end if
%>
<div class="form_item">
  <div class="cfclear">&nbsp;</div>
</div>

<div class="form_item">
  <div class="form_element cf_button">
  <input type="reset" name="btnCerrar" value="Aceptar" onClick="window.close();" style="float:none; clear:none; display:inline" /> <input type="reset" name="btnVolver" value="Ingresar otro código" onClick="location.href='http://alumnos.ecas.cl/certificados/';" style="float:none; clear:none; display:inline;" />
  </div>
<%
	Else 
%>
									
			                        						        <form name="ChronoContact_Certificados" id="ChronoContact_Certificados" method="post" action="http://alumnos.ecas.cl/certificados/validacion.asp">
		
				<div class="form_item">
  <div class="form_element cf_textbox">
    <label class="cf_label" style="width:300px;" for="codigo_verificacion">Ingrese el código de verificación del certificado</label>
    <input class="cf_inputbox required" size="30" title="Debes ingresar el código!" id="text_1" name="codigo_verificacion" type="text" maxlength="8" />
  
  </div>
  <div class="cfclear">&nbsp;</div>
</div>

<div class="form_item">
  <div class="form_element cf_textbox"></div>
</div>
<div class="form_item">
  <div class="cfclear">&nbsp;</div>
</div>

<div class="form_item">
  <div class="form_element cf_button">

    <input value="Aceptar" name="btnValidar" type="submit" /><input type="reset" name="btnCerrar" value="Cancelar" onClick="window.close();" />
  </div>
  <% End If %>
  <div class="cfclear">&nbsp;</div>
</div>
		<input type="hidden" name="2b6a0231f640c7ed9f5a2178cf90d4e8" value="1" />	
                	<input type="hidden" name="1cf1" value="0c1916161a125bb32cb8871493375565" />
                </form>		
					</div>
														</div></div>
								<div class="clearbreak"></div>
			</div></div></div>
		  </div>
</div>
	
</body>
</html>