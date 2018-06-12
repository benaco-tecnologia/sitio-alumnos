<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%ScriptTimeOut = 100000%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title>Desarrollo \ Sincronizaci&oacute;n de plataformas</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link href="css/default.css" rel="stylesheet" type="text/css" />
<script language="javascript" type="text/javascript" src="http://code.jquery.com/jquery-1.5.1.min.js"></script>
<script language="javascript" type="text/javascript">
	$(document).ready(function(){
		$.ajaxSetup( {
			type: "POST",
			dataType: "html",
			timeout: 100000
		});
		
		// Sincronizacion de numero ecas
		$("#SyncNumeroEcas").click(function(){
			$.ajax({
			 	url: "sincronizar_numero_ecas.asp",
				beforeSend: function(objeto){
					$('#SyncNumeroEcas').attr("disabled", true);
            		$("#SyncNumeroEcasMsj").html("Sincronizando...");
        		},
				error: function(objeto, textStatus){
					$('#SyncNumeroEcas').attr("disabled", false);
					$("#SyncNumeroEcasMsj").html("Ocurri&oacute; un error al intentar sincronizar. " + textStatus);
				},
				success: function(datos){
					$('#SyncNumeroEcas').attr("disabled", false);
           			$("#SyncNumeroEcasMsj").html("Sincronizaci&oacute;n exitosa. " + datos);
        		}
			});
		});
		
		// Sincronizacion de alumnos y profesores
		$("#SyncAlumnosProfesores").click(function(){
			$.ajax({
			 	url: "sincronizar_usuarios.asp",
				beforeSend: function(objeto){
					$('#SyncAlumnosProfesores').attr("disabled", true);
            		$("#SyncAlumnosProfesoresMsj").html("Sincronizando...");
        		},
				error: function(objeto, textStatus){
					$('#SyncAlumnosProfesores').attr("disabled", false);
					$("#SyncAlumnosProfesoresMsj").html("Ocurri&oacute; un error al intentar sincronizar. " + textStatus);
				},
				success: function(datos){
					$('#SyncAlumnosProfesores').attr("disabled", false);
           			$("#SyncAlumnosProfesoresMsj").html("Sincronizaci&oacute;n exitosa. " + datos);
        		}
			});
		});
		
		// Sincronizacion ramos
		$("#SyncRamos").click(function(){
			var periodo = $('#SyncPeriodo').val();
			
			if (periodo == '1')
				periodo = 'Otoño';
			else
				periodo = 'Primavera';
				
			if (!confirm('Realmente desea sincronizar la asignacion de ramos para ' + periodo + ' ' + $('#SyncAnio').val()))
				return false;
				
			$.ajax({
			 	url: "sincronizar_ramos.asp",
				data: "ano=" + $('#SyncAnio').val() + "&periodo=" + $('#SyncPeriodo').val(),
				beforeSend: function(objeto){
					$('#SyncRamos').attr("disabled", true);
            		$("#SyncRamosMsj").html("Sincronizando...");
        		},
				error: function(objeto, textStatus){
					$('#SyncRamos').attr("disabled", false);
					$("#SyncRamosMsj").html("Ocurri&oacute; un error al intentar sincronizar. " + textStatus);
				},
				success: function(datos){
					$('#SyncRamos').attr("disabled", false);
					$("#SyncRamosMsj").html("Sincronizaci&oacute;n exitosa. " + datos);
        		}
			});
		});
		
		// Sincronizacion claves alumnos.ecas.cl
		$("#SyncClaveAlumnos").click(function(){
			$.ajax({
			 	url: "arreglar_usuarios_sinclave.asp",
				beforeSend: function(objeto){
					$('#SyncClaveAlumnos').attr("disabled", true);
            		$("#SyncClaveAlumnosMsj").html("Sincronizando...");
        		},
				error: function(objeto, textStatus){
					$('#SyncClaveAlumnos').attr("disabled", false);
					$("#SyncClaveAlumnosMsj").html("Ocurri&oacute; un error al intentar sincronizar. " + textStatus);
				},
				success: function(datos){
					$('#SyncClaveAlumnos').attr("disabled", false);
					$("#SyncClaveAlumnosMsj").html("Sincronizaci&oacute;n exitosa. " + datos);
        		}
			});
		});
	});
</script>
<style type="text/css" media="all">
body   
{
    font-size: .80em;
    font-family: "Helvetica Neue", "Lucida Grande", "Segoe UI", Arial, Helvetica, Verdana, sans-serif;
    margin: 10px;
    padding: 0px;
    color: #696969;
}
</style>
</head>

<body>
<p>Seleccione per&iacute;odo y a&ntilde;o a sincronizar: 
	<select name="SyncPeriodo" id="SyncPeriodo">
		<option value="1">Oto&ntilde;o</option>
		<option value="2">Primavera</option>
	</select>
	<select name="SyncAnio" id="SyncAnio">
		<option value="2015">2015</option>
		<option value="2011">2011</option>
		<option value="2010">2010</option>
	</select>
</p>
<table border="0" cellpadding="5">
	<tr>
		<td width="450" align="left"><h4>Alumnos y profesores</h4>Ingresa alumnos y profesores en Ecas Virtual.</td>
		<td width="80"><input name="SyncAlumnosProfesores" id="SyncAlumnosProfesores" type="button" value="Sincronizar" /></td>
		<td><span id="SyncAlumnosProfesoresMsj"></span></td>
	</tr>
	<tr>
		<td align="left"><h4>Ramos</h4>Asigna ramos a profesores y alumnos en Ecas Virtual.</td>
		<td><input name="SyncRamos" id="SyncRamos" type="button" value="Sincronizar" /></td>
		<td><span id="SyncRamosMsj"></span></td>
	</tr>
	<tr>
		<td align="left"><h4>N&uacute;mero ECAS</h4>Asigna el n&uacute;mero ECAS correspondiente a cada alumno en Ecas Virtual.</td>
		<td><input name="SyncNumeroEcas" id="SyncNumeroEcas" type="button" value="Sincronizar" /></td>
		<td><span id="SyncNumeroEcasMsj"></span></td>
	</tr>
	<tr>
		<td align="left"><h4>Usuarios sin clave</h4>Asigna clave de acceso a usuarios sin acceso a alumnos.ecas.cl (primeros 5 d&iacute;gitos de rut).</td>
		<td><input name="SyncClaveAlumnos" id="SyncClaveAlumnos" type="button" value="Sincronizar" /></td>
		<td><span id="SyncClaveAlumnosMsj"></span></td>
	</tr>
</table>
</body>
</html>