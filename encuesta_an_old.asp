<!--#INCLUDE file="top.asp" -->
<!--#INCLUDE file="top1.asp" -->
<!--#INCLUDE file="f-izq.asp" -->
<!--#INCLUDE file="SubMenu.asp" -->
<!--#INCLUDE FILE="include/conexion.inc" -->
<!--#INCLUDE FILE="include/periodos.inc" -->
<!--#INCLUDE FILE="include/funciones.inc" -->
<!--#include file="include/lib.asp" -->
<!--#INCLUDE file="analytics.asp" -->

<%
strParame="SELECT coalesce(dbo.Fn_ValorParame('USAANALYTICS'),'NO')Parame"
set rstParame= Session("Conn").Execute(strParame)		
if not rstParame.eof then
	USAANALYTICS=rstParame("Parame")
end if

if USAANALYTICS="SI" then
	analitycs() 
end if  


  Function EsAlumnoNuevo_EAN()
	Dim StrSql
	Dim Rst
	   
	   StrSql = "Select 1 from mt_Alumno where Codcli = '" & Session("Codcli") & "' And Ano = Ano_mat and ano = dbo.Fn_ValorParame('anoadmision') " 
	   if BCL_ADO(StrSql, Rst) then
		  EsAlumnoNuevo_EAN = true
	   else 
		  EsAlumnoNuevo_EAN = false
	   end if
	End Function

%>

<%
'response.Write(getEncuestaVigente(2))


'Response.Write("paso")
'Response.Write(Session("CodCli"))
'Response.Write("<br>")
'Response.Write("***")
'Response.Write(getFaltaPorResponder(Session("CodCli"), cInt(getEncuestaVigente())))
'Response.Write("<br>")
'Response.Write("***")

'ENCUESTA INSTITUCIONAL = 2
'response.Write(

if EsAlumnoNuevo_EAN = false then 
	session("MensajeBloqueosVarios") ="Esta opción está disponible sólo para alumnos nuevos."
    response.Redirect("MensajeBloqueo.asp")
end if


if cInt(getEncuestaVigente(11)) = 0 	then
    'response.Write("no hay")
		response.Redirect("nohay.asp")
		response.End()
else
        'response.Write("si hay")
    if getFaltaPorResponder(Session("CodCli"), cInt(getEncuestaVigente(11)))<>0 then
        'response.Write("si hay")
		response.Redirect("yahecha.asp")
		response.End()
		end if
end if 

	if EstaHabilitadaNW (902)="S" then 
		if GetPermisoNW(902) ="N" then
			response.Redirect("MensajesBloqueos.asp")	
		end if
	else
		response.Redirect("MensajeBloqueoHabilita.asp")
	end if 
Audita 902,"Ingresa a Encuesta de alumnos nuevos" 
'Response.Write("paso2")

NroEncuesta= cInt(getEncuestaVigente(11))
Session("NroEncuesta") = NroEncuesta

'response.Write("0.1")

if NroEncuesta = 0 _
	then
		'response.Redirect("index.asp")
		response.Redirect("alumnos.asp")
		response.End()
end if

Set rsBloques = getBloques(NroEncuesta)

'response.Write("1")

AnnoEncuesta = getEncuestaVigenteAnno()
Session("AnnoEncuesta") = AnnoEncuesta


PeriodoEncuesta = getEncuestaVigentePeriodo()
Session("PeriodoEncuesta") = PeriodoEncuesta

'Response.Write(AnnoEncuesta)
'Response.Write(PeriodoEncuesta)


%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
    <title>Portal de Alumnos en L&iacute;nea</title>
    <link rel="stylesheet" href="css/tablas.css" type="text/css">

    <script language="javascript" type="text/javascript" src="include/prototype.js"></script>

    <script language="javascript" type="text/javascript"><!--//
if (history.forward(1)){
	location.replace(history.forward(1))	
} 

function jsProcesaEntrevista(){
	if (jsIsValidForm()){	    
		document.getElementById('btngrabar').disabled = true;
		$('FrmEntrevista').submit();
		
	}	
    return false;
}

function jsIsValidForm(){
	<%

							while not rsBloques.Eof										
								Set rsPreguntas = getPreguntas(NroEncuesta, rsBloques("Orden"))
										'Response.write("dentro del while")
									while not rsPreguntas.Eof
										%>
										if ($('Select_<%=rsBloques("Orden") & "_" & rsPreguntas("Numero")%>').value==0){
											alert('DEBE SELECCIONAR OPCION:\nBloque: <%=uCase(rsBloques("Descripcion"))%>\nPregunta: <%=rsPreguntas("Pregunta")%>\n');
											$('Select_<%=rsBloques("Orden") & "_" & rsPreguntas("Numero")%>').focus();
											return false;
										}
										<%
										rsPreguntas.MoveNext	
									wend
								rsBloques.MoveNext
							wend
							rsBloques.MoveFirst
	%>
	
	return true;
}
//--></script>

    <style type="text/css">
        <!
        -- .cssBloque
        {
            font-family: Verdana, Arial, Helvetica, sans-serif;
            font-weight: bold;
        }
        .cssPregunta
        {
            font-family: Verdana, Arial, Helvetica, sans-serif;
            font-size: 12px;
        }
        -- ></style>
    <body>
        <table border="0" cellpadding="0" cellspacing="0" align="left">
            <tr>
                <td>
                    <table width="200" border="0" cellpadding="0" cellspacing="0" align="center">
                        <tr>
                            <td colspan="2" valign="top" align="right">
                                <% CargarTop()%><% 
				if trim(Session("NomAlum"))="" then
					CargarMenu() 
				else
					CargarMenu2() 
				end if
                                %>
                            </td>
                            <td valign="top">
                                <% CargarTop1()%><% SubMenu()%>
                                <div style="width: 825px; height: 600px; overflow-x: hidden; overflow-y: auto;" align="left">
                                    <table width="800" border="0" cellpadding="5" cellspacing="15">
                                        <tr class="cssBloque">
                                            <td>
                                                <span style="font-size: 14px"><p style="font-size: 25px" class="text-menu">Encuesta de alumnos nuevos</p></span>
                                            </td>
                                        </tr>
                                        <tr class="cssBloque">
                                            <td align="center" bgcolor="#DBECF2" class="Tit-celdas" style="color: #000000">
                                                <div align="justify">
                                                  <p>Muy bienvenido(a)! Nos gustar&iacute;a conocerte mejor por lo que te pedimos nos respondas las siguientes preguntas:
¡Muchas gracias!”
</p>
                                                </div>
                                            </td>
                                        </tr>
                                    </table>
                                    <form id="FrmEntrevista" name="FrmEntrevista" action="graba_ei.asp" method="post">
                                    <div align="left">
                                        <%
					Set rsBloques = getBloques(NroEncuesta)
					
					while Not rsBloques.Eof
                                        %>
                                        <table width="99%" border="0" cellpadding="3" cellspacing="15" >
                                            <tr align="left">
                                                <td width="802" align="left">
                                                    <table width="766" border="0" cellpadding="5" cellspacing="1">
                                                        <tr class="cssBloque" background="Imagenes/fdo-cabecera-cel22.jpg">
                                                            <td colspan="5" background="Imagenes/fdo-cabecera-cel22.jpg">
                                                                <font size="3" class="text-cabecera-celda">
                                                                    <%=server.HTMLEncode(uCase(rsBloques("Descripcion")))%></font>
                                                            </td>
                                                        </tr>
                                                        <%
							Set rsPreguntas = getPreguntas(NroEncuesta, rsBloques("Orden"))
							
							while Not rsPreguntas.Eof
                                                        %>
                                                        <tr class="cssPregunta">
                                                            <td width="4%" align="right" valign="top" bgcolor="#DBECF2">
                                                                <%=rsPreguntas("Nro_Linea")%>)
                                                            </td>
                                                            <td colspan="3" align="left" valign="top" bgcolor="#DBECF2" class="tex-totales-celda">
                                                                <%=server.HTMLEncode(uCase(rsPreguntas("Pregunta")))%>
                                                            </td>
                                                            <td width="35%" align="left" valign="top" bgcolor="#DBECF2">
                                                                <%
										set rsRespuestas = getRespuestas(rsPreguntas("TipoRespuesta"))
                                                                %>
                                                                <select id="Select_<%=rsBloques("Orden") & "_" & rsPreguntas("Numero")%>" name="Select_<%=rsBloques("Orden") & "_" & rsPreguntas("Numero")%>">
                                                                    <option value="0">===[ SELECCIONE UNA OPCI&Oacute;N ]===</option>
                                                                    <%while not rsRespuestas.Eof
                                                                    %><option value="<%=rsBloques("Orden") & "_" & rsPreguntas("Numero") & "_" & rsRespuestas("Respuesta")%>">
                                                                        <%=server.HTMLEncode(uCAse(rsRespuestas("Descripcion")))%></option>
                                                                    <%rsRespuestas.MoveNext
										wend%>
                                                                </select>
                                                            </td>
                                                        </tr>
                                                        <%
								rsPreguntas.MoveNext
							wend
                                                        %>
                                                    </table>
                                                </td>
                                            </tr>
                                        </table>
                                        <%
						rsBloques.MoveNExt
					wend
                                        %>
                                        <table width="99%" border="0" cellspacing="15">
                                            <tr align="left">
                                                <td width="802" align="left">
                                                    <table width="761" border="0" cellpadding="5" cellspacing="0">
                                                        <tr class="cssBloque">
                                                            <td width="41%" align="left" valign="top" bgcolor="#DBECF2" class="Tit-celdas" style="color: #000000">
                                                                Registre otros comentarios que contribuyan a mejorar
                                                            </td>
                                                            <td width="59%" align="left" valign="top" bgcolor="#DBECF2" class="cssPregunta">
                                                                <textarea maxlength="500"name="txtObservaciones" cols="70" rows="5" wrap="virtual" id="txtObservaciones"></textarea>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                        </table>
                                        <br />
                                        <table align ="center">
                                        <input src="Imagenes/botones/grabar-of.gif" type="image"  id ="btngrabar" name="Submit" value="Procesar encuesta"
                                            onclick="return jsProcesaEntrevista()"   />
                                        </table>
                                    </div>
                                    </form>
                                    <div align="center">
                                        <br />
                                    </div>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </body>
</html>
