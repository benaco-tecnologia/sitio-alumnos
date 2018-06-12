<%

Const CdoBodyFormatHTML = 0
Const CdoBodyFormatText = 1
Const CdoMailFormatMime = 0
Const CdoMailFormatText = 1
Const CdoLow    = 0
Const CdoNormal = 1
Const CdoHigh   = 2

Function SendMail(CodCli)
Dim StrSql,StrSqlc
Dim Rst,Rstc

  CodCli = Session("CodCli") 
  CodSede = Session("CodSede")
  Ano = Session("Ano")
  Periodo = Session("Periodo")

  StrSql= "SELECT b.codramo, b.nombre, a.inscrito, a.codsecc, a.Ramoequiv, c.nombre as NombreReal " & _ 
          "FROM ra_carga a with (nolock), ra_ramo b with (nolock), ra_ramo c with (nolock) " & _
          "WHERE a.codramo = b.codramo and " & _
          "a.ramoequiv *= c.codramo and " & _
          "a.inscrito = 'S' and " & _
          "a.codcli ='" & CodCli & "' and " & _
          "a.ano = '" & ano & "' and " & _
          "a.periodo = '" & periodo & "' Order By a.prioridad "
        
  Set Rstc=Conn.execute(StrSql)

  Mensaje = "<html><body> <br>"
  Mensaje = Mensaje & "<font face='Verdana, Arial, Helvetica, sans-serif' color='#4a5da1' size='2'>"
  Mensaje = Mensaje & "Si tienes problemas para ver el formato de este mensaje, activa la <br>"
  Mensaje = Mensaje & "visualización HTML de este mensaje.<br>"
  Mensaje = Mensaje & "Si usas el WebMail de la Universidad, esto lo podrás activar en :<br>"
  Mensaje = Mensaje & "Opciones --> Preferencias de Pantalla --> Mostrar versión HTML por omisión (SI).<br>"
  Mensaje = Mensaje & "Luego de esto debes seleccionar la opción ENVIAR (al final de la página).<br>"
  Mensaje = Mensaje & "SI usas Oultlook, Outlook Express u otro sistema de correo POP, activa la <br>"
  Mensaje = Mensaje & " visualización HTML </font>"
  Mensaje = Mensaje & " <br> "
  Mensaje = Mensaje & " <br> "
  Mensaje = Mensaje & " <b><font face='Verdana, Arial, Helvetica, sans-serif' color='#4a5da1' size='2'><b>Alumno : " & GetNombreAlumno(Codcli) &  "</font></b> <br>"
  Mensaje = Mensaje & " <b><font face='Verdana, Arial, Helvetica, sans-serif' color='#4a5da1' size='2'><b>Fecha : " & date &  "</font></b> "
  Mensaje = Mensaje & " <br> "
  Mensaje = Mensaje & " <b><font face='Verdana, Arial, Helvetica, sans-serif' color='#4a5da1'><b><font size='2'>Ud. Posee inscritas las siguientes asignaturas </b><br><br></font>"
  Mensaje = Mensaje & " <table border='0' cellpadding='0' cellspacing='0' align='left' width='584'> "
  Mensaje = Mensaje & " <tr bgcolor='4a5da1'>"
  Mensaje = Mensaje & " <td height='16' width='48' bgcolor='4a5da1'>  "
  Mensaje = Mensaje & " <font face='Verdana, Arial, Helvetica, sans-serif' color='#FFFFFF'><b><font size='1'> "
  Mensaje = Mensaje & " <div align='center'>C&oacute;digo</div> "
  Mensaje = Mensaje & " </font></b></font>"
  Mensaje = Mensaje & " </td> "
  Mensaje = Mensaje & " <td height='16' width='200'> "
  Mensaje = Mensaje & " <font face='Verdana, Arial, Helvetica, sans-serif' color='#FFFFFF'><b><font size='1'> "
  Mensaje = Mensaje & " <div align='center'>Asignatura</div> "
  Mensaje = Mensaje & " </font></b></font>"
  Mensaje = Mensaje & " </td> "
  Mensaje = Mensaje & " <td height='16' width='65'> "
  Mensaje = Mensaje & " <font face='Verdana, Arial, Helvetica, sans-serif' color='#FFFFFF'><b><font size='1'> "
  Mensaje = Mensaje & " <div align='center'>secci&oacute;n</div> "
  Mensaje = Mensaje & " </font></b></font>"
  Mensaje = Mensaje & " </td> "
  Mensaje = Mensaje & " <td height='16' width='91'> "
  Mensaje = Mensaje & " <font face='Verdana, Arial, Helvetica, sans-serif' color='#FFFFFF'><b><font size='1'> "
  Mensaje = Mensaje & " <div align='center'>Horario</div> "
  Mensaje = Mensaje & " </font></b></font>"
  Mensaje = Mensaje & " </td> "
  Mensaje = Mensaje & " <td height='16' width='200'> "
  Mensaje = Mensaje & " <font face='Verdana, Arial, Helvetica, sans-serif' color='#FFFFFF'><b><font size='1'> "
  Mensaje = Mensaje & " <div align='center'>Ramo</div> "
  Mensaje = Mensaje & " </font></b></font>"
  Mensaje = Mensaje & " </td> "
  Mensaje = Mensaje & " </tr> "

  do while not Rstc.eof
    
    Mensaje = Mensaje & " <tr bgcolor='ffc172' height='25'> "
    Mensaje = Mensaje & " <td width='48' height='8' align='center'><font face='Verdana, Arial, Helvetica, sans-serif' size='1'> " & Rstc("CodRamo") & "</font></td> "
    Mensaje = Mensaje & "  <td width='286' height='0' align='center'><font face='Verdana, Arial, Helvetica, sans-serif' size='1'> " & Rstc("Nombre") & " </font></td> "
    Mensaje = Mensaje & "  <td width='65' height='0' align='center'><font face='Verdana, Arial, Helvetica, sans-serif' size='1'>" & Rstc("CodSecc") & "</font></td> "
    Mensaje = Mensaje & "  <td width='91' height='16' align='center'><font face='Verdana, Arial, Helvetica, sans-serif' size='1'> " & GetHor2(Rstc("CodRamo"), Codsede, Rstc("CodSecc"), ano, Periodo ) & " </font></td> "
    Mensaje = Mensaje & "  <td width='91' height='16' align='center'><font face='Verdana, Arial, Helvetica, sans-serif' size='1'> " & Rstc("NombreReal") & " </font></td> "
    Mensaje = Mensaje & "  </tr> "

    Rstc.movenext
     
  loop
  
  StrSql = "Select e_mail from mt_webmail with (nolock) where Rut = '" & Session("RutAlum") & "' "
  Set rst = Conn.execute(StrSql)

  'Mensaje = Mensaje & "<tr><td>" & StrSql & " </td></tr>"
  'Mensaje = Mensaje & " </table></body></html> "
  Rstc.close()

  'if not rst.eof then
    Set objMail = CreateObject("CDONTS.NewMail")
    objMail.From = "ora@udd.cl"
    'objMail.To = Rst("e_mail")
    objMail.To = "amlapostol@udd.cl"
    'objMail.To = "rcristensen@bettersoft.cl"
    objMail.Subject = "Resultado de inscripcion de asignaturas"
    objMail.BodyFormat = CdoBodyFormatHTML
    objMail.MailFormat = CdoMailFormatMime
    objMail.Body = Mensaje
    objMail.Send
    Set objMail = Nothing
  'end if  
End Function

Function SendAnyMail()
    Set objMail = CreateObject("CDONTS.NewMail")
    objMail.From = "ora@udd.cl"
    objMail.To = "rcristensen@bettersoft.cl"
    ' Add the subject.
    objMail.Subject = "Resultado de inscripcion de asignaturas"
    ' Set the mail format.
    objMail.BodyFormat = CdoBodyFormatHTML
    objMail.MailFormat = CdoMailFormatMime
    ' Add the Web page.
    objMail.Body = "Prueba"
    ' Send the email.
    objMail.Send
  Set objMail = Nothing
End Function
%>