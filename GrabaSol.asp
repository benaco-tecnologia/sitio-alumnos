<% 
   Response.expires = -1 
%>


<!--#INCLUDE FILE="include/conexion.inc" -->

<%

Dim StrSql
Dim Rst

  'Response.write("A=" & Trim(Request("A")))
  Accion = Ucase(Trim(Request("A")))
  Curso = Trim(Request("C"))
  Ramo = Trim(Request("R"))
  Seccion = Trim(Request("S"))
  CodCli = Session("CodCli")
  Carrera = Request("CodCarr")
  Ano = Session("Ano")
  Periodo = Session("Periodo")
  Texto = Request("T")
  
  'Response.Write(" A grabar " & Texto) 
  'Response.End
  
  'Response.Redirect("Horario.asp?P=S")
  
  StrSql = "Select * from tmpSolici "
  StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
  StrSql = StrSql & "  And CodRamo = '" & Curso & "' "
  StrSql = StrSql & "  And Ano = " & Ano
  StrSql = StrSql & "  And Periodo = " & Periodo

  Set Rst = Session("Conn").Execute (StrSql)
  'Response.Write(StrSql + "<br>")
  'Response.End
  'Response.write("eof = " & Rst.eof)
  'Response.End

  if not Rst.eof then
    Rst.close()
    'Response.Write("Accion = " & Accion & "<br>")
    if Accion = "D" then
      StrSql = "delete tmpSolici "
      StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
      StrSql = StrSql & "  And CodRamo = '" & curso & "' "
      StrSql = StrSql & "  And Ano = " & Ano
      StrSql = StrSql & "  And Periodo = " & Periodo
    else
      if (Seccion <> 0) then
        StrSql = "Update tmpSolici set CodSecc = '" & Seccion & "', Glosa = '" & Texto & "', Ramoequiv = '" + Ramo + "' "
        StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
        StrSql = StrSql & "  And CodRamo = '" & Curso & "' "
        StrSql = StrSql & "  And Ano = " & Ano
        StrSql = StrSql & "  And Periodo = " & Periodo
      else
        StrSql = "Update tmpSolici set Glosa = '" & Texto & "' "
        StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
        StrSql = StrSql & "  And CodRamo = '" & Curso & "' "
        StrSql = StrSql & "  And Ano = " & Ano
        StrSql = StrSql & "  And Periodo = " & Periodo
      end if
    end if
    Session("Conn").Execute StrSql
    'Response.Write(Strsql)
    'Response.End
  else
    'Response.Write("Accion = " & Accion & "<br>")
    if Accion <> "D" then
      StrSql = "Insert into tmpSolici (Codcli, CodRamo, Codsecc, ano, Periodo, Glosa, codCarr, RamoEquiv) values ("
      StrSql = StrSql & "'" & Codcli & "'," 
      StrSql = StrSql & "'" & Curso & "'," 
      StrSql = StrSql & "'" & Seccion & "','" & Ano & "','" & Periodo & "', '" & Texto & "', '" & Carrera & "','" + Ramo + "' )" 

      'Response.Write(Strsql)
      Session("Conn").Execute StrSql
    end if
  end if
  
  'Response.Write(StrSql)
  'Response.End
  Response.Redirect("Horario.asp?P=N&S=S&T=N")
  
%>

<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
</HEAD>
<BODY>

<P>&nbsp;</P>

</BODY>
</HTML>
<!--#INCLUDE file="include/desconexion.inc" -->
