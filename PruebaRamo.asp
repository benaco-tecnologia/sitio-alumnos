<%@ Language=VBScript %>
<!--#INCLUDE FILE="include/conexion.inc" -->

<%
'response.Write("Hola")
'response.End()
Dim StrSql
Dim Rst

  Accion = Ucase(Trim(Request("A")))
  Ramo = Trim(Request("R"))
  Seccion = Trim(Request("S"))
  CodCli = Session("CodCli")
  Ano = Session("Ano")
  Periodo = Session("Periodo")
  
  StrSql = "Select * from tmpTest "
  StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
  StrSql = StrSql & "  And CodRamo = '" & Ramo & "' "
  StrSql = StrSql & "  And Ano = " & Ano
  StrSql = StrSql & "  And Periodo = " & Periodo

  Set Rst = Session("Conn").Execute (StrSql)
  
  if not Rst.eof then
    Rst.close()
    if Accion = "D" then
      StrSql = "delete tmpTest "
      StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
      StrSql = StrSql & "  And CodRamo = '" & Ramo & "' "
      StrSql = StrSql & "  And Ano = " & Ano
      StrSql = StrSql & "  And Periodo = " & Periodo
    else
      StrSql = "Update tmpTest set CodSecc = '" & Seccion & "' "
      StrSql = StrSql & " Where Codcli = '" & Codcli & "' "  
      StrSql = StrSql & "  And CodRamo = '" & Ramo & "' "
      StrSql = StrSql & "  And Ano = " & Ano
      StrSql = StrSql & "  And Periodo = " & Periodo

    end if
    Session("Conn").Execute StrSql
  else
    StrSql = "Insert into tmpTest (Codcli, CodRamo, Codsecc, ano, Periodo) values ("
    StrSql = StrSql & "'" & Codcli & "'," 
    StrSql = StrSql & "'" & Ramo & "'," 
    StrSql = StrSql & "'" & Seccion & "','" & Ano & "','" & Periodo & "' )" 
	
    Session("Conn").Execute StrSql
  end if
  
  	'response.Write(StrSql)
	'response.End()
	
  Response.Redirect("Horario.asp?P=S&T=S&S=S")

  
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