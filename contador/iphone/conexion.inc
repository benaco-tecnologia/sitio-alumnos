<%

Dim Conn, ConnectStr,RstTemp,strTMP
Dim STR_, NUM_, DAT_

STR_ = 0
NUM_ = 1
DAT_ = 2

'sConn =	"DRIVER={SQL SERVER};SERVER=" & cStr(Application("Servidor")) & _
'			";UID=" & cStr(Application("Usuario")) & _
'			";PWD=" & cStr(Application("password")) & _
'			";DATABASE=" & cStr(Application("Base")) & ";"
'	Set Session("Conn") = Server.CreateObject("ADODB.Connection")
'	Session("Conn").open sConn

'****** nuevo metodo, ahora reemplazamos la variable de conn, por una variable de sesión session("Conn")

Conn = Session("Conn")

Function BCL_ADO(strTMP, RstTemp) 
	'response.Write(strTMP)
	strTMP = f_NormSql(strTMP)
	'response.write(strTMP)
    Set RstTemp = Session("Conn").Execute(strTMP)
    BCL_ADO = Not (RstTemp.BOF And RstTemp.EOF)
End Function

Function f_NormSql(sQry)
Dim sNormQry
Dim i
   sQry = UCase(sQry)
   If (InStr(1, sQry, "SELECT") > 0) And (InStr(1, sQry, "FROM") > 0) Then
     posini = InStr(1, sQry, "FROM")
     posfin = InStr(posini + 4, sQry, "WHERE") - 1
     If posfin <= 0 Then posfin = InStr(posini + 4, sQry, "GROUP BY") - 1
     If posfin <= 0 Then posfin = InStr(posini + 4, sQry, "ORDER BY") - 1
     If posfin <= 0 Then posfin = Len(sQry)
     sNormQry = ""
     For i = posini + 4 To posfin
        If Mid(sQry, i, 1) = "," Then
          sNormQry = sNormQry & " WITH (NOLOCK),"
        Else
          sNormQry = sNormQry & Mid(sQry, i, 1)
        End If
     Next
     sNormQry = sNormQry & " WITH (NOLOCK) "
      f_NormSql = Mid(sQry, 1, posini + 4) + _
                  sNormQry + _
                  Mid(sQry, posfin + 1)
   Else
      f_NormSql = sQry
   End If
End Function

Function ValNulo(varDum, intTip) 
    If VarType(varDum) = vbNull Or varDum = "" Then
        Select Case intTip
            Case STR_
                ValNulo = ""
            Case NUM_
                ValNulo = 0
            Case DAT_
                ValNulo = 0 'Null
            Case Else
                ValNulo = "0"
        End Select
    Else
        If intTip = NUM_ Then
            'Response.write("var = " & VarType(varDum) & ":" & vbNull & ":" & vbString & ":" & varDum)
            'response.end
            ValNulo = cDbl(varDum)
        Else
            ValNulo = varDum
        End If
    End If
End Function
%>
