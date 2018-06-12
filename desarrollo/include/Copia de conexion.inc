<%
Dim Conn, ConnectStr,RstTemp,strTMP
Dim STR_, NUM_, DAT_

STR_ = 0
NUM_ = 1
DAT_ = 2

		
	Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.ConnectionTimeout = 100
        'Conn.Open "DSN=matricula","matricula","dtb01s"
		'ConnectStr = "PROVIDER=SQLOLEDB;Data Source=10.132.250.41;Initial Catalog=matricula_des"
		ConnectStr = "PROVIDER=SQLOLEDB;Data Source=TRUENO;Initial Catalog=matricula"
		Conn.Open ConnectStr,"matricula","dtb01s"


Function BCL_ADO(strTMP, RstTemp) 
    Set RstTemp = conn.Execute(strTMP)
    BCL_ADO = Not (RstTemp.BOF And RstTemp.EOF)
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