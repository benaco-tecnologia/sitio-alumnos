<%
	STR_ = 0
	NUM_ = 1
	DAT_ = 2

	Set uMas = Server.CreateObject("ADODB.Connection")
	uMas.ConnectionTimeout = 0
	ConnectStr="Driver={SQL Server};Server=172.168.1.26;UID=sa;PWD=vc.2012,mssql;Database=QA;"
   	uMas.Open ConnectStr


Function BCL_ADO(strTMP, RstTemp) 
	Set RstTemp = uMas.Execute(strTMP)
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
