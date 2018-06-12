<%
Dim CodCalc, SecCalc, HorCalc

CodCalc = Request("CodCalc")
SecCalc = Request("SecCalc")
HorCalc = Request("HorCalc")


Session("CodCalc")=CodCalc
Session("SecCalc")=SecCalc
Session("HorCalc")=HorCalc

REsponse.write "Cod :" & Session("CodCalc") & "<br>"
REsponse.write "Sec :" & Session("SecCalc") & "<br>"
REsponse.write "Hor :" & Session("HorCalc") & "<br>"
%>
<script>
window.opener.location.href = "inscrip-asigna.asp";
window.close();
</script>