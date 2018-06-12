<%
Dim CodAlg, SecAlg, HorAlg

CodAlg = Request("CodAlg")
SecAlg = Request("SecAlg")
HorAlg = Request("HorAlg")


Session("CodAlg")=CodAlg
Session("SecAlg")=SecAlg
Session("HorAlg")=HorAlg

REsponse.write "Cod :" & Session("CodAlg") & "<br>"
REsponse.write "Sec :" & Session("SecAlg") & "<br>"
REsponse.write "Hor :" & Session("HorAlg") & "<br>"
%>
<script>
window.opener.location.href = "inscrip-asigna.asp";
window.close();
</script>