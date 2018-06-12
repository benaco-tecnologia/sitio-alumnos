<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#INCLUDE FILE="include/funciones.inc" -->

<%
str = "sp_traeDatosBiblioPearson_pa '" + Session("Rut") + "'"

set rst= Session("Conn").Execute(str)
if not rst.eof then
	un = "'" + rst(0) + "'"
	ul1 = "'" + rst(1) + "'"
	ul2 = "'" + rst(2) + "'"
	ue = "'" + rst(3) + "'" 
	i = "'" + rst(4) + "'"
	ik = "'" + rst(5) + "'"
end if 
                
%>
<script src="Scripts/jquery-2.1.1.js"></script>
    <script type="text/javascript">

        var un; //Nombre del usuario codificado.
        var ul1; //Apellido paterno del usuario codificado.
        var ul2; //Apellido materno del usuario codificado.
        var ue; //Correo electrÃ³nico del usuario codificado.
        var ui; //DirecciÃ³n IP del usuario codificada.
        var i; //ID de la instituciÃ³n en la base de datos de Pearson (entero)
        var ic; //ID del campus de la instituciÃ³n en la base de datos de Pearson (entero)
        var ik; //InstituciÃ³n Key en la base de datos de Pearson (entero)
        var res;
        var urlInitial = "";

        var attemptCount = 0;
        GetUrlBV();

        //Comienza la asignación de valores
        function GetUrlBV() {
            un = <%=un%>;
            ul1 = <%=ul1%>;
            ul2 = <%=ul2%>;
            ue = <%=ue%>;
            ui = ''; //vacío
            i = <%=i%>;
            ic = ''; //vacío
            ik = <%=ik%>;
            res = '';

            GetURLPage();
        };

        function GetURLPage() {
            //se concatenan lo valores asignados
            urlInitial = "https://www.biblionline.pearson.com/Services/GenerateURLAccess.svc/GetUrl?firstname=" + un + "&lastname1=" + ul1 + "&lastname2=" + ul2 + "&email=" + ue + "&ip=" + ui + "&idInstitution=" + i + "&idCampus=" + ic + "&institutionKey=" + ik + "&$callback=successCall&$format=json";
            $.ajax({
                dataType: "jsonp",
                contentType: "application/json; charset=utf-8",
                url: urlInitial,
                jsonpCallback: "successCall",
                error: function() {
                    if (attemptCount == 0) {
                        attemptCount++;
                        GetURLPage();
                    }
                    else {
                        alert("Error");
                    }
                },
                success: successCall
            });

            function parseJSON(jsonData) {

                return jsonData.Message;

            }


            function successCall(result) {

                res = parseJSON(result.GetUrlAccessResult);
                location.replace(res); //Redirecciona a la pÃ¡gina de la BV.
                //document.getElementById("lnkUrl").href = res; //Establece la URL de un hipervÃ­nculo en la pÃ¡gina.
                //document.getElementById("txtUrlResult").value = res; //Establece el valor de un elemento oculto con el texto de la URL generada.
                //document.getElementById("txtUrl").innerHTML = res; //Muestra el texto de la URL generada.
            }
        };

    </script>
    
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    
    </div>
    </form>
</body>
</html>