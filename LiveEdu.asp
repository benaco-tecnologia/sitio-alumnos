<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>http://login.live.com</title>
<script language="JavaScript" type="text/javascript">
function postSlt(liveLogonUrl, slt) {
            var tempSpan = window.document.createElement("span");
            var formHtml = "<form method='POST' action='" + liveLogonUrl + "' id='SLT_Form' name='SLT_Form'";
            formHtml += ">";
            formHtml += "<input type='hidden' name='slt' value='" + slt + "'>";
            formHtml += "</form>";

            tempSpan.innerHTML = formHtml;
            window.document.body.appendChild(tempSpan);

            var formTag = window.document.getElementById("SLT_Form");
            if (!formTag) {
                var formTags = window.document.getElementsByTagName("form");
                if (formTags) {
                    if (formTags.length) {
                        for (var index = 0; index < formTags.length; index++) {
                            if (formTags[index].id == "SLT_Form") {
                                formTag = formTags[index];
                                break;
                            }
                        }
                    }
                    else if (formTags.id == "SLT_Form") {
                        formTag = formTags;
                    }
                }
            }
            if (formTag) {
                formTag.submit();
            }
        }
        </script>
</head>
<body onLoad="javascript: <%  
     
	  dim strLive,rsLive,LiveSC
	 strLive = "Exec Pr_LiveEdu '"&Session("Rut")&"','" &  chr(34) & "Data Source="& Application("Servidor")&";Initial Catalog="&Application("Base")&";Persist Security Info=True;User ID="&Application("Usuario")&";Password="&Application("Password")&"" & chr(34) &"'"
	'response.Write(strLive)
	'response.End()


	Session("Conn").Execute(strLive)
	
	StrLive = "Select url from LiveEdu where rut = '"&Session("Rut")&"'"	
	Set rsLive = Session("Conn").execute(StrLive)
	if(not rsLive.EOF) then
	Set LiveSC = rsLive("url")
	'response.Write(rsLive("url"))
	'response.End()
	response.Write(rsLive("url"))
	else
	response.Write("alert('Ha ocurrido un error , contactese con el administrador'); window.close();")

	end if
	   %>" topmargin=0 > 
      <!-- <a href="javascript:; LiveCs "><img src="imagenes/botones/A-cerrar-sesion-of.jpg"  id="Image1" width="162" height="21" border="0"></a></div></td> -->    
</body>
</html>
