
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#INCLUDE FILE="include/conexion.inc" -->

<html>
<head>
<title>::Contador::</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<%
strParamePOP="SELECT dbo.Fn_ValorParame('ACTIVATRAUTO')ACTIVATRAUTO"
set rstParamePOP= Session("Conn").Execute(strParamePOP)

if not rstParamePOP.eof then
	ACTIVATRAUTO = rstParamePOP("ACTIVATRAUTO")
else
	ACTIVATRAUTO =""
end if

if ACTIVATRAUTO = "SI" then%>
    <body>
    Contador del Sitio: <% response.write (application("auxactivos"))%><br>
    Alumnos Inscritos 
    <% 
    
    AnoP = 0
    PeriodoP = 0
    FechaP = ""
        
    strParame= "select dbo.Fn_ValorParame('ANO')ANO,dbo.Fn_ValorParame('PERIODO')PERIODO,dbo.Fn_ValorParame('FECHATRA')FECHATRA"
    Set RstParame = Session("Conn").Execute(strParame)
    if not RstParame.eof then
        AnoP = RstParame("ANO")
        PeriodoP = RstParame("PERIODO")
        FechaP = RstParame("FECHATRA")
    end if 
    %>
    Año   <%=AnoP%> Periodo  <%=PeriodoP%> :
    <%
    set rst = createobject("adodb.recordset")
    rst.cursortype = 1
    
    rst.open "select count ( * ) as valor from  sis_reg_inscripcion WHERE ano="& AnoP &" AND periodo="& PeriodoP &" AND Acepto='SI' ", Session("Conn")
    if rst.recordcount > 0 then
       response.write rst("valor")
    else
       response.write "0"
    end if
    
    set rst=nothing
    
    %>
    <br>
      Alumnos Inscritos 07CS: 
    
    <%
    
    set rst = createobject("adodb.recordset")
    rst.cursortype = 1
    
    rst.open "select count(a.codcli) as valor from sis_reg_inscripcion a, mt_alumno b where a.fechaacepto>='"& FechaP &"' and acepto ='SI' and a.codcli=b.codcli and b.codcarpr='07CS'", Session("Conn")
    if rst.recordcount > 0 then
       response.write rst("valor")
    else
       response.write "0"
    end if
    
    set rst=nothing
    
    %>
    
    
    <br>
    </body>
<%else%>
    <body>
    
    Contador del Sitio: <% response.write (application("auxactivos"))%><br>
    
    Alumnos Inscritos en portal: 
    <% 
    set rst = createobject("adodb.recordset")
    rst.cursortype = 1
    
    rst.open "select COUNT(distinct codcli)valor from SIS_REG_INSCRIPCION WHERE fecha>=(select MIN(fecha_ini) from ra_calendario_tr WHERE ano=(SELECT dbo.Fn_ValorParame('ano')Parame) AND periodo=(SELECT dbo.Fn_ValorParame('periodo')Parame))", Session("Conn")
    if rst.recordcount > 0 then
       response.write rst("valor")
    else
       response.write "0"
    end if
    
    set rst=nothing
	%>
	<br>Alumnos Inscritos en total: 
    <% 
    set rst = createobject("adodb.recordset")
    rst.cursortype = 1
    
    rst.open "select COUNT(distinct codcli)valor from ra_nota where ano=(SELECT dbo.Fn_ValorParame('ano')Parame)", Session("Conn")
    if rst.recordcount > 0 then
       response.write rst("valor")
    else
       response.write "0"
    end if
    
    set rst=nothing
    %>
    
        <br>Solicitudes Realizadas: 
    
    <%
    set rst = createobject("adodb.recordset")
    rst.cursortype = 1
    
    rst.open "select COUNT(distinct codcli)valor from sis_reg_solicitud WHERE fecha>=(select MIN(fecha_ini) from ra_calendario_tr WHERE ano=(SELECT dbo.Fn_ValorParame('ano')Parame) AND periodo=(SELECT dbo.Fn_ValorParame('periodo')Parame))",Session("Conn")
    if rst.recordcount > 0 then
       response.write rst("valor")
    else
       response.write "0"
    end if
    
    set rst=nothing
    
    %>
    <br>
    </body>

<%end if%>
</html>
<!--#INCLUDE file="include/desconexion.inc" -->
