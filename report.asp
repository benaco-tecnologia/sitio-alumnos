<%
' Get QS variables
rpttoview = request.querystring("rpt")
viewer = request.querystring("init")

'build full path for report

rpttoview = MID(request.ServerVariables("PATH_TRANSLATED"), 1, (LEN(request.ServerVariables("PATH_TRANSLATED"))-11)) & "\xtreme\" & rpttoview & ".rpt"

' build path to MDB



' Only create the Crystal Application Object on first time through
If Not IsObject ( session ("oApp")) Then
Set session ("oApp") = Server.CreateObject("Crystal.CRPE.Application")
End If

' Turn off all Error Message dialogs
Set oGlobalOptions = Session ("oApp").Options
oGlobalOptions.MorePrintEngineErrorMessages = 0

' Open the report
If IsObject(session("oRpt")) then
	Set session("oRpt") = nothing
End if   
Set session("oRpt") = session("oApp").OpenReport(rpttoview,1)

' Turn off sepecific report error messages
Set oRptOptions = Session("oRpt").Options
oRptOptions.MorePrintEngineErrorMessages = 0


' Opening the page engine will cause the data to be read
Set session("oPageEngine") = session("oRpt").PageEngine

' Now decide what viewer to create
Select Case viewer

	Case "java"
%>
<html>
<head>
<title>Seagate Crystal Smart Viewer for Java</title>
</head>
<body bgcolor=C6C6C6>
<SCRIPT LANGUAGE="JavaScript"><!--
 	var _ns3 = false;
 	var _ns4 = false;
 	//--></SCRIPT>
 	<COMMENT><SCRIPT LANGUAGE="JavaScript1.1"><!--
 	var _info = navigator.userAgent;
 	var _ns3 = (navigator.appName.indexOf("Netscape") >= 0 && _info.indexOf("Win16") < 0 && _info.indexOf("Mozilla/3") >= 0);
 	var _ns4 = (navigator.appName.indexOf("Netscape") >= 0 && _info.indexOf("Win16") < 0 && _info.indexOf("Mozilla/4") >= 0 );
 	//--></SCRIPT></COMMENT>
 		<SCRIPT LANGUAGE="JavaScript"><!--
 			if(_ns3==true)
 document.writeln( '<applet code=com.seagatesoftware.img.ReportViewer.ReportViewer 		 codebase="/viewer/JavaViewer" 		 id=ReportViewer width=100% height=95%  archive="/viewer/JavaViewer/ReportViewer.zip">' );
 			else if (_ns4 == true)
 document.writeln( '<applet code=com.seagatesoftware.img.ReportViewer.ReportViewer 		 codebase="/viewer/JavaViewer" 		 id=ReportViewer width=100% height=95%  archive="/viewer/JavaViewer/ReportViewer.jar">' );
 			else
 document.writeln( '<applet code=com.seagatesoftware.img.ReportViewer.ReportViewer 		 codebase="/viewer/JavaViewer" 		 id=ReportViewer width=100% height=95%  >' );
 		//--></SCRIPT>

 		<param name=ReportName value="rptserver.asp">
		<param name=HasGroupTree value=true>
		<param name=ShowGroupTree value=true>
		<param name=HasRefreshButton value=true>
		<param name=HasPrintButton value=true>
		<param name=HasExportButton value=true>
 		<param name=cabbase value="/viewer/JavaViewer/ReportViewer.cab">
		</applet>

</body>
</html>

<%
	Case "actx"
%>

<HTML>
<HEAD>
<TITLE>Seagate Crystal Smart Viewer for ActiveX</TITLE>
</HEAD>
<BODY BGCOLOR=C6C6C6 LANGUAGE=VBScript ONLOAD="Page_Initialize">

<OBJECT ID="CRViewer"
	CLASSID="CLSID:C4847596-972C-11D0-9567-00A0C9273C2A"
	WIDTH=100% HEIGHT=95%
	CODEBASE="/viewer/activeXViewer/activexviewer.cab#Version=2,2,4,36">
<PARAM NAME="EnableRefreshButton" VALUE=1>
<PARAM NAME="EnableGroupTree" VALUE=1>
<PARAM NAME="DisplayGroupTree" VALUE=1>
<PARAM NAME="EnablePrintButton" VALUE=1>
<PARAM NAME="EnableExportButton" VALUE=1>
<PARAM NAME="EnableDrillDown" VALUE=1>
<PARAM NAME="EnableSearchControl" VALUE=1>
<PARAM NAME="EnableAnimationControl" VALUE=1>
<PARAM NAME="EnableZoomControl" VALUE=1>
</OBJECT>

<SCRIPT LANGUAGE="VBScript">
<!--
Sub Page_Initialize
	On Error Resume Next
	Dim webBroker
	Set webBroker = CreateObject("WebReportBroker.WebReportBroker")
	if ScriptEngineMajorVersion < 2 then
		window.alert "IE 3.02 users on NT4 need to get the latest version of VBScript or install IE 4.01 SP1. IE 3.02 users on Win95 need DCOM95 and latest version of VBScript, or install IE 4.01 SP1. These files are available at Microsoft's web site."
		CRViewer.ReportName = Location.Protocol + "//" + Location.Host +"/scrreports/rptserver.asp"
	else
		Dim webSource
		Set webSource = CreateObject("WebReportSource.WebReportSource")
		webSource.ReportSource = webBroker
		webSource.URL = Location.Protocol + "//" + Location.Host + "/scrreports/rptserver.asp"
		webSource.PromptOnRefresh = True
		CRViewer.ReportSource = webSource
	end if
	CRViewer.ViewReport
End Sub
-->
</SCRIPT>

</BODY>
</HTML>

<%
	Case "html_frame"
		response.redirect "htmstart.asp"

	Case "html_page"

	response.redirect "rptserver.asp"



	end select

%>
