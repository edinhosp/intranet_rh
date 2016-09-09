<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.redirect "intranet.asp"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>RH Online</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="images/rho.png">
</head>

<%if session("usuariogrupo")="RH" then%>
<frameset cols="139,*,160" framespacing="0" border="1" color="#000000" frameborder="0">
	<frame name="frmMenu" id="frmMenu" scrolling="auto" target="frmMain" src="frmMenu.asp">
	<frame name="frmMain" id="frmMain" src="frmMain.asp" scrolling="auto">
	<frame marginwidth="1px" marginheight="1px" name="frmRodape" id="frmRodape" target="frmRodape" src="frmRodape.asp" frameborder="2" scrolling="auto">
	<noframes>
	<body>
	<p>Esta página usa quadros mas seu navegador não aceita quadros.</p>
	</body>
	</noframes>
</frameset>
<%else%>
<frameset cols="139,*" framespacing="0" border="0" frameborder="0">
	<frame name="frmMenu" id="frmMenu" scrolling="auto" target="frmMain" noresize src="frmMenu.asp">
	<frame name="frmMain" id="frmMain" src="frmMain.asp" scrolling="auto">
	<noframes>
	<body>
	<p>Esta página usa quadros mas seu navegador não aceita quadros.</p>
	</body>
	</noframes>
</frameset>
<%end if%>

</html>