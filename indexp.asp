<%@ Language=VBScript %>
<!-- #Include file="adovbs.inc" -->
<!-- #Include file="funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='http://10.0.1.91/intranet.asp';</script>"
if session("acesso")>2 then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
'accesso func 1 prof 2
if session("a100")="N" or session("a100")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Informações do Professor</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
</head>
<body>
<%
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
'sqla="SELECT dc_carga.CURSO FROM dc_carga GROUP BY dc_carga.CURSO;"
'rs.Open sql1, ,adOpenStatic, adLockReadOnly
chapa=session("usuariomaster")
%>
<!-- -->
<table border="0" cellpadding="3" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse" >
<tr><td valign=top style="border-right:3px double silver;border-bottom:3px double silver" width=150 height=600>
<!-- -->
<p style="margin-top:0;margin-bottom:0" class=titulo><%=session("usuarioname")%></p>
<hr>
<p style="margin-top:0;margin-bottom:5"><a href="academic/disponibilidade.asp">
<img src="images/Clock.gif" width="16" height="16" border="0" alt="">Disponibilidade</a></p>

<p style="margin-top:0;margin-bottom:5"><a href="academic/aderencia.asp">
<img src="images/BookO.gif" width="16" height="16" border="0" alt="">Aderência</a></p>

<br><br>
<p style="margin-top:0;margin-bottom:5"><a href="academic/meusplanos.asp">
<img src="images/BookO.gif" width="16" height="16" border="0" alt="">Plano de Ensino</a></p>

<br><br>
<p style="margin-top:0;margin-bottom:5"><a href="academic/espelho.asp">
<img src="images/espelho.jpg" width="16" height="16" border="0" alt="">Marcação de Ponto</a></p>


<!-- -->
</td><td valign=top style="border-bottom:3px double silver" width=500>
<!-- -->
<%
hora=hour(now())
if hora<12 then 
	cumprimento="Bom dia"
elseif hora<18 then
	cumprimento="Boa tarde"
else
	cumprimento="Boa noite"
end if
sql="select sexo, apelido from corporerm.dbo.ppessoa p, corporerm.dbo.pfunc f where f.codpessoa=p.codigo and f.chapa='" & chapa & "' "
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs("sexo")="F" then suf1="a" else suf1=""
%>
<br>
<p align="center" class=titulo><%=cumprimento%>, professor<%=suf1%>&nbsp;<%=rs("apelido")%></p>
<br>
<br><br>
<b><font style="color:blue;font-size:12pt">Disponibilidade: </font></b>
<font style="color:black;font-size:11pt">informe aqui a sua disponibilidade de horários nos dias da semana.
<br><br>
<b><font style="color:blue;font-size:12pt">Aderência: </font></b>
<font style="color:black;font-size:11pt">pesquise os cursos e grades curriculares atuais e marque quais disciplinas você está apto a ministrar.
<br><br>
<b><font style="color:blue;font-size:12pt">Plano de Ensino: </font></b>
<font style="color:black;font-size:11pt">monte o conteúdo das disciplinas ministradas no semestre.
<br><br>
<b><font style="color:blue;font-size:12pt">Marcação de Ponto: </font></b>
<font style="color:black;font-size:11pt">verifique os horários marcados no ponto eletrônico. Caso haja falta de marcação procure o RH para a justificativa.
<br>


<!-- -->
</td><td valign="top">
<a href="javascript:window.history.back()"><img src="images/arrowb.gif" border="0" WIDTH="13" HEIGHT="10"></a>
</td></tr></table>
<table border="0" cellpadding="1" cellspacing="0" bordercolor="#000000" style="border-collapse: collapse">
<tr><td class=campo align="left" width=150>
<img src="images/logo_centro_universitario_unifieo_big.gif" width=110 alt="" border=0>

</td><td class="campop" align="right" width=510 valign=top>
Recursos Humanos
</td></tr></table>

<!-- -->
<%
'rs.close
'set rs=nothing
'conexao.close
'set conexao=nothing
%>

</body>
</html>