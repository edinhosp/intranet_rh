<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a45")="N" or session("a45")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Pontos importantes da Convenção Coletiva 2005</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
</head>
<body>
<!-- -->
<table><tr><td>
<!-- -->
<%
'dim conexao, rs, rs2
'set conexao=server.createobject ("ADODB.Connection")
'conexao.Open application("conexao")
'set rs=server.createobject ("ADODB.Recordset")
'Set rs.ActiveConnection = conexao
'sqla="SELECT dc_carga.CURSO FROM dc_carga GROUP BY dc_carga.CURSO;"
'rs.Open sql1, ,adOpenStatic, adLockReadOnly
%>
<!-- auxiliares -->
<table border="0" cellpadding="4" cellspacing="5" style="border-collapse: collapse" width=690>
<tr><td valign=top class=titulop colspan=2>Hora de trabalho de professor: 50 ou 60 minutos</td></tr>
<tr>
	<td valign=top class=campo width=50% style="border-right: 1px solid #000000">
	<p class="artigo">Devido à grande polêmica que se criou na rede, com relação à duração da hora de trabalho
	do professor, o CPP consultou o DRHU tendo obtido a seguinte resposta:
	<p class="artigo">"Em razão de várias consultas a respeito, lembramos que a duração de hora de trabalho docente,
	de acordo com o que dispõe o parágrafo 10 da Lei Complementar nº 836, de 30 de dezembro de 1997, é de 60
	minutos, dentro os quais, 50 minutos dedicados à tarefa de ministrar aulas.
	<p class="artigo">As Resoluções nº 6, de 28 de janeiro de 2005 e nº 11, de 11 de fevereiro de 2005, dispõem sobre
	a organização curricular do ensino fundamental e do ensino médio e estabelecem duração das aulas de 50 minutos.
	</td>
	<td valign=top class=campo width=50%>
	<p class="artigo">Cabe ao Diretor da Escola proceder, em conjunto com a equipe escolar, a compatibilização do 
	funcionamento dos turnos escolares com as horas de trabalho dos docentes, ficando mantido o limite máximo de 8 aulas
	de trabalho por dia, atendendo ao disposto no artigo 5º do decreto nº 39.931, de 30 de janeiro de 1995 (8 horas = 
	480 minutos)".
	<p class="artigo">	
	</td>
</tr>
<tr>
	<td colspan=2  style="border-bottom: 1px solid #000000"></td></tr>


	
</table>

<br>

<!-- professores -->

<%
'rs.close
'set rs=nothing
'conexao.close
'set conexao=nothing
%>
<!-- -->
</td><td valign="top">
<a href="javascript:window.history.back()"><img src="../images/arrowb.gif" border="0" WIDTH="13" HEIGHT="10"></a>
</td></tr></table>
<!-- -->
</body>
</html>