<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a91")="N" or session("a91")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Livro de Ponto</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="VBScript">
	Sub datapt_onChange
		temp=document.form.datapt.value
		temp=weekday(temp)
		temp=weekdayname(temp)
		document.form.diasem.value=temp
	End Sub
</script>

</head>
<body style="margin-left:40px">
<form method="POST" action="pontob.asp" name="form">
<%
'******************************** inicio impressao
tamanho=640
%>
<!-- borda -->
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="<%=tamanho+10%>" height=1020>
<tr><td class=campo valign=top height=100%>
<!-- ponto -->
<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="<%=tamanho%>">
<tr>
	<td class="campop" align="left" valign="center"><b>PONTO DO PESSOAL DOCENTE DO CURSO DE <input class=form_ponto type="text" name="curso" size="35" value="digite aqui o nome do curso"></b></td>
	<td class="campop" align="right" valign="center"><input style="font-size:14pt;text-align:right;border-bottom:1px solid #000000" class=form_ponto type="text" name="pagina" size="1" value=""></td>
</tr>
<tr>
	<td class="campop" align="left">
	<input class=form_ponto type="text" name="diasem" size="11" value="<%=weekdayname(weekday(now))%>" disabled>
	<input style="font-size:10pt;text-align:left" class=form_ponto type="text" name="datapt" size="20" value="<%=formatdatetime(now,2)%>"></td>
	<td class="campop" align="right" nowrap><input style="font-size:10pt;text-align:right;border-bottom:1px solid #000000" class=form_ponto type="text" name="dial" size="1" value="">º dia letivo</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="<%=tamanho%>">
<tr>
	<td class="campop" align="left"><b>Período <select name="per" class=form_ponto style="font-size:10pt;text-align:left"><option>Matutino</option><option>Vespertino</option><option>Noturno</option></select></td>
	<td class="campop" align="center"><b>Grade <select name="grade" class=form_ponto style="font-size:10pt;text-align:left"><option>Semestral</option><option>Anual</option></select></td>
	<td class="campop" align="right"><b><input style="font-size:10pt;text-align:right" class=form_ponto type="text" name="perlet" size="10" value="<%=year(now)%>/1"></td>
</tr>

<table border="1" bordercolor="#000000" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="<%=tamanho%>">
<tr>
	<td class=titulo align="center">Nº</td>
	<td class=titulo align="center">Nome / Disciplina</td>
	<td class=titulo align="center">Assinatura</td>
	<td class=titulo align="center" colspan=2 width=50>Inicio</td>
	<td class=titulo align="center" colspan=2>Termino</td>
	<td class=fundo align="center">Faltas</td>
	<td class=fundo align="center">Observações</td>
</tr>	
<!--
<tr>
	<td class="campor" align="center" height=35>&nbsp;</td>
	<td class="campor"><b><input style="font-size:8pt;text-align:left" class=form_input type="text" name="professor" size="40" value="Nome do professor">
				</b><br><input style="font-size:7pt;text-align:left" class=form_input type="text" name="materia" size="40" value="MATÉRIA"></td>
	<td class="campop" width=150>&nbsp;</td>
<%for a=1 to 4 %>	
	<td class=campo align="center"><input style="font-size:8pt;text-align:center" class=form_input type="text" name="classe" size="1" value="1A"></td>
<%next%>
	<td class=campo>&nbsp;</td>
	<td class=campo>&nbsp;</td>
</tr>
-->
<%
linha=4
for l=5 to 20
%>
<tr>
	<td class="campor" align="center" height=35>&nbsp;</td>
	<td class="campor"><b><input style="font-size:8pt;text-align:left" class=form_input type="text" name="professor" size="40" value="">
				</b><br><input style="font-size:7pt;text-align:left" class=form_input type="text" name="materia" size="40" value=""></td>
	<td class="campop" width=150>&nbsp;</td>
<%for a=1 to 2 '14 to rs.fields.count-1 %>	
	<td class=campo colspan=2 align="center">&nbsp;&nbsp;&nbsp;:&nbsp;&nbsp;&nbsp;</td>
<%next%>
	<td class=campo>&nbsp;</td>
	<td class=campo>&nbsp;</td>
</tr>
<%
linha=linha+1
next
%>
</table>
<!-- ponto -->

<!-- borda -->
</td></tr>
<tr><td class="campor" valign=top height=30>
Diretor do curso: <input style="font-size:9pt;text-align:left" class=form_ponto type="text" name="diretor" size="50" value="">
<br>Coordenador do curso: <input style="font-size:9pt;text-align:left" class=form_ponto type="text" name="coord" size="50" value="">
</td></tr>
</table>
</body>
</html>