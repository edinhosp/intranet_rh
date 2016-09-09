<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" --><html>
<%
	'Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a64")="N" or session("a64")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<title>Protocolo de entrega</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open Application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao
sessao=session.sessionid

if request.form="" then
	sql="DELETE FROM ttbolsaprotocolo where sessao='" & sessao & "' "
	conexao.execute sql
	sql="INSERT INTO ttbolsaprotocolo ( sessao, chapa, NOME, nome_bolsista, matricula, tipocurso, curso, Campus, renovacao, validade, excecao, situacao, descricao, status, protocolo, id_lanc ) " & _
	"SELECT '" & sessao & "', b.chapa, f.NOME, b.nome_bolsista, b.matricula, b.tipocurso, b.curso, " & _
	"campus=case when tipocurso<>'Graduação' and tipocurso<>'Tecnológico' then 'Narciso' else case when b.curso='Direito' then 'Narciso' else 'Vila Yara' end end, l.renovacao, l.validade, l.excecao, l.situacao, s.descricao, " & _
	"status=case when l.situacao='M' then 'DEFERIDO' else case when l.situacao='I' then 'INDEFERIDO' else case when l.situacao='C' then 'CANCELADO' else case when l.situacao='S' then 'SUSPENSA (DP)' else 'VERIFICAR' end end end END, l.protocolo, l.id_lanc " & _
	"FROM bolsistas AS b, corporerm.dbo.pfunc AS f, bolsistas_lanc AS l, bolsistas_situacao s " & _
	"WHERE l.situacao=s.id_sit and b.chapa=f.chapa collate database_default and b.id_bolsa=l.id_bolsa AND l.protocolo=0 " 
	conexao.execute sql
sql="SELECT p.Campus, p.tipocurso, p.sessao, p.chapa, p.NOME, p.nome_bolsista, p.matricula, p.curso, p.renovacao, p.validade, p.excecao, p.situacao, p.descricao, p.status, p.protocolo, p.id_lanc, p.emitiu " & _
"FROM ttbolsaprotocolo AS p WHERE p.sessao='" & sessao & "' ORDER BY p.nome, p.nome_bolsista, p.Campus, p.tipocurso "
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs.recordcount=0 then
	response.write "<font size=""3"">Não existem protocolos a serem emitidos."
else
rs.movefirst
%>
<form name="protocolo" action="protocolo_entrega.asp" method="post">
<table border="1" cellpadding="0" cellspacing="1" style="border-collapse: collapse">
<tr>
	<td class=titulo>Campus/Tipo</td>
	<td class=titulo>Bolsista</td>
	<td class=titulo>Curso</td>
	<td class=titulo>Situação</td>
	<td class=titulo>Exceção</td>
	<td class=titulo>Emitir?</td>
</tr>
<%
vezes=1
do while not rs.eof
%>
<tr>
	<td class=campo><%=rs("campus")%>/<%=rs("tipocurso")%></td>
	<td class=campo><%=rs("nome_bolsista")%></td>
	<td class=campo><%=rs("curso")%></td>
	<td class=campo><%=rs("status")%></td>
	<td class=campo><%=rs("excecao")%></td>
	<td class=campo align="center">
		<input type="checkbox" name="emitir<%=vezes%>" value="ON" <%="checked"%> >
		<input type="hidden" name="id<%=vezes%>" value="<%=rs("id_lanc")%>">
	</td>
</tr>
<%
vezes=vezes+1
rs.movenext
loop
session("vezesprot")=vezes-1
%>
</table>
<input type="submit" value="Emitir protocolos" class="button" name="B1">
</form>
<%
end if 'rs.recordcount>0
else 'request.form
	vez=session("vezesprot")
	for a=1 to vez
		id=request.form("id" & a)
		emitir=request.form("emitir" & a)
		'response.write id & " " & tabela & " " & emitir & "<br>"
		if emitir="ON" then
			sql="UPDATE bolsistas_lanc SET protocolo=1 WHERE id_lanc=" & id 
			conexao.execute sql
			sql="UPDATE ttbolsaprotocolo SET emitiu=1 WHERE id_lanc=" & id & " and sessao='" & sessao & "' "
			conexao.execute sql
		end if
	next

	sql="SELECT p.Campus, p.tipocurso, p.sessao, p.chapa, p.NOME, p.nome_bolsista, p.matricula, p.curso, p.renovacao, p.validade, p.excecao, p.situacao, p.descricao, p.status, p.protocolo, p.id_lanc, p.emitiu " & _
	"FROM ttbolsaprotocolo AS p WHERE p.sessao='" & sessao & "' and emitiu=1 ORDER BY p.Campus, p.tipocurso, p.nome, p.nome_bolsista "
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	total=rs.recordcount

rs.movefirst:do while not rs.eof

'******* cabecalho **********
if rs("tipocurso")<>lasttipo or rs("campus")<>lastcampus then
if rs.absoluteposition>1 then 
	response.write "</table>"
	response.write "<br>"
	response.write "<table style='border-collapse: collapse' border=0 bordercolor=#000000 cellpadding=2 cellspacing=0 width=630>"
	response.write "<tr><td colspan=3 class=""campop"">Osasco,&nbsp;" & day(now()) & " de " & monthname(month(now())) & " de " & year(now()) &"</td></tr>"
	response.write "<tr><td class=""campop"" valign=top><br>Atenciosamente<br><br><br>______________________________<br>Recursos Humanos</td>"
	response.write "<td class=""campop"" valign=top style='border: 1px solid #000000'>Protocolo Secretaria Geral</td>"
	response.write "<td class=""campop"" valign=top style='border: 1px solid #000000'>Protocolo Tesouraria</td></tr></table>"
	response.write "<DIV style=""page-break-after:always""></DIV>"
end if
%>
<div align="right">
<table style="border-collapse: collapse" border="0" cellpadding="2" width="630" cellspacing="0">
<tr><td><img border="0" src="../images/logo_centro_universitario_unifieo_big.gif" width="225" height="50"></td></tr>
<tr><td class="campop" align="center"><br><b>PROTOCOLO DE ENTREGA DE PROCESSOS DE BOLSA DE ESTUDOS</td></tr>
<%
if rs("Campus")="Vila Yara" and (rs("tipocurso")="Graduação" or rs("tipocurso")="Tecnológico") then destinatario="Secretaria Geral-VY":copia="Tesouraria"
if rs("Campus")="Narciso" and rs("tipocurso")="Graduação" then destinatario="Secretaria Geral-NS":copia="Tesouraria"
if rs("Campus")="Narciso" and rs("tipocurso")<>"Graduação" then destinatario="Secretaria Pós-NS":copia="Tesouraria"
%>
	<tr><td class="campop">At.: <%=destinatario%><br>C/Cópia: <%=copia%></td></tr>
</table>
<table style="border-collapse: collapse" border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" width="630">
<tr><td class=titulo>Nome do funcionário</td>
	<td class=titulop>Bolsista</td>
	<td class=titulop>Matrícula</td>
	<td class=titulop>Situação</td>
	<td class=titulop>Ressalva</td></tr>
<%
end if
%>
<tr><td class="campor"><%=rs("nome")%></td>
	<td class="campop"><%=rs("nome_bolsista")%></td>
	<td class="campop"><%=rs("matricula")%></td>
	<td class="campop"><input type="text" name="txt1" class="form_input" size="10" value="<%=rs("status")%>" style="font-size:10pt"></td>
	<td class="campop"><%=rs("excecao")%></td></tr>
<%
lasttipo=rs("tipocurso")
lastcampus=rs("campus")
rs.movenext
loop
%>
</table>
<br>
<table style="border-collapse: collapse" border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" width="630">
<tr><td colspan=3 class="campop">Osasco,&nbsp;<%=day(now()) & " de " & monthname(month(now())) & " de " & year(now()) %></td></tr>
<tr><td class="campop" valign=top><br>Atenciosamente<br><br><br>______________________________<br>Recursos Humanos</td>
	<td class="campop" valign=top style="border: 1px solid #000000">Protocolo Secretaria Geral</td>
	<td class="campop" valign=top style="border: 1px solid #000000">Protocolo Tesouraria</td>
</tr>
</table>
<%	
end if 'request.form
%>

<%
conexao.close
set conexao=nothing
set rs=nothing
set rsc=nothing
%>
</body>
</html>