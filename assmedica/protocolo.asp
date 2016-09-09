<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	'if response.buffer=true then Response.buffer=true
	Response.buffer=true:	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a85")="N" or session("a85")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
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
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
sessao=session.sessionid

if request.form="" then
	sql="DELETE FROM ttamprotocolo where sessao='" & sessao & "' "
	conexao.execute sql
	sql="INSERT INTO ttamprotocolo ( sessao,chapa,Funcionario,Beneficiario,empresa,plano,codigo,compr,id,tabela,parentesco) " & _
	"SELECT '" & sessao & "',am.chapa, ab.NOME AS Funcionario, ab.NOME AS Beneficiario, am.empresa, am.plano, am.codigo, am.compr, am.id_mudanca,'tit','Titular' " & _
	"FROM assmed_mudanca am, assmed_beneficiario ab WHERE am.chapa = ab.CHAPA and am.compr=0 "
	conexao.execute sql
	sql="INSERT INTO ttamprotocolo ( sessao,chapa,Funcionario,Beneficiario,empresa,plano,codigo,compr,id,tabela,parentesco) " & _
	"SELECT '" & sessao & "', ad.chapa, ab.NOME AS Funcionario, ad.dependente AS Beneficiario, adm.empresa, adm.plano, adm.codigo, adm.compr, adm.id_mud,'dep',ad.parentesco " & _
	"FROM assmed_beneficiario  ab, assmed_dep ad, assmed_dep_mudanca adm " & _
	"WHERE ab.chapa=ad.chapa and ad.id_dep=adm.id_dep AND adm.compr=0 "
	conexao.execute sql
	sql="SELECT t.sessao, t.chapa, t.Funcionario, t.Beneficiario, t.empresa, t.plano, t.codigo, t.compr, t.id, t.tabela,t.parentesco " & _
	", f.codsecao, s.descricao " & _
	"FROM ttamprotocolo t, corporerm.dbo.pfunc f, corporerm.dbo.psecao s " & _
	"where t.chapa=f.chapa collate database_default and f.codsecao=s.codigo and sessao='" & sessao & "' " & _
	"ORDER BY Funcionario, Beneficiario "
	rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<form name="protocolo" action="protocolo.asp" method="post">
<table border="1" cellpadding="0" cellspacing="1" style="border-collapse: collapse">
	<tr>
		<td class=titulo>Beneficiário</td>
		<td class=titulo>Empresa</td>
		<td class=titulo>Plano</td>
		<td class=titulo>Código</td>
		<td class=titulo>Parentesco</td>
		<td class=titulo>Emitir?</td>
	</tr>
<%
if rs.recordcount>0 then
	rs.movefirst
	vezes=1
	do while not rs.eof
	if lastchapa=rs("chapa") then
	else
		response.write "<tr><td class=titulo colspan=6>" & rs("chapa") & " - " & rs("funcionario") & "</td></tr>"
	end if
%>
	<tr>
		<td class=campo><%=rs("beneficiario")%>&nbsp;</td>
		<td class=campo align="center"><%=rs("empresa")%>&nbsp;</td>
		<td class=campo><%=rs("plano")%>&nbsp;</td>
		<td class=campo><%=rs("codigo")%>&nbsp;</td>
		<td class=campo><%=rs("parentesco")%>&nbsp;</td>
		<td class=campo align="center">&nbsp;
			<input type="checkbox" name="emitir<%=vezes%>" value="ON" <%="checked"%> >
			<input type="hidden" name="id<%=vezes%>" value="<%=rs("id")%>">
			<input type="hidden" name="tabela<%=vezes%>" value="<%=rs("tabela")%>">
			</td>
	</tr>
<%
	lastchapa=rs("chapa")
	vezes=vezes+1
	rs.movenext
	loop
	session("vezesprot")=vezes-1
%>
</table>
<input type="submit" value="Emitir protocolos" class="button" name="B1">
<%
else
	response.write "<tr><td class=""campop"" colspan=6><b><font color=blue>Não existem protocolos a serem emitidos</td></tr>"
end if

%>
</form>
<%
	else 'request.form
		vez=session("vezesprot")
		for a=1 to vez
			id=request.form("id" & a)
			emitir=request.form("emitir" & a)
			tabela=request.form("tabela" & a)
			'response.write id & " " & tabela & " " & emitir & "<br>"
			if emitir="ON" then
				if tabela="tit" then
					sql="UPDATE assmed_mudanca SET compr = 1 WHERE id_mudanca=" & id 
					conexao.execute sql
				end if
				if tabela="dep" then
					sql="UPDATE assmed_dep_mudanca SET compr = 1 WHERE id_mud=" & id 
					conexao.execute sql
				end if
				sql="UPDATE ttamprotocolo SET compr = 1 WHERE id=" & id & " AND tabela='" & tabela & "'" & " and sessao='" & sessao & "' "
				conexao.execute sql
			end if
		next
		sql="SELECT t.sessao, t.chapa, t.Funcionario, t.Beneficiario, t.empresa, t.plano, t.codigo, t.compr, t.id, t.tabela,t.parentesco " & _
		", f.codsecao, s.descricao, f.codsituacao " & _
		"FROM ttamprotocolo t, corporerm.dbo.pfunc f, corporerm.dbo.psecao s " & _
		"where t.chapa=f.chapa collate database_default and f.codsecao=s.codigo and compr=1 and t.sessao='" & sessao & "' " & _
		"ORDER BY Funcionario, Beneficiario "

		rs.Open sql, ,adOpenStatic, adLockReadOnly
		total=rs.recordcount
		imprime=1
		meiapagina=1
		rs.movefirst
		do while not rs.eof
		if rs("empresa")="U" then operadora="Unimed Seguros"
		if rs("empresa")="V" then operadora="Metlife Odonto"
		situacao="Afastado"
		if rs("codsituacao")="A" or rs("codsituacao")="F" or rs("codsituacao")="Z" then situacao="Ativo"
		if rs("codsituacao")="D" then situacao="Demitido"
if lastchapa<>rs("chapa") then
	set rsc=server.createobject ("ADODB.Recordset")
	Set rsc.ActiveConnection = conexao
	sqlc="SELECT Count(ttamprotocolo.chapa) AS total FROM ttamprotocolo WHERE compr=1 and chapa='" & rs("chapa") & "' and sessao='" &  sessao & "' "
	rsc.Open sqlc, ,adOpenStatic, adLockReadOnly
	total=rsc("total")
	rsc.close
	imprime=1
%>
<!-- table recibo -->
<table border="1" cellpadding="5" width="690" cellspacing="0">
<tr><td class=titulo colspan="5" align="center"><font size="4">PROTOCOLO DE ENTREGA<br>ASSISTÊNCIA MÉDICA E ODONTOLÓGICA</font></td></tr>
<tr>
	<td class=campo colspan="4">Recebi nesta data os seguintes cartões de assistência médica e odontológica
	da Unimed Seguros e Metlife Odonto:</td>
</tr>
<tr>
	<td class=titulo>Nome</td>
	<td class=titulo>Parentesco</td>
	<td class=titulo>Operadora</td>
	<td class=titulo>Carteirinha</td>
</tr>
<%
	if meiapagina=0 then meiapagina=1 else meiapagina=0
else 'lastchapa
	imprime=imprime+1
end if 'lastchapa
%>
<tr>
	<td class=campo><%=rs("beneficiario")%></td>
	<td class=campo><%=rs("parentesco")%></td>
	<td class=campo><%=operadora%></td>
	<td class=campo align="center">&nbsp;<%=rs("codigo")%></td>
</tr>
<%
	lastchapa=rs("chapa"):lastnome=rs("funcionario")
	secao=rs("codsecao") & " - " & rs("descricao")
	rs.movenext
	if imprime=total then
%>
<tr>
	<td class=campo colspan="4">
	Declaro ter conhecimento de que a Unimed Seguros cobra os seguintes valores da contratante (FIEO):<br>
	- 2ª via de cartão de titular ou dependente: R$ 15,00<br>
	- novo Manual de orientação: R$ 10,00<br>
	e que caso venha a solicitar a reemissão de cartão ou manual repassarei estes valores à FIEO 
	através de desconto em folha de pagamento.
	<br>(&nbsp;&nbsp;&nbsp;) Recebi um Manual de orientação.</td>
</tr>
<tr>
	<td class="campop" colspan=4>
	<b>Tomo ciência nesta data que a assistência médica prevista na convenção coletiva da categoria passa de:</b>
	<table border="1" cellpadding="0" cellspacing="1" style="border-collapse: collapse" width=75%>	
	<tr><td class=fundop width=35%>Operadora</td><td class=fundop width=25%>Plano</td><td class=fundop width=15%>Valor gasto</td><td class=fundop width=25%>Valor reembolsado</td></tr>
<%
	sqlp1="select operadora, plano, valor=sum(valor), reembolso=sum(reembolso) from ( " & _
"select chapa, operadora, empresa, m.plano, parentesco='Titular', valor, reembolso " & _
"from assmed_mudanca m inner join assmed_planos p on m.empresa=p.codigo and m.plano=p.plano inner join assmed_empresa e on e.codigo=p.codigo " & _
"where chapa='" & lastchapa & "' and fvigencia='07/31/2010' and empresa in ('M','O') " & _
"union " & _
"select chapa, operadora, empresa, m.plano, parentesco, valor, " & _
"reembolso=case when reembolso>0 then (case parentesco when 'Esposo' then valor when 'Companheiro' then valor else reembolso end) else 0 end " & _
"from assmed_dep_mudanca m inner join assmed_dep d on d.id_dep=m.id_dep inner join assmed_planos p on m.empresa=p.codigo and m.plano=p.plano inner join assmed_empresa e on e.codigo=p.codigo " & _
"where chapa='" & lastchapa & "' and fvigencia='07/31/2010' and empresa in ('M','O') " & _
") z group by operadora, plano"
	rs2.Open sqlp1, ,adOpenStatic, adLockReadOnly
	do while not rs2.eof
%>
	<tr>
		<td class="campop"> <%=rs2("operadora")%></td>
		<td class="campop"> <%=rs2("plano")%></td>
		<td class="campop"> <%=formatnumber(rs2("valor"),2)%></td>
		<td class="campop"> <%=formatnumber(rs2("reembolso"),2)%></td>
	</tr>
<%
	rs2.movenext:loop
	rs2.close
%>
	</table>	
	<br><b>para a seguinte:</b>
	<table border="1" cellpadding="0" cellspacing="1" style="border-collapse: collapse" width=75%>	
	<tr><td class=fundop width=35%>Operadora</td><td class=fundop width=25%>Plano</td><td class=fundop width=15%>Valor gasto</td><td class=fundop width=25%>Valor reembolsado</td></tr>
<%
	sqlp2="select operadora, plano, valor=sum(valor), reembolso=sum(reembolso) from ( " & _
"select chapa, operadora, empresa, m.plano, parentesco='Titular', valor, reembolso " & _
"from assmed_mudanca m inner join assmed_planos p on m.empresa=p.codigo and m.plano=p.plano inner join assmed_empresa e on e.codigo=p.codigo " & _
"where chapa='" & lastchapa & "' and ivigencia='08/01/2010' and empresa in ('U','V') " & _
"union " & _
"select chapa, operadora, empresa, m.plano, parentesco, valor, " & _
"reembolso=case when reembolso>0 then (case parentesco when 'Esposo' then valor when 'Companheiro' then valor else reembolso end) else 0 end " & _
"from assmed_dep_mudanca m inner join assmed_dep d on d.id_dep=m.id_dep inner join assmed_planos p on m.empresa=p.codigo and m.plano=p.plano inner join assmed_empresa e on e.codigo=p.codigo " & _
"where chapa='" & lastchapa & "' and ivigencia='08/01/2010' and empresa in ('U','V') " & _
") z group by operadora, plano"
	rs2.Open sqlp2, ,adOpenStatic, adLockReadOnly
	do while not rs2.eof
%>
	<tr>
		<td class="campop"> <%=rs2("operadora")%></td>
		<td class="campop"> <%=rs2("plano")%></td>
		<td class="campop"> <%=formatnumber(rs2("valor"),2)%></td>
		<td class="campop"> <%=formatnumber(rs2("reembolso"),2)%></td>
	</tr>
<%
	rs2.movenext:loop
	rs2.close
%>
	</table>	
	<br><b>e de que
	<br>-Comprometo-me a devolver todas as carteirinhas do plano anterior;
	<br>-Tenho o prazo de até o dia 16/08/2010 para fazer retificações nos planos e dados dos dependentes.
	</b>
	
	
	
	
	</td>
</tr>  
<tr>
	<td class=campo colspan="4">
	Osasco,&nbsp;<%=day(now()) & " de " & monthname(month(now())) & " de " & year(now()) %><br><br><br>
	______________________________________________________<br>
	<%=lastchapa%> - <%=lastnome%><br><%=secao%>
	</td>
</tr>

	</table>
<!-- table recibo -->
<br>
<br>
Observação:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b><%=situacao%></b><br>
-Caso esta entrega de carteirinhas se refira à substituição de carteira anterior ou de mudança de
plano, favor devolver ou enviar <br>ao Recursos Humanos as carteirinhas anteriores para serem devolvidas
à empresa de saúde.



<%
		'response.write "<p style='margin-top:0; margin-bottom:0'><font size=1>Recursos Humanos - FIEO"
		response.write "<hr>"
		response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	end if

	loop

	end if 'request.form

conexao.close
set conexao=nothing
set rs=nothing
%>
</body>
</html>