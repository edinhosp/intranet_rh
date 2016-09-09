<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a66")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Relatório de Pagamento de Reembolso</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, rt(10), rd(10)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

if request.form("Gerar")="" then 
%>
<p class=titulo>Geração de relatório de pagamento de Reembolso de Mensalidades
<form method="POST" action="reembolsorel.asp" name="form">
  <p>Reembolso a: <select size="1" name="sindicato" class=a onChange="javascript:submit()">
  <option value="01" <%if request.form("sindicato")="01" then response.write "selected"%> >Funcionários</option>
  <option value="03" <%if request.form("sindicato")="03" then response.write "selected"%> >Professores</option>
  </select>
  <br>
  Data de Pagamento: <select size="1" name="datapagamento" class=a>
<%
if request.form("sindicato")="" then sindicato="01" else sindicato=request.form("sindicato")
sql="SELECT b.data_pagamento FROM bolsistas_reembolso b INNER JOIN corporerm.dbo.pfunc f ON b.chapa=f.CHAPA collate database_default WHERE "
if sindicato="01" then sql=sql & " f.CODSINDICATO<>'03' " else sql=sql & " f.CODSINDICATO='03' "
sql=sql & "GROUP BY b.data_pagamento order by b.data_pagamento desc "
rsc.Open sql, ,adOpenStatic, adLockReadOnly
rsc.movefirst
do while not rsc.eof
%>  
		<option value="<%=rsc("data_pagamento")%>"><%=rsc("data_pagamento")%></option>
<%
rsc.movenext
loop
rsc.close
%>
		</select></p>
<p><input type="submit" value="Visualizar relatório" name="Gerar" class="button"></p>
</form>
<%
else
	sindicato=request.form("sindicato")
	datapagamento=cdate(request.form("datapagamento"))
	sessao=session.sessionid
	if sindicato="01" then 
		tipo="Funcionários" 
		sql2="<>'03' "
	else
		tipo="Professores"
		sql2="='03' "
	end if
%>
<table border="0" cellpadding="2" width="1000" cellspacing="0" style="border-collapse: collapse">
<tr><td class="campop">Para: DEPARTAMENTO FINANCEIRO</td></tr>
<tr><td class="campop">De  : PRÓ-REITORIA ADMINISTRATIVA</td></tr>
<tr><td class="campop">Ref. Planilha de Reembolso a <%=tipo%>-bolsistas - <%=ucase(monthname(month(datapagamento)))%>/<%=year(datapagamento)%>
<tr><td class="campor">&nbsp;</td></tr>
<tr><td class=campo>De conformidade com o que estabelece a Comunicação Interna nº 11/00, da Pró-Reitoria Administrativa, datada de
30/05/2000, autorizo o reembolso, através de crédito em conta corrente, dos valores expressos nos recibos anexos e adiante
relacionados, obedecendo a porcentagem indicada, aos <%=tipo%>-bolsistas abaixo relacionados:
<tr><td class="campor">&nbsp;</td></tr>
</td></tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="2" width="1000" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo align="center">Funcionários</td>
	<td class=titulo align="center">Mês/Ano</td>
	<td class=titulor align="center">Valor Mens.<br>sem multa</td>
	<td class=titulo align="center">% Deferida</td>
	<td class=titulor align="center">Valor<br>Reembolso R$</td>
	<td class=titulo align="center">Total R$</td>
	<td class=titulo align="center">Agência</td>
	<td class=titulor align="center">Conta<br>Corrente</td>
	<td class=titulo align="center">Código</td>
	<td class=titulo align="center">Centro de<br>Custo</td>
	<td class=titulo align="center">Departamento</td>
</tr>

<%
linha=2
SQL="SELECT BR.chapa, F.NOME, Sum(BR.reembolso) AS total, BR.porcentagem, S.DESCRICAO, F.CODAGENCIAPAGTO as agencia, F.CONTAPAGAMENTO as conta, f.codsecao " & _
"FROM (bolsistas_reembolso AS BR INNER JOIN corporerm.dbo.PFUNC AS F ON BR.chapa = F.CHAPA collate database_default) INNER JOIN corporerm.dbo.PSECAO AS S ON F.CODSECAO = S.CODIGO " & _
"GROUP BY BR.chapa, F.NOME, BR.porcentagem, BR.data_pagamento, F.CODSINDICATO, S.DESCRICAO, F.CODAGENCIAPAGTO, F.CONTAPAGAMENTO, f.codsecao " & _
"HAVING BR.data_pagamento='" & dtaccess(datapagamento) & "' AND F.CODSINDICATO" & sql2 & " " & _
"ORDER BY F.NOME "
rs.Open sql, ,adOpenStatic, adLockReadOnly
totalam=0:totalrb=0
rs.movefirst
do while not rs.eof
if linha>75 then
	pagina=pagina+1
	response.write "</table>"
	response.write "<br>"
	response.write "<p><font size='1'>Recursos Humanos - FIEO    -    Página " & pagina & "    -    " & now() & "</font></p>"
	response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
	linha=2
end if
'totalam=totalam+cdbl(rs("rateio"))
agencia=left(rs("agencia"),len(rs("agencia"))-1) & "-" & right(rs("agencia"),1)
conta=left(rs("conta"),len(rs("conta"))-1) & "-" & right(rs("conta"),1)
select case mid(rs("codsecao"),4,1)
	case "3"
		codigo="401"
	case "2"
		codigo="516"
	case "1"
		codigo="631"
	case else	
		codigo=""
end select
sql3="select mes_base, mensalidade, reembolso from bolsistas_reembolso where chapa='" & rs("chapa") & "' and data_pagamento='" & dtaccess(datapagamento) & "' "
rsc.Open sql3, ,adOpenStatic, adLockReadOnly
rsc.movefirst
%>
<tr>
	<td class=campo ><%=rs("nome")%></td>
	<td class=campo align="center">
	<% rsc.movefirst:do while not rsc.eof
	if rsc.recordcount>1 and rsc.absoluteposition>1 then response.write "<br>"
	response.write numzero(month(rsc("mes_base")),2) & "/" & year(rsc("mes_base"))
	rsc.movenext:loop
	%></td>
	<td class=campo align="right">
	<% rsc.movefirst:do while not rsc.eof
	if rsc.recordcount>1 and rsc.absoluteposition>1 then response.write "<br>"
	response.write formatnumber(rsc("mensalidade"),2) & "&nbsp;"
	rsc.movenext:loop
	%></td>
	<td class=campo align="center"><%=rs("porcentagem")%></td>
	<td class=campo align="right">
	<% rsc.movefirst:do while not rsc.eof
	if rsc.recordcount>1 and rsc.absoluteposition>1 then response.write "<br>"
	response.write formatnumber(rsc("reembolso"),2) & "&nbsp;"
	rsc.movenext:loop
	%></td>
	<td class=campo align="right"><%=formatnumber(rs("total"),2)%>&nbsp;</td>
	<td class=campo align="center"><%=agencia%></td>
	<td class=campo align="center"><%=conta%></td>
	<td class=campo align="center"><%=codigo%></td>
	<td class=campo align="center"><%=rs("codsecao")%></td>
	<td class="campor" ><%=rs("descricao")%></td>
</tr>
<%
rsc.close
totalrb=totalrb+cdbl(rs("total"))
linha=linha+1
rs.movenext
loop
rs.close
%>
<tr>
	<td class=titulo colspan=5>&nbsp;</td>
	<td class=titulo align="right"><%=formatnumber(totalrb,2)%>&nbsp;</td>
	<td class=titulo colspan=5>&nbsp;</td>
</tr>
</table>
<p style="margin-top: 0; margin-bottom: 0">(<%=extenso2(totalrb)%>)

<table border="0" bordercolor="#000000" cellpadding="2" width="1000" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=campo width=600></td>
	<td class="campop" align="center">
	<br>Osasco, <%=day(datapagamento)%> de <%=monthname(month(datapagamento))%> de <%=year(datapagamento)%>
	<br><br><br><br>
	<b>LUIZ FERNANDO DA COSTA E SILVA
	<br>Pró-Reitor Administrativo
	</td>
</tr>
</table>

<%
linha=linha+1
pagina=pagina+1
end if
%>
</body>
</html>
<%
set rs=nothing
set rsc=nothing
conexao.close
set conexao=nothing
%>