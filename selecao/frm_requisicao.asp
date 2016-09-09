<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a68")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Formulário para Abertura de Vaga</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body style="margin-left:20px">
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rsc=server.createobject ("ADODB.Recordset")
Set rsc.ActiveConnection = conexao

sessao=session.sessionid
id_requisicao=request("codigo")	
	
sql="SELECT r.id_requisicao, r.descricao, r.funcao AS cfuncao, c.NOME AS nfuncao, r.secao AS csecao, s.DESCRICAO AS nsecao, " & _
"r.requisitante AS creq, f.NOME AS nreq, r.motivo AS cmotivo, r.id_area, r.id_faixa, r.faixa, r.salario, r.tipo, " & _
"ntipo=case when tipo='1' then 'Normal' else 'Estagiário' end, r.horario AS chorario, r.exp_cumpre, h.DESCRICAO AS nhorario, " & _
"r.escolaridade AS cinstrucao, i.DESCRICAO AS ninstrucao, r.idademin, r.idademax, r.experiencia, r.sexo, r.cursos, " & _
"r.deficiente, ndeficiente=case when deficiente='0' then 'Indiferente' else case when deficiente='1' then 'Não deficiente' else 'Deficiente' end end, " & _
"r.tp_def, r.outros, r.dt_abertura, r.dt_encerramento, r.qt_vagas, r.chapasubst, f1.NOME AS nchapasubst " & _
"FROM ((((((rs_requisicao r LEFT JOIN corporerm.dbo.PFUNCAO c ON r.funcao=c.CODIGO collate database_default) " & _
"LEFT JOIN corporerm.dbo.PSECAO s ON r.secao=s.CODIGO collate database_default) " & _
"LEFT JOIN corporerm.dbo.PCODINSTRUCAO i ON r.escolaridade=i.CODCLIENTE collate database_default) " & _
"LEFT JOIN corporerm.dbo.PSUBSTCHEFE ch ON r.secao=ch.CODSECAO collate database_default) " & _
"LEFT JOIN corporerm.dbo.PFUNC f ON ch.CHAPASUBST=f.CHAPA) " & _
"LEFT JOIN corporerm.dbo.AHORARIO h ON r.horario=h.CODIGO collate database_default) " & _
"LEFT JOIN corporerm.dbo.PFUNC f1 ON r.chapasubst=f1.CHAPA collate database_default WHERE r.id_requisicao=" & id_requisicao & _
"and (ch.datafim is null or ch.datafim>getdate() )"

	rs.Open sql, ,adOpenStatic, adLockReadOnly

select case left(rs("csecao"),2)
	case "01"
		campus="NARCISO"
	case "03"
		campus="V. YARA"
	case "04"
		campus="JD. WILSON"
	case else
		campus=""
end select
select case rs("cmotivo")
	case "02"
		nmotivo="Substituição"
	case "03"
		nmotivo="Vaga Nova"
	case "04"
		nmotivo="Aumento de Quadro"
	case else
		nmotivo=""
end select
select case rs("sexo")
	case "F"
		nsexo="Feminino"
	case "M"
		nsexo="Masculino"
	case else
		nsexo="Indiferente"
end select

if rs("salario")="" or isnull(rs("salario")) then salario=0 else salario=cdbl(rs("salario"))
if rs("exp_cumpre")=1 and rs("tipo")=1 then fator=0.95 else fator=1
salario_exp=cdbl(salario*fator)

%>
<br><br>
<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr>
	<td><img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="225" height="50" alt=""></td>
	<td><b><input type="text" value="FORMULÁRIO PARA ABERTURA DE VAGA" size=50 class=form_input10 style="font-weight:bold;"></b><td>
</tr>
</table>
<br><br>
<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>FUNÇÃO/CARGO</i></td></tr>
<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<%=rs("cfuncao")%>&nbsp;&nbsp;<b><%=rs("nfuncao")%></b></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>SEÇÃO/DEPTº</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<i>CAMPUS</i></td></tr>
<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<%=rs("csecao")%>&nbsp;&nbsp;<b><%=rs("nsecao")%></b></td>
	<td class="campop" style="border-right:1px solid #000000">
	<b>Campus <%=campus%></b></td>	</tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>REQUISITANTE</i></td></tr>
<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<%=rs("creq")%>&nbsp;&nbsp;<b><%=rs("nreq")%></b></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;;border-right:1px solid #000000">
	<i>Motivo da Contratação</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<i>se substituição, informar nome do Substituído</i></td>
	</tr>

<tr><td class="campop" style="border-left:1px solid #000000;;border-right:1px solid #000000">
	<%=rs("cmotivo")%>&nbsp;&nbsp;<%=nmotivo%></td>
	<td class="campop" style="border-right:1px solid #000000" valign=top>
	<%if rs("chapasubst")<>"0" then response.write rs("chapasubst")%>&nbsp;&nbsp;<b><%=rs("nchapasubst")%></b></td>
	</tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Tipo do Contratado</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<i>Faixa Salarial</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000" align="center">
	<i>Salário Contratação / Efetivação</i></td></tr>
<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<%=rs("ntipo")%>&nbsp;&nbsp;</td>
	<td class="campop" style="border-right:1px solid #000000">
	<b><%=rs("faixa")%></b>&nbsp;</td>
	<td class="campop" style="border-right:1px solid #000000" align="center">
	<b><%=formatnumber(salario_exp,2)%>&nbsp;/&nbsp;<%=formatnumber(salario,2)%></b></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Horário de Trabalho</i></td></tr>
<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<%=rs("chorario")%>&nbsp;&nbsp;<b><%=rs("nhorario")%></b></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;;border-right:1px solid #000000">
	<i>Instrução / Formação</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<i>Experiência Mínima</i></td>
	</tr>

<tr><td class="campop" style="border-left:1px solid #000000;;border-right:1px solid #000000">
	<%=rs("cinstrucao")%>&nbsp;&nbsp;<%=rs("ninstrucao")%></td>
	<td class="campop" style="border-right:1px solid #000000" valign=top>
	&nbsp;<%=rs("experiencia")%>&nbsp;&nbsp;anos</td>
	</tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Sexo</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<i>Idade</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<i>Deficiente / Tipo def.</i></td></tr>

<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	&nbsp;<%=nsexo%>&nbsp;</td>
	<td class="campop" style="border-right:1px solid #000000">
	mínima&nbsp;<%=rs("idademin")%>&nbsp;&nbsp;máxima&nbsp;<%=rs("idademax")%></td>
	<td class="campop" style="border-right:1px solid #000000">
	&nbsp;<%=rs("ndeficiente")%>&nbsp;/&nbsp;<%=rs("tp_def")%></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Curso Técnico e/ou Exigido</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<%=rs("cursos")%>&nbsp;</td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650" height=85>
<tr><td height=15 class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Outras Informações sobre o cargo</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000" valign=top>
	<%=rs("outros")%>&nbsp;</td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Data de Abertura</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<i>Data de Encerramento</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<i>Quant. Vagas disponíveis</i></td></tr>
<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	&nbsp;<%=rs("dt_abertura")%>&nbsp;</td>
	<td class="campop" style="border-right:1px solid #000000">
	&nbsp;<%=rs("dt_encerramento")%>&nbsp;</td>
	<td class="campop" style="border-right:1px solid #000000">
	&nbsp;<%=rs("qt_vagas")%>&nbsp;</td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>&nbsp;</i></td></tr>
<tr><td class="campop" style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	&nbsp;</td></tr>
</table>
<p style="margin-top: 0; margin-bottom: 0"><font size=1><%=rs("descricao")%></p>
<%for a=1 to 4%>
<br>
<%next%>
<p align="center">Recursos Humanos</p>
<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>