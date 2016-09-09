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
<title>Autorização para Contratação de Funcionário</title>
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
id_candidato=request("codigo")	
	
sql="SELECT r.id_requisicao, r.descricao, r.funcao AS cfuncao, " & _
	"c.NOME AS nfuncao, r.secao AS csecao, s.DESCRICAO AS nsecao, " & _
	"r.motivo AS cmotivo, r.id_area, r.id_faixa, r.faixa, r.salario, r.tipo, " & _
	"ntipo=case when tipo='1' then 'Normal' else 'Estagiario' end, r.horario AS chorario, " & _
	"h.DESCRICAO AS nhorario, r.chapasubst, f.NOME AS s_nome, f.salario s_salario, f.jornadamensal/60 s_jornada, " & _
	"ca.id_candidato, ca.nome_candidato, processo, r.exp_cumpre " & _
	"FROM (((((rs_requisicao r LEFT JOIN corporerm.dbo.PFUNCAO c ON r.funcao=c.CODIGO collate database_default) " & _
	"LEFT JOIN corporerm.dbo.PSECAO s ON r.secao=s.CODIGO collate database_default) " & _
	"LEFT JOIN corporerm.dbo.AHORARIO h ON r.horario=h.CODIGO collate database_default) " & _
	"LEFT JOIN corporerm.dbo.PFUNC AS f ON r.chapasubst=f.CHAPA collate database_default) " & _
	"INNER JOIN rs_candidato ca ON r.id_requisicao=ca.id_requisicao) " & _
	"INNER JOIN rs_agenda ON ca.id_candidato = rs_agenda.id_candidato " & _
	"WHERE ca.id_candidato=" & id_candidato & " AND rs_agenda.processo='7' "

rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<br><br>
<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr>
		<td><img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="225" height="50" alt=""></td>
		<td><b><input type="text" value="AUTORIZAÇÃO PARA CONTRATAÇÃO DE FUNCIONÁRIO" size=50 class=form_input10 style="font-weight:bold;">
	</tr>
</table>
<br><br>

<table cellpadding="5" cellspacing="0" width="650" style="border:1px solid #000000">
    <tr><td class="campop">Entrevistado por: <input type="text" value="" size=50 class=form_input10></td></tr>
	<tr><td class="campop">Processo Seletivo: <input type="text" value="" size="50" class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Departamento Requisitante</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<%=rs("csecao")%>&nbsp;&nbsp;<b><%=rs("nsecao")%></b></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Nome do Candidato</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<b><%=rs("nome_candidato")%></b></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Cargo</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<%=rs("cfuncao")%>&nbsp;&nbsp;<b><%=rs("nfuncao")%></b></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Local de Trabalho</i></td></tr>
<%
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
%>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<b>Campus <%=campus%></b></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Horário de Trabalho</i></td>
	<td class="campop" width=35% style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	<i>Jornada Mensal:&nbsp;<input type="text" value="220" size="3" class=form_input10>&nbsp;horas</i></td>
	</tr>
	<tr><td colspan=2 class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<%=rs("chorario")%>&nbsp;&nbsp;<b><%=rs("nhorario")%></b></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td colspan=2 class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Salário</i></td>
	<td class="campop" width=15% style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	<i>Faixa:&nbsp;<input type="text" value="<%=rs("faixa")%>" size="3" class=form_input10></i></td>
	</tr>
<%
if rs("salario")="" or isnull(rs("salario")) then salario=0 else salario=cdbl(rs("salario"))
if rs("exp_cumpre")=-1 and rs("tipo")=1 then fator=0.95 else fator=1
salario_exp=cdbl(salario*fator)
if isnull(rs("s_salario")) then s_salario=0 else s_salario=rs("s_salario")
if isnull(rs("s_jornada")) then s_jornada=0 else s_jornada=rs("s_jornada")
%>
	<tr>
		<td class="campop" style="border-left:1px solid #000000" width=50%>
	inicial: <b><%=formatnumber(salario_exp,2)%></b></td>
		<td class="campop" style="border-right:1px solid #000000" width=50% colspan=2>
	após experiência: <b><%=formatnumber(salario,2)%></b></td>
	</tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Data de Admissão</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" name="admissao" class=form_input10></td></tr>
</table>

<table border="0" bordercolor=#000000 cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000">
	<i>Motivo da Admissão</i></td>
	<td class="campop" style="border-top:1px solid #000000;border-left: 1px solid;border-right:1px solid #000000" colspan=2>
	<i>Nome do Substituído</i></td>
	</tr>

	<tr><td class="campop" style="border-left:1px solid #000000" rowspan=4>
	<input type="radio" name="motivo" value="02" <%if rs("cmotivo")="02" then response.write "checked"%>> Substituição<br>
 	<input type="radio" name="motivo" value="03" <%if rs("cmotivo")="03" then response.write "checked"%>> Vaga Nova<br>
 	<input type="radio" name="motivo" value="04" <%if rs("cmotivo")="04" then response.write "checked"%>> Aumento de Quadro
	</td>
	<td class="campop" style="border-right:1px solid #000000;border-left: 1px solid" valign=top colspan=2>
	<%=rs("chapasubst")%>&nbsp;&nbsp;<b><%=rs("s_nome")%></b></td>
	</tr>
	<tr>
		<td class="campop" style="border-left: 1px solid;border-right:1px solid #000000;border-top: 1px solid" valign=top><i>Salário do substituido</td>
		<td class="campop" style="border-right: 1px solid;border-top: 1px solid" valign=top><i>Jornada do substituido</td>
	</tr>
	<tr>
		<td class="campop" style="border-lefT: 1px solid;border-right:1px solid #000000" valign=top>
		<%a=formatnumber(s_salario,2)%>
		<% 
		if cdbl(s_salario)>0 then variacao=((cdbl(rs("salario"))) / (cdbl(s_salario)))-1
		if variacao<0 then s_texto=" (Redução de " else s_texto=" (Aumento de "
		if variacao<>"" then variacao=" " & formatpercent(variacao,2) & ")"
		%> <%b=s_texto & variacao%>
		<input type="text" class="form_input10" value="<%=a & b%>" size="28"> 
		</td>
		<td class="campop" style="border-right:1px solid #000000" valign=top><%=s_jornada%> horas</td>
	</tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Entrevistado(a) pela Chefia:</i>&nbsp;<input type="text" value="" size=64 class=form_input10 style="border-bottom:1px solid #000000"></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Comentários:</i>&nbsp;<input type="text" value="" size=76 class=form_input10 style="border-bottom:1px solid #000000"></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="text" value="" size=88 class=form_input10 style="border-bottom:1px solid #000000"></td></tr>


	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<input type="radio" name="aprov1" value="A"> Aprovado &nbsp;
 	<input type="radio" name="aprov1" value="N"> Não aprovado &nbsp; &nbsp;
	<input type="text" value="" size=35 class=form_input10 style="border-bottom:1px solid #000000"></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Recursos Humanos:</td>
<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<i>Pró-Reitoria Administrativa:</td></tr>

	<tr><td class="campop" height=50 style="border-bottom:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	&nbsp;</td>
<td class="campop" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
	&nbsp;</td></tr>

</table>
<%for a=1 to 4%>
<br>
<%next%>
<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>