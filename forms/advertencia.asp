<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a76")="N" or session("a76")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Formulário para Transferência de Funcionário</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function nome1() { form.chapa.value=form.nome.value;form.submit() }
function chapa1() { form.nome.value=form.chapa.value;form.submit() }
function secao1() { form.nsecao.value=form.secao.value; }
function secao2() { form.secao.value=form.nsecao.value; }
function funcao1() { form.nfuncao.value=form.funcao.value; }
function funcao2() { form.funcao.value=form.nfuncao.value; }
function horario1() { form.nhorario.value=form.horario.value; }
function horario2() { form.horario.value=form.nhorario.value; }
--></script>
</head>
<body style="margin-left:20px">
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
sessao=session.sessionid

set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form("B1")="" then
frmchapa=request.form("chapa")
%>
<form name="form" action="advertencia.asp" method="post">
<table border="1" bordercorlor="#CCCCCC" cellpadding="4" cellspacing="0" style="border-collapse: collapse" width=500>
<tr>
	<td class=titulo colspan=3>Advertência/Suspensão de funcionário</td>
</tr>
<tr>
	<td class=campo>Funcionário</td>
	<td class=campo><input type="text" name="chapa" size="5" maxlength="5" onchange="chapa1()" value="<%=request.form("chapa")%>"></td>
	<td class=campo><select name="nome" class=a onchange="nome1()">
		<option value="0"> Selecione o funcionário</option>
<%
sql="select p.chapa, p.nome from corporerm.dbo.pfunc p where  p.codsituacao<>'D' and p.codsindicato in ('03','01') and p.codtipo='N' " & _
"order by p.nome "
rs.Open sql, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof
if frmchapa=rs("chapa") then tempc="selected" else tempc=""
%>
		<option value="<%=rs("chapa")%>" <%=tempc%>> <%=rs("nome")%></option>
<%
rs.movenext:loop
%>
	</select>
	</td>
</tr>
</table>
<%
if request.form("chapa")<>"" then achapa=request.form("chapa") else achapa="00000"
sqla="select a.codpessoa, a.nroanotacao, a.texto, a.dtanotacao, a.dtresolucao, a.tipo " & _
"from corporerm.dbo.panotac a, corporerm.dbo.pfunc f where f.codpessoa=a.codpessoa and a.tipo in (8,9) and f.chapa='" & achapa & "' "
rs2.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<table border="1" bordercorlor="#CCCCCC" cellpadding="4" cellspacing="0" style="border-collapse: collapse" width=600>
<tr>
	<td class=titulo>&nbsp;</td>
	<td class=titulo>Tipo</td>
	<td class=titulo>Dt.Anotação</td>
	<td class=titulo>Dt.Resolução</td>
	<td class=titulo>Descrição</td>
</tr>
<%
if rs2.recordcount>0 then
rs2.movefirst
do while not rs2.eof
tipo=""
if rs2("tipo")=8 then tipo="Advertência"
if rs2("tipo")=9 then tipo="Suspensão"
%>
<input type="hidden" name="codpessoa" value="<%=rs2("codpessoa")%>">
<tr>
	<td class=campo><input type="radio" name="anotacao" value="<%=rs2("nroanotacao")%>"></td>
	<td class=campo><%=tipo%></td>
	<td class=campo><%=rs2("dtanotacao")%></td>
	<td class=campo><%=rs2("dtresolucao")%></td>
	<td class=campo><%=rs2("texto")%></td>
</tr>
<%
rs2.movenext
loop
end if
rs2.close
%>
<tr>
	<td class=campo colspan=5>&nbsp;
<%if tipo="" then%>
	Para emissão de advertência ou suspensão disciplinar é necessário o cadastro da mesmo no RM Labore.
<%else%>
	<input type="submit" value="Visualizar" class=button name="B1">
<%end if%>
	</td>
</tr>
</table>

</form>
<%
else

frmchapa=request.form("chapa")
frmanotacao=request.form("anotacao")
sql="select f.chapa, f.nome, p.carteiratrab, p.seriecarttrab, a.tipo, a.texto, a.nroanotacao, a.dtanotacao, a.dtresolucao, s.descricao as secao " & _
"from corporerm.dbo.pfunc f, corporerm.dbo.ppessoa p, corporerm.dbo.panotac a, corporerm.dbo.psecao s " & _
"where f.codpessoa=p.codigo and a.codpessoa=f.codpessoa and f.codsecao=s.codigo " & _
"and a.nroanotacao=" & frmanotacao & " and f.chapa='" & frmchapa & "' "
rs.Open sql, ,adOpenStatic, adLockReadOnly
if rs("tipo")=8 then tipo="Advertência"
if rs("tipo")=9 then tipo="Suspensão"
%>
<br><br>
<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr>
		<td><img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="225" height="50" alt=""></td>
		<td><p style="font-size:18pt"><b><%=ucase(tipo)%> DISCIPLINAR</b><td>
	</tr>
</table>
<br><br>

<table border="0" cellpadding="5" cellspacing="0" width="650" bordercolor="#000000">
	<tr><td class="campop" style="border-left: 1px solid;border-right: 1px solid;border-top: 1px solid;font-size:12pt">
	<i>Nome do Empregador</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid;border-right:1px solid;border-bottom: 1px solid;font-size:12pt">
	FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</b></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650" bordercolor="#000000">
	<tr><td class="campop" style="border-left:1px solid;border-right:1px solid;font-size:12pt">
	<i>Nome do Empregado</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid;border-right:1px solid;border-bottom: 1px solid;font-size:12pt">
	<b><%=rs("nome")%></b>&nbsp;&nbsp;&nbsp;&nbsp;(<%=rs("chapa")%>)</td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650" bordercolor="#000000">
	<tr><td class="campop" style="border-left:1px solid;border-right:1px solid;font-size:12pt">
	<i>CTPS / Série</i></td>
	<td class="campop" style="border-right:1px solid;font-size:12pt">
	<i>Departamento</i></td></tr>
	<tr><td class="campop" style="border-right:1px solid;border-left: 1px solid;border-bottom: 1px solid;font-size:12pt">
	<%=rs("carteiratrab")%> / <%=rs("seriecarttrab")%></td>
	<td class="campop" style="border-right:1px solid;border-bottom: 1px solid;font-size:12pt">
	<%=rs("secao")%></td></tr>
</table>

<br>

<table border="0" cellpadding="5" cellspacing="0" width="650" bordercolor="#000000">
	<tr><td class="campop" style="border:1px solid;font-size:12pt;text-align:justify">
<%
if tipo="Suspensão" then
	retorno=rs("dtresolucao")
	suspensao=rs("dtanotacao")
	dias=retorno-suspensao
	if dias>1 then plural="s" else plural=""
	complemento=", por " & dias & " dia" & plural & " a partir desta data"
else
	complemento=""
end if
%>
Esta tem a finalidade de aplicar-lhe a pena de <%=tipo%> Disciplinar<%=complemento%>, em razão da seguinte ocorrência:	
<br>
<br>
<b><%=rs("texto")%></b>
<br>
<br>
Esclarecemos, ainda, que a repetição de procedimentos como este poderá ser considerada como ato faltoso, passível de
dispensa por justa causa.<br>
<%if tipo="Advertência" then%>
Para que não tenhamos, no futuro, de tomar as medidas que nos facultam a legislação vigente, solicitamos-lhe que observar
as normas reguladoras da relação de emprego.<br>
<%end if:if tipo="Suspensão" then%>
Ao reassumir suas funções em <b><%=rs("dtresolucao")%></b>, solicitamos-lhe observar as normas reguladoras da relação de emprego
para que não tenhamos, no futuro, de tomar as medidas que nos facultam a legislação vigente.
<%end if%>
Favor dar seu ciente na cópia desta.
	</td></tr>
</table>

<br>

<table border="0" cellpadding="5" cellspacing="0" width="650" bordercolor="#000000">
	<tr><td class="campop" style="border:1px solid;font-size:12pt" width=300>
	<br>Osasco, <%=rs("dtanotacao")%>
	<br><br><br><br>_______________________________
	<br>    Empregador
	</td>
	<td class="campop" style="border-top:1px solid;border-right:1px solid;border-bottom: 1px solid;font-size:12pt">
	<br>Ciente em: ______/______/______
	<br><br><br><br>_______________________________
	<br>    Empregado
	</td></tr>
</table>

<%
rs.close

end if 'request.form

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>