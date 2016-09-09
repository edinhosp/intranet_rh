<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Verificação de Números</title>
<meta name="viewport" content="width=device-width" />
<link rel="stylesheet" type="text/css" href="../diversos.css"
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body width="300px">
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

visualiza=0:erro=0

if request.form<>"" then
	chapa=request.form("chapa")
	ncheck=request.form("ncheck")
	checagem=request.form("checagem")

	sql="select f.chapa, ncheck0=DAY(p.DTNASCIMENTO), ncheck1=YEAR(f.dataadmissao), ncheck2=LEFT(intranet_rh.dbo.textopuro(p.cartidentidade,2),4), " & _
	"ncheck3=intranet_rh.dbo.primeironome(f.nome), ncheck4=LEFT(p.cpf,3) " & _
	"from corporerm.dbo.PFUNC f inner join corporerm.dbo.PPESSOA p on p.CODIGO=f.CODPESSOA " & _
	"where CODSITUACAO<>'D' and CODTIPO='N' and CODSINDICATO<>'03' " & _
	"and f.chapa='" & chapa & "' "
	rs.Open sql, ,adOpenStatic, adLockReadOnly
	nome1=rs("ncheck3")
	if ncheck="0" and cstr(rs("ncheck0"))=cstr(checagem) then visualiza=1
	if ncheck="1" and cstr(rs("ncheck1"))=cstr(checagem) then visualiza=1
	if ncheck="2" and cstr(rs("ncheck2"))=cstr(checagem) then visualiza=1
	if ncheck="3" and ucase(rs("ncheck3"))=ucase(checagem) then visualiza=1
	if ncheck="4" and cstr(rs("ncheck4"))=cstr(checagem) then visualiza=1
	if visualiza=0 then erro=1
	if erro=1 then msgerr="A informação digitada não está correta." else msgerr=""
'response.write "<br>1: " & ncheck
'response.write "<br>1: " & checagem
	rs.close
end if

dim check(4), mcheck(4)
check(0)="DiaAniv"	: mcheck(0)="Informe o dia do seu aniversário"
check(1)="AnoAdm"	: mcheck(1)="Informe o ano de entrada no Unifieo"
check(2)="4RG"		: mcheck(2)="Informe os 4 (quatro) primeiros números do seu R.G."
check(3)="Nome"		: mcheck(3)="Qual o seu primeiro nome?"
check(4)="3CPF"		: mcheck(4)="Informe os 3 (três) primeiros números do seu C.P.F."
randomize timer
ncheck0=rnd*5
ncheck1=int(ncheck0):if ncheck1=5 then ncheck=4
'response.write "<br>" & ncheck0
'response.write "<br>" & ncheck1
%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="322">
<tr><td class="grupo" colspan="2" height="25" align="center">Consulta aos números dos Cartões</td></tr>
<tr>
	<td><img src="logo_centro_universitario_unifieo.jpg" border="0"></td>
	<td><img src="bradesco.png" border="0"></td>
</tr>


<form method="POST" action="index.asp">
<input type="hidden" name="ncheck" value="<%=ncheck1%>">
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="322">
<tr>
	<td class=titulo width="80%">Informe o número de sua Chapa</td>
	<td class=titulo width="20%"><input type="text" name="chapa" size="5" maxsize="5" value="<%=request.form("chapa")%>"></td>
</tr>
<tr>
	<td class=titulo><%=mcheck(ncheck1)%></td>
	<td class=titulo><input type="text" name="checagem" size="5" value="<%request.form("checagem")%>" class=a></td>
</tr>
<tr>
	<td class=titulo colspan="2"><input type="submit" value="Visualizar os números" name="Gerar" class="button"></td>
</tr>
</table>
</form>
<p><font color="red"><b><%=msgerr%></b>
<font color="black">
<%
if visualiza=1 then
response.write "<br>"
sqlc="select nome='TITULAR (" & nome1 & ")', 'Carteira'=codigo, up from assmed_mudanca where chapa='" & chapa & "' and empresa='BS' " & _
"union all select NOME=intranet_rh.dbo.primeironome(d.NOME), 'Carteira'=m.codigo, up " & _
"from assmed_dep_mudanca m inner join corporerm.dbo.PFDEPEND d on d.CHAPA collate database_default=m.chapa and d.NRODEPEND=m.nrodepend " & _
"where m.chapa='" & chapa & "' and empresa='BS' "
rs.Open sqlc, ,adOpenStatic, adLockReadOnly
%>
<table border="1" cellpadding="2" cellspacing="2" style="border-collapse: collapse" width="322">
<tr><td class="titulop">Nome</td><td class="titulop">Nº Cartão</td><td class="titulop">Beneficiário</td></tr>
<%
do while not rs.eof
%>
<tr><td class="campop"><%=rs("nome")%></td><td class="campop"><%=rs("carteira")%></td><td class="campop"><%=rs("up")%></td></tr>
<%
rs.movenext
loop
rs.close
%>
<tr><td class="campol" colspan="3"><font size="3">
Prezado(a) Funcionário(a):<br>
Caso necessite utilizar o plano e ainda não tenha recebido os cartões, anote os seus números e leve juntamente um documento com foto no atendimento médico.
</td></tr>
</table>
<br>
Central de Atendimento Bradesco: 4004.2700 - Opção 2<br>
Endereço da Rede: <a href="http://www.bradescosaude.com.br" target="_blank">Clique aqui</a>
<br><br>
<a href="http://www.bradescosaude.com.br" target="_blank">
<img src="site.png" border="0">
</a>
<%end if%>

<%
set rs=nothing
conexao.close
set conexao=nothing
%> 
</body>
</html>