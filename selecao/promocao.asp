<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a53")="N" or session("a53")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
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

if request.form("B1")="" then
	sql="select p.chapa, p.nome, p.codnivelsal from corporerm.dbo.pfunc p where  p.codsituacao<>'D' and p.codsindicato<>'03' and p.codtipo='N' " & _
	"order by p.nome "
else
	frmchapa=request.form("chapa")
	sql="SELECT f.CHAPA, f.NOME as nomef, f.CODSECAO, s.DESCRICAO as secao, f.CODFUNCAO, " & _
	"c.NOME as funcao, f.CODHORARIO, h.DESCRICAO as horario, f.SALARIO, f.codnivelsal, f.jornadamensal " & _
	"FROM ((corporerm.dbo.PFUNC f INNER JOIN corporerm.dbo.PFUNCAO c ON f.CODFUNCAO=c.CODIGO) INNER JOIN corporerm.dbo.PSECAO s ON f.CODSECAO=s.CODIGO) INNER JOIN corporerm.dbo.AHORARIO h ON f.CODHORARIO=h.CODIGO " & _
	"WHERE f.CHAPA='" & frmchapa & "' "
end if
frmchapa=request.form("chapa")

set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
rs.Open sql, ,adOpenStatic, adLockReadOnly

if request.form("B1")="" then

if request.form("chapa")<>"" then
	sql2="select s.codigo, s.descricao from corporerm.dbo.psecao s, corporerm.dbo.pfunc f where f.codsecao=s.codigo and f.chapa='" & frmchapa & "' "
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	secaoc=rs2("codigo"):secaon=rs2("descricao")
	rs2.close
	sql2="select s.codigo, s.nome as descricao from corporerm.dbo.pfuncao s, corporerm.dbo.pfunc f where f.codfuncao=s.codigo and f.chapa='" & frmchapa & "' "
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	funcaoc=rs2("codigo"):funcaon=rs2("descricao")
	rs2.close
	sql2="select s.codigo, s.descricao from corporerm.dbo.ahorario s, corporerm.dbo.pfunc f where f.codhorario=s.codigo and f.chapa='" & frmchapa & "' "
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	horarioc=rs2("codigo"):horarion=rs2("descricao")
	rs2.close
	sql2="select salario from corporerm.dbo.pfunc f where f.chapa='" & frmchapa & "' "
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	salario=rs2("salario")
	rs2.close
end if
if request.form("secao")="" then secao=secaoc else secao=request.form("secao")
if request.form("funcao")="" then funcao=funcaoc else funcao=request.form("funcao")
if request.form("horario")="" then horario=horarioc else horario=request.form("horario")

%>
<form name="form" action="promocao.asp" method="post">
<table border="1" bordercorlor="#CCCCCC" cellpadding="4" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo colspan=3>Promoção de Funcionário</td>
</tr>
<tr>
	<td class=campo>Funcionário</td>
	<td class=campo><input type="text" name="chapa" size="5" maxlength="5" onchange="chapa1()" value="<%=request.form("chapa")%>"></td>
	<td class=campo><select name="nome" class=a onchange="nome1()">
		<option value="0"> Selecione o funcionário</option>
<%
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
<tr>
	<td class=campo>Seção atual</td>
	<td class=campo><%=secaoc%><input type="hidden" name="secao0" value="<%=secaoc%>"></td>
	<td class=campo><%=secaon%></td>
</tr>
<tr>
	<td class=campo>Seção transferida</td>
	<td class=campo><input type="text" name="secao" size="8" class=a onchange="secao1()" value="<%=secao%>"></td>
	<td class=campo><select name="nsecao" class=a onchange="secao2()">
	<option value="0">Selecione a nova seção</option>
<%
sql2="select codigo, descricao from corporerm.dbo.psecao order by descricao"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
if secao=rs2("codigo") then temps="selected" else temps=""
%>
	<option value="<%=rs2("codigo")%>" <%=temps%> > <%=rs2("descricao")%></option>
<%
rs2.movenext
loop:rs2.close
%>
	</select>
	</td>
</tr>
<tr>
	<td class=campo>Função atual</td>
	<td class=campo><%=funcaoc%><input type="hidden" name="funcao0" value="<%=funcaoc%>" ></td>
	<td class=campo><%=funcaon%></td>
</tr>
<tr>
	<td class=campo>Função na transferência</td>
	<td class=campo><input type="text" name="funcao" size="8" class=a onchange="funcao1()" value="<%=funcao%>"></td>
	<td class=campo><select name="nfuncao" class=a onchange="funcao2()">
	<option value="0">Selecione a nova função</option>
<%
sql2="select codigo, nome from corporerm.dbo.pfuncao order by nome"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
if funcao=rs2("codigo") then tempf="selected" else tempf=""
%>
	<option value="<%=rs2("codigo")%>" <%=tempf%> > <%=rs2("nome")%></option>
<%
rs2.movenext
loop
rs2.close
%>
	</select>
	</td>
</tr>
<tr>
	<td class=campo>Horário atual</td>
	<td class=campo><%=horarioc%><input type="hidden" name="horario0" value="<%=horarioc%>" ></td>
	<td class=campo><%=horarion%></td>
</tr>
<tr>
	<td class=campo>Horário na transferência</td>
	<td class=campo><input type="text" name="horario" size="8" class=a onchange="horario1()" value="<%=horario%>"></td>
	<td class=campo>
	<select name="nhorario" class=small onchange="horario2()">
	<option value="0">Selecione o novo horário</option>
<%
sql2="select codigo, descricao from corporerm.dbo.ahorario order by descricao"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst:do while not rs2.eof
if horario=rs2("codigo") then temph="selected" else temph=""
%>
	<option value="<%=rs2("codigo")%>" <%=temph%> > <%=rs2("descricao")%></option>
<%
rs2.movenext:loop
rs2.close

%>
	</select>
	</td>
</tr>

<tr>
	<td class=campo>Salário Atual</td>
	<td class=campo colspan=2>
	<input type="text" name="salario1" size="10" class=a value="<%=formatnumber(salario,2)%>">
	
	<input type="hidden" name="salario0" size="10" value="<%=salario%>"></td>
</tr>
<tr>
	<td class=campo>Salário na promoção</td>
	<td class=campo colspan=2><input type="text" name="salario2" size="10" class=a></td>
</tr>
	<tr>
		<td class=campo colspan=3>&nbsp;
		<input type="submit" value="Visualizar" class=button name="B1">
		</td>
	</tr>
</table>

</form>
<%
else
%>
<br><br>
<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr>
		<td><img src="../images/logo_centro_universitario_unifieo_big.jpg" border="0" width="225" height="50" alt=""></td>
		<td><b>AUTORIZAÇÃO PARA MUDANÇA DE FUNÇÃO E/OU SALÁRIO</b><td>
	</tr>
</table>
<br><br>
<table cellpadding="5" cellspacing="0" width="650" style="border:1px solid #000000">
    <tr><td class="campop">Entrevistado por: <input type="text" value="" size=50 class=form_input10></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Nome do Funcionário</i></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<%=rs("chapa")%>&nbsp;&nbsp;<b><%=rs("nomef")%></b></td></tr>
</table>

<%
sql2="select descricao from corporerm.dbo.psecao where codigo='" & request.form("secao") & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
secaot=rs2("descricao")
rs2.close
adepto=rs("codsecao") & " - " & rs("secao")
tdepto=request.form("secao") & " - " & secaot
if tdepto=adepto then tdepto="o mesmo"
%>
<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Departamento anterior: </i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<%=adepto%></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Departamento atual: </i></td>
	<td class="campop" style="border-right:1px solid #000000">
	<%=tdepto%></td></tr>
</table>

<%
select case left(rs("codsecao"),2)
	case "01"
		campus="NARCISO"
	case "03"
		campus="V. YARA"
	case "04"
		campus="JD. WILSON"
	case else
		campus=""
end select
campus="Campus " & campus
select case left(request.form("secao"),2)
	case "01"
		campust="NARCISO"
	case "03"
		campust="V. YARA"
	case "04"
		campust="JD. WILSON"
	case else
		campust=""
end select
campust="Campus " & campust
if campust=campus then campust="o mesmo"
%>
<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000" width=300>
	<i>Local de trabalho anterior: </i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<%=campus%>&nbsp;</td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Local de trabalho atual: </i></td>
	<td class="campop" style="border-right:1px solid #000000">
	<%=campust%>&nbsp;&nbsp;</td></tr>
</table>

<%
sql2="select nome from corporerm.dbo.pfuncao where codigo='" & request.form("funcao") & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
funcaot=rs2("nome")
rs2.close
afuncao=rs("codfuncao") & " - " & rs("funcao")
tfuncao=request.form("funcao") & " - " & funcaot
if tfuncao=afuncao then tfuncao="o mesmo"
%>
<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1pxpx solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Função/Cargo anterior: </i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<%=afuncao%></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Função/Cargo atual/proposto: </i></td>
	<td class="campop" style="border-right:1px solid #000000">
	<%=tfuncao%></td></tr>
</table>

<%
sql2="select descricao from corporerm.dbo.ahorario where codigo='" & request.form("horario") & "' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
horariot=rs2("descricao")
rs2.close
ahorario=rs("codhorario") & " - " & rs("horario")
thorario=request.form("horario") & " - " & horariot
if thorario=ahorario then thorario="o mesmo"
%>
<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Horário anterior: </i></td>
	<td class=campo style="border-top:1px solid #000000;border-right:1px solid #000000">
	<%=ahorario%></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Horário atual: </i></td>
	<td class=campo style="border-right:1px solid #000000">
	<%=thorario%></td></tr>
</table>

<%
asalario=formatnumber(cdbl(rs("salario")),2)
asalario=formatnumber(cdbl(request.form("salario1")),2)
tsalario=formatnumber(cdbl(request.form("salario2")),2)
perc=tsalario/asalario-1
perc=formatpercent(perc,2)
if cdbl(tsalario)=cdbl(asalario) then tsalario="o mesmo"
%>
<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000" width=250>
	<i>Salário anterior: </i></td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000">
	<%=asalario%></td></tr>
	<tr><td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Salário proposto: </i></td>
	<td class="campop" style="border-right:1px solid #000000">
	<b><%=tsalario%></b> (variação: <%=perc%>)</td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650">
	<tr><td class="campop" valign=top style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Data da mudança</i> <input type="text" value="" name="admissao" class=form_input10 style="border-bottom:1px solid #000000">
	</td>
	<td class="campop" style="border-top:1px solid #000000;border-right:1px solid #000000" width=75%>Motivo da proposta:<br>
	<input type="radio" name="motivo" value="1"> Promoção &nbsp;<Br>
 	<input type="radio" name="motivo" value="2"> Enquadramento &nbsp;<Br>
	<input type="radio" name="motivo" value="3"> Mérito &nbsp;<Br>
 	<input type="radio" name="motivo" value="4"> <input type="text" value="" size=30 class=form_input10>&nbsp;
	</td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650" height=60>
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Solicitado por:</i>&nbsp;<input type="text" value="" size=76 class=form_input10 style="border-bottom:1px solid #000000"></td></tr>
</table>
<table border="0" cellpadding="5" cellspacing="0" width="650" height=60>
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000">
	<i>Recursos Humanos:</i>&nbsp;<input type="text" value="" size=76 class=form_input10 ></td></tr>
</table>

<table border="0" cellpadding="5" cellspacing="0" width="650" height=60>
	<tr><td class="campop" style="border-top:1px solid #000000;border-left:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
	<i>Pró-Reitoria Administrativa:</i>&nbsp;<input type="text" value="" size=56 class=form_input10>
	</td></tr>
</table>
<%for a=1 to 4%>
<br>
<%next%>
<p align="center">Recursos Humanos</p>
<%
rs.close

end if 'request.form

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>