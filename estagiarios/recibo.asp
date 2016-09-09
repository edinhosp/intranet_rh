<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a72")="N" or session("a72")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Demonstrativo de Pagamento</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao

if request.form("ver")="" then
ano=0:mes=0:per=0
%>
<form action="recibo.asp" method="post" name="form">
<p class=realce>Emissão de Demonstrativo de Pagamento</p>
Selecionar período: <select name="periodo" onchange="javascript:submit()">
<option value="0">Selecione um período</option>
<%
sql2="SELECT ff.ANOCOMP, ff.MESCOMP, ff.NROPERIODO " & _
"FROM corporerm.dbo.PFPERFF ff INNER JOIN corporerm.dbo.PFUNC f ON ff.CHAPA=f.CHAPA " & _
"WHERE f.CODTIPO='T' " & _
"GROUP BY ff.ANOCOMP, ff.MESCOMP, ff.NROPERIODO " & _
"ORDER BY ff.ANOCOMP desc, ff.MESCOMP desc, ff.NROPERIODO "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof
if rs2("mescomp")<10 then mes="0" & rs2("mescomp") else mes=rs2("mescomp")
if request.form("periodo")=rs2("anocomp") & mes & rs2("nroperiodo") then tempsel="selected" else tempsel=""
%>
	<option value="<%=rs2("anocomp") & mes & rs2("nroperiodo")%>" <%=tempsel%>><%=mes & "/" & rs2("anocomp") & " - Periodo: " & rs2("nroperiodo")%></option>
<%
rs2.movenext
loop
rs2.close
%>
</select>
<br>
Estagiário: <select name="chapa">
<option value="0">Todos</option>
<%
	vartemp=request.form("periodo")
	ano=left(vartemp,4)
	mes=mid(vartemp,5,2)
	per=mid(vartemp,7,2)
	if ano="" then ano=0
	if mes="" then mes=0
	if per="" then per=0
sql2="SELECT ff.CHAPA, f.NOME " & _
"FROM corporerm.dbo.PFFINANC ff, corporerm.dbo.PFUNC f, corporerm.dbo.PEVENTO e WHERE ff.CHAPA=f.CHAPA AND ff.CODEVENTO=e.CODIGO " & _
"AND ff.ANOCOMP=" & ano & " AND ff.MESCOMP=" & mes & " AND ff.NROPERIODO=" & per & " " & _
"AND e.PROVDESCBASE<>'B' " & _
"GROUP BY ff.CHAPA, f.NOME " & _
"ORDER BY f.NOME "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
'rs2.movefirst
do while not rs2.eof
%>
	<option value="<%=rs2("chapa")%>"><%=rs2("nome")%></option>
<%
rs2.movenext
loop
rs2.close
%>
</select>
<br>

<input type="submit" name="ver" value="Visualizar" class=button>
</form>
<%

else 'request.form
	vartemp=request.form("periodo")
	ano=left(vartemp,4)
	mes=cint(mid(vartemp,5,2))
	per=mid(vartemp,7,2)
	pagina=0
	inicio=1
	chapa=request.form("chapa")
	numero=0
if chapa="0" then
	sql2="SELECT ff.CHAPA, f.NOME " & _
	"FROM corporerm.dbo.PFFINANC ff, corporerm.dbo.PFUNC f, corporerm.dbo.PEVENTO e WHERE ff.CHAPA=f.CHAPA AND ff.CODEVENTO=e.CODIGO " & _
	"AND ff.ANOCOMP=" & ano & " AND ff.MESCOMP=" & mes & " AND ff.NROPERIODO=" & per & " " & _
	"AND e.PROVDESCBASE<>'B' " & _
	"GROUP BY ff.CHAPA, f.NOME " & _
	"ORDER BY f.NOME "
	rs2.Open sql2, ,adOpenStatic, adLockReadOnly
	rs2.movefirst
	do while not rs2.eof
	redim preserve chapai(numero)
	chapai(numero)=rs2("chapa")
	rs2.movenext
	numero=numero+1
	loop
	rs2.close
else
	redim preserve chapai(0)
	chapai(0)=chapa
end if
for a=0 to ubound(chapai)
	chapap=chapai(a)
	sql="SELECT f.CHAPA, f.NOME, c.NOME AS funcao, f.DATAADMISSAO, p.RUA, p.NUMERO, p.COMPLEMENTO, p.BAIRRO, p.ESTADO, p.CIDADE, p.CEP, f.PISPASEP, p.CPF, p.CARTIDENTIDADE AS RG, f.NRODEPIRRF AS DEPIR, f.NRODEPSALFAM AS DEPSF, f.CONTAPAGAMENTO AS CONTA, f.CODBANCOPAGTO AS BANCO, f.CODAGENCIAPAGTO AS AGENCIA, f.SALARIO, f.JORNADAMENSAL " & _
"FROM (corporerm.dbo.PFUNC f INNER JOIN corporerm.dbo.PFUNCAO c ON f.CODFUNCAO=c.CODIGO) INNER JOIN corporerm.dbo.PPESSOA p ON f.CODPESSOA=p.CODIGO " & _
"WHERE f.CHAPA='" & chapap & "' ORDER BY f.NOME "
	rs.Open sql, ,adOpenStatic, adLockReadOnly
%>
<table border="0" width="660" height="700">
<tr><td valign=top>

<table border="0" cellspacing="0" width="650" cellpadding="0" style="border-collapse: collapse">
<tr><td class=campo align="center"><font size=3><b>Demonstrativo de Pagamento Mensal</b></font></td></tr>
<tr><td class=campo align="center"><font size=2>FIEO - Fundação Instituto de Ensino para Osasco</font></td></tr>
<table>

<table border="1" bordercolor="#000000" cellspacing="0" width="650" cellpadding="3" style="border-collapse: collapse">
<tr>
	<td valign=top class="campop">Matrícula<br><%=rs("chapa")%></td>
	<td valign=top class="campop">Nome<br><b><%=rs("nome")%></b></td>
</tr>
</table>
<table border="" bordercolor="#000000" cellspacing="0" width="650" cellpadding="4" style="border-collapse: collapse">
<tr>
	<td valign=top class="campop">Função<br><%=rs("funcao")%></td>
	<td valign=top class="campop">Data Admissão<br><%=rs("dataadmissao")%></td>
	<td valign=top class="campop">Endereço<br><%=rs("rua") & " " & rs("numero") & " " & rs("complemento")%></td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellspacing="0" width="650" cellpadding="4" style="border-collapse: collapse">
<tr>
	<td valign=top class="campop">Bairro<br><%=rs("bairro")%></td>
	<td valign=top class="campop">Cidade<br><%=rs("cidade")%></td>
	<td valign=top class="campop">CEP<br><%=rs("cep")%></td>
	<td valign=top class="campop">UF<br><%=rs("estado")%></td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellspacing="0" width="650" cellpadding="4" style="border-collapse: collapse">
<tr>
	<td valign=top class="campop">PIS<br><%=rs("pispasep")%></td>
	<td valign=top class="campop">CPF<br><%=rs("cpf")%></td>
	<td valign=top class="campop">Identidade<br><%=rs("rg")%></td>
	<td valign=top class="campop">Data Crédito<br>&nbsp;</td>
	<td valign=top class="campop">Referência<br><%=ucase(monthname(mes)) & "/" & ano%></td>
	<td valign=top class="campop">Dep.Sal.Fam.<br><%=rs("depsf")%></td>
	<td valign=top class="campop">Dep.IRRF<br><%=rs("depir")%></td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellspacing="0" width="650" cellpadding="4" style="border-collapse: collapse">
<tr>
	<td valign=top class="campop">Composição do Salário</td>
	<td valign=top class="campop" colspan=3 align="center">Local do Pagamento</td>
</tr>
<tr>
	<td valign=top class="campop">Salário Hora<br>&nbsp;<%=formatnumber(cdbl(rs("salario"))/(rs("jornadamensal")/60),2)%></td>
	<td valign=top class="campop">Banco<br><%=rs("banco")%></td>
	<td valign=top class="campop">Agência<br><%=rs("agencia")%></td>
	<td valign=top class="campop">C/C<br><%=rs("conta")%></td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellspacing="0" width="650" cellpadding="4" style="border-collapse: collapse">
<tr>
	<td valign=top class="campop" colspan=5 align="center">Discriminação das Parcelas</td>
</tr>
<tr>
	<td valign=top class="campop" align="center">Evento</td>
	<td valign=top class="campop" align="center">Discriminação</td>
	<td valign=top class="campop" align="center">Ref.</td>
	<td valign=top class="campop" align="center">Proventos</td>
	<td valign=top class="campop" align="center">Descontos</td>
</tr>
<%
liquido=0
sql2="SELECT f.CHAPA, f.ANOCOMP, f.MESCOMP, f.NROPERIODO, f.CODEVENTO, e.DESCRICAO, f.REF, e.PROVDESCBASE, f.VALOR " & _
"FROM corporerm.dbo.PFFINANC f INNER JOIN corporerm.dbo.PEVENTO e ON f.CODEVENTO=e.CODIGO " & _
"WHERE f.CHAPA='" & chapap & "' AND f.ANOCOMP=" & ano & " AND f.MESCOMP=" & mes & " AND f.NROPERIODO=" & per & " " & _
"AND e.PROVDESCBASE<>'B' "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
rs2.movefirst
do while not rs2.eof
if rs2("provdescbase")="P" then
	proventos=formatnumber(rs2("valor"),2)
	descontos="&nbsp;"
	liquido=liquido+cdbl(rs2("valor"))
	tproventos=tproventos+cdbl(rs2("valor"))
else
	descontos=formatnumber(rs2("valor"),2)
	proventos="&nbsp;"
	liquido=liquido-cdbl(rs2("valor"))
	tdescontos=tdescontos+cdbl(rs2("valor"))
end if
%>
<tr>
	<td valign=top class="campop"><%=rs2("codevento")%></td>
	<td valign=top class="campop"><%=rs2("descricao")%></td>
	<td valign=top class="campop" align="center"><%=rs2("ref")%></td>
	<td valign=top class="campop" align="right"><%=proventos%>&nbsp;</td>
	<td valign=top class="campop" align="right"><%=descontos%>&nbsp;</td>
</tr>
<%
rs2.movenext
loop
rs2.close
%>
</table>
</td></tr></table>
<table border="1" bordercolor="#000000" cellspacing="0" width="650" cellpadding="4" style="border-collapse: collapse">
<tr>
	<td valign=top class="campop" width=100 rowspan=6>&nbsp;</td>
	<td valign=top class="campop">Base para FGTS</td>
	<td valign=top class="campop">FGTS do mês</td>
	<td valign=top class="campop">Total de Proventos</td>
</tr>
<tr>
	<td valign=top class="campop" align="right">0,00&nbsp;</td>
	<td valign=top class="campop" align="right">0,00&nbsp;</td>
	<td valign=top class="campop" align="right"><%=formatnumber(tproventos,2)%>&nbsp;</td>
</tr>
<tr>
	<td valign=top class="campop">Base Calc. IRRF</td>
	<td valign=top class="campop">Pensão Alim. Extra Folha</td>
	<td valign=top class="campop">Total de Descontos</td>
</tr>
<tr>
	<td valign=top class="campop" align="right"><%=formatnumber(tproventos,2)%>&nbsp;</td>
	<td valign=top class="campop" align="right">0,00&nbsp;</td>
	<td valign=top class="campop" align="right"><%=formatnumber(tdescontos,2)%>&nbsp;</td>
</tr>
<tr>
	<td valign=top class="campop">Sal. Contr. INSS</td>
	<td valign=top class="campop" valign="center" rowspan=2><br><b>Líquido a Receber =></td>
	<td valign=top class="campop" valign="center" align="right" rowspan=2><br><b><%=formatnumber(liquido,2)%>&nbsp;</b></td>
</tr>
<tr>
	<td valign=top class="campop" align="right">0,00&nbsp;</td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellspacing="0" width="650" cellpadding="4" style="border-collapse: collapse">
<tr>
	<td valign=top class="campop" width=100>Data<br>&nbsp;</td>
	<td valign=top class="campop">Assinatura<br>&nbsp;</td>
</tr>
</table>




<%
rs.close
if a<>ubound(chapai) then response.write "<DIV style=""page-break-after:always""></DIV>"
tproventos=0
tdescontos=0

next

end if 'request.form

set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>