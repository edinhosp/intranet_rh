<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a92")="N" or session("a92")="" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Visualização de Funcionários</title>
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
dim total
setor=request("setor")

sqla="SELECT codpessoa, dataexame, codmemoobs from vexamespront WHERE codpessoa=" & chapa & " ORDER BY dataexame"

sqla="SELECT O.CHAPA, ats.NOME, ats.CODSITUACAO, ats.SALARIO, ats.ats, ats.DATAADMISSAO, C.NOME AS FUNCAO, S.DESCRICAO AS SETOR " & _
"FROM organograma_pessoas O, pfunc_ats ats, corporerm.dbo.PSECAO S, corporerm.dbo.PFUNCAO C " & _
"WHERE O.CHAPA=ats.CHAPA collate database_default AND ats.CODSECAO=S.CODIGO AND ats.CODFUNCAO=C.CODIGO " & _
"AND O.ORGANOGRAMA='" & setor & "' ORDER BY ats.NOME "
rs.Open sqla, ,adOpenStatic, adLockReadOnly
totals=0
totala=0
totalg=0
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=515>
<th class=titulo colspan=9>Funcionários alocados no Setor: <%=setor%></th>
<tr>
	<td class=titulor>Chapa</td>
	<td class=titulor>Nome</td>
	<td class=titulor>Admissão</td>
	<td class=titulor>Função</td>
	<td class=titulor>Setor</td>
	<td class=titulor>Sit.</td>
<%if session("a92")="T" then%>
	<td class=titulor>Salário</td>
	<td class=titulor>A.T.S.</td>
	<td class=titulor>Total</td>
<%end if%>
</tr>
<%
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
if isnull(rs("ats"))=true then ats=0 else ats=cdbl(rs("ats"))
total=cdbl(rs("salario")) + ats
totals=totals+cdbl(rs("salario"))
totala=totala+ats
totalg=totals+totala
%>
<tr>
	<td class="campor"><%=rs("chapa")%></td>
	<td class="campor"><%=rs("nome")%></td>
	<td class="campor"><%=rs("dataadmissao")%></td>
	<td class="campor"><%=rs("funcao")%></td>
	<td class="campor"><%=rs("setor")%></td>
	<td class="campor"><%=rs("codsituacao")%></td>
<%if session("a92")="T" then%>
	<td class="campor" align="right"><%=formatnumber(rs("salario"),2)%>&nbsp;</td>
	<td class="campor" align="right"><%=formatnumber(ats,2)%>&nbsp;</td>
	<td class="campor" align="right"><%=formatnumber(total,2)%>&nbsp;</td>
<%end if%>
</tr>
<%
rs.movenext
loop
else
	response.write "<tr><td class=campo colspan=3>Sem lançamentos cadastrados</td></tr>"
end if
%>
<%if session("a92")="T" then%>
<tr>
	<td class=titulor colspan=6>Total do setor</td>
	<td class="campor" align="right" style="border-top:2 solid #000000"><%=formatnumber(totals,2)%>&nbsp;</td>
	<td class="campor" align="right" style="border-top:2 solid #000000"><%=formatnumber(totala,2)%>&nbsp;</td>
	<td class="campor" align="right" style="border-top:2 solid #000000"><%=formatnumber(totalg,2)%>&nbsp;</td>
</tr>
<%end if%>
</table>

<%
rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>