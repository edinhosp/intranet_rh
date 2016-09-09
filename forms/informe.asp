<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a75")="N" or session("a75")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Informe de Rendimentos - Funcionários</title>
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
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

if request.form("B1")="" then
if request.form("saida")="" then tiposaida="I" else tiposaida=request.form("saida")
%>
<form method="POST" action="informe.asp" name="form">
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo colspan=3>Seleções para emissão de Informe de Rendimentos</td>
</tr>
<tr>
	<td class="campoa">Nome do Beneficiário</td>
	<td class="campoa">Chapa: <input type="text" name="chapaform" SIZE="5" value="<%=anoform%>">
<!--
	<select size=1 name="anochapa">
<%
sqltemp="SELECT f.NOME, ff.CHAPA, Year(dtpagto) AS ANO FROM (SELECT CHAPA, DTPAGTO FROM PFFINANC UNION ALL SELECT CHAPA, DTPAGTO FROM PFFINANCCOMPL) ff INNER JOIN PFUNC f ON ff.CHAPA=f.CHAPA " & _
"WHERE ff.DTPAGTO Between #1/1/2003# And #12/31/2006# GROUP BY f.NOME, ff.CHAPA, Year(dtpagto) ORDER BY f.NOME, Year(dtpagto);"
sqltemp="SELECT f.NOME, f.CHAPA FROM corporerm.dbo.PFUNC F ORDER BY f.NOME;"
rs.Open sqltemp, ,adOpenStatic, adLockReadOnly
'rs.movefirst:do while not rs.eof 
%>
<option value="<%=rs("chapa")%>"><%=rs("nome") & " (" & rs("chapa") & ")"%></option>
<%
'rs.movenext:loop
rs.close
if request.form("anoform")="" then anoform=year(now)-1 else anoform=request.form("anoform")
%>
	</select>
-->
	</td>
	<td class="campoa">Ano: <input type="text" name="anoform" SIZE="4" value="<%=anoform%>"></td>
</tr>
<tr>
	<td class="campoa" colspan=3>
	<input type="radio" name="saida" value="I" <%if tiposaida="I" then response.write "checked"%> >Impressora
	<input type="radio" name="saida" value="E" <%if tiposaida="E" then response.write "checked"%> >E-mail
	</td>
</tr>
<tr>
	<td class=titulo colspan=3>
	<input type="submit" class=button value="Visualizar" name="B1">
	</td>
</tr>
</table>
</form>
<%
end if

if request.form("B1")<>"" then
chapa=left (request.form("anochapa"),5)
chapa=request.form("chapaform")
ano  =right(request.form("anochapa"),4)
ano  =request.form("anoform")
tiposaida=request.form("saida")

sql1="SELECT f.CHAPA, f.NOME, p.CPF, p.email, f.codsituacao FROM corporerm.dbo.pfunc AS f, corporerm.dbo.ppessoa AS p " & _
"WHERE f.CHAPA='" & chapa & "' and p.codigo=f.codpessoa; "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
nome=rs("nome")
cpf=rs("cpf")
email=rs("email"):codsituacao=rs("codsituacao")
rs.close
if cpf<>"" then cpf=left(cpf,3) & "." & mid(cpf,4,3) & "." & mid(cpf,7,3) & "-" & right(cpf,2)

sql1="SELECT CHAPA, NRODEPENDIRRF FROM corporerm.dbo.PFHSTNDP WHERE CHAPA='" & chapa & "' AND Year(DTMUDANCA)<=" & ano & " order by dtmudanca desc;"
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
	ndep=rs("nrodependirrf")
else
	ndep=0
end if
rs.close

sql1="select VALOR from corporerm.dbo.PVALFIX where CODIGO='04' and '" & ano & "'+'1231' between INICIOVIGENCIA and FINALVIGENCIA "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
	vdep=rs("valor")
else
	vdep=0
end if
rs.close
dedudep=ndep*cdbl(vdep)

sql1="SELECT Year(DTPAGTO) AS Expr1, ff.CHAPA " & _
",'3_01'=sum(case when b.grupo='3.01' then valor*fator else 0 end) " & _
",'3_02'=sum(case when b.grupo='3.02' then valor*fator else 0 end) " & _
",'3_03'=sum(case when b.grupo='3.03' then valor*fator else 0 end) " & _
",'3_04'=sum(case when b.grupo='3.04' then valor*fator else 0 end) " & _
",'3_05'=sum(case when b.grupo='3.05' then valor*fator else 0 end) " & _
",'4_01'=sum(case when b.grupo='4.01' then valor*fator else 0 end) " & _
",'4_02'=sum(case when b.grupo='4.02' then valor*fator else 0 end) " & _
",'4_03'=sum(case when b.grupo='4.03' then valor*fator else 0 end) " & _
",'4_04'=sum(case when b.grupo='4.04' then valor*fator else 0 end) " & _
",'4_05'=sum(case when b.grupo='4.05' then valor*fator else 0 end) " & _
",'4_06'=sum(case when b.grupo='4.06' then valor*fator else 0 end) " & _
",'4_07'=sum(case when b.grupo='4.07' then valor*fator else 0 end) " & _
",'4_08'=sum(case when b.grupo='4.08' then valor*fator else 0 end) " & _
",'4_09'=sum(case when b.grupo='4.09' then valor*fator else 0 end) " & _
",'5_01'=sum(case when b.grupo='5.01' then valor*fator else 0 end) " & _
",'5_02'=sum(case when b.grupo='5.01' and b.codevento='063' then valor*abs(fator) else 0 end) " & _
",'6_01'=sum(case when b.grupo='6.01' then valor*fator else 0 end) " & _
"FROM base_informe b INNER JOIN ( " & _
"SELECT * FROM corporerm.dbo.PFFINANC WHERE CHAPA='" & chapa & "' UNION ALL SELECT * FROM corporerm.dbo.PFFINANCCOMPL WHERE CHAPA='" & chapa & "' " & _
") ff ON b.CODEVENTO=ff.CODEVENTO collate database_default " & _
"WHERE Year(DTPAGTO)=" & ano & " AND ff.CHAPA='" & chapa & "'  " & _
"GROUP BY Year(DTPAGTO), ff.CHAPA "

if tiposaida="I" then

rs.Open sql1, ,adOpenStatic, adLockReadOnly
'rs.movefirst
'do while not rs.eof 
%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=campo align="left"><img src="../images/republica.gif" border="0" width="70">
	</td>
	<td class=campo><p style="font-size: 12 pt;font-family: Tahoma;margin-top: 0; margin-bottom: 0"><b>MINISTÉRIO DA FAZENDA</b></p>
			<p style="font-size: 10 pt;font-family: Tahoma;margin-top: 0; margin-bottom: 0"><b>Secretaria da Receita Federal</b></p>
	</td>
	<td class=campo align="center" width=50%><p style="font-size: 9 pt;font-family: Tahoma;margin-top: 0; margin-bottom: 0">COMPROVANTE DE RENDIMENTOS PAGOS E DE RETENÇÃO DE IMPOSTO DE RENDA NA FONTE
	<br>Ano-Calendário <%=ano%>
	</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" width="650">
<tr>
	<td colspan=2 class=campo><b>1. Fonte Pagadora Pessoa Jurídica ou Pessoa Física</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		Nome Empresarial/Nome</td>
	<td class="campor" style="border-top:1px solid #000000;border-right:1px solid #000000">CNPJ/CPF</td>
</tr>
<tr>
	<td class="campop" style="border-left:1px solid black;border-bottom:1px solid #000000;border-right:1px solid #000000">
		&nbsp;&nbsp;&nbsp;FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</td>
	<td class="campop" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
		&nbsp;&nbsp;&nbsp;73.063.166/0001-20</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: separate" width="650">
<tr>
	<td colspan=2 class=campo><b>2. Pessoa Física Beneficiária dos Rendimentos</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">CPF</td>
	<td class="campor" style="border-top:1px solid #000000;border-right:1px solid #000000">Nome Completo</td>
</tr>
<tr>
	<td class="campop" style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">&nbsp;&nbsp;&nbsp;<%=cpf%></td>
	<td class="campop" style="border-bottom:1px solid #000000;border-right:1px solid #000000">&nbsp;&nbsp;&nbsp;<%=nome%></td>
</tr>
<tr>
	<td class="campor" colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">Natureza do Rendimento</td>
</tr>
<tr>
	<td class="campop" colspan=2 style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">&nbsp;&nbsp;&nbsp;Rendimentos do trabalho assalariado</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=campo><b>3. Rendimentos Tributáveis, Deduções e Imposto sobre a Renda Retido na Fonte</td>
	<td class=campo align="center" width=110><b>Valores em reais</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		01. Total dos Rendimentos (inclusive Férias)</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%if isnull(rs("3_01")) then v301="0,00" else v301=formatnumber(rs("3_01"),2)%>
	<input type="text" class="form_input" style="text-align:right;background-color:white;font-size:10px" size="10" name=totalf value="<%=v301%>">&nbsp;&nbsp;
	</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		02. Contribuição Previdenciária Oficial</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%if isnull(rs("3_02")) then response.write "0,00" else response.write formatnumber(rs("3_02"),2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		03. Contribuição à Previdência Privada e ao Fundo de Aposentadoria Programada Individual-FAPI</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%if isnull(rs("3_03")) then response.write "0,00" else response.write formatnumber(rs("3_03"),2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		04. Pensão Alimentícia (preencher também o quadro 7)</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%if isnull(rs("3_04")) then response.write "0,00" else response.write formatnumber(rs("3_04"),2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		05. Imposto sobre a Renda Retido na Fonte</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%if isnull(rs("3_05")) then response.write "0,00" else response.write formatnumber(rs("3_05"),2)%>&nbsp;&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=campo><b>4. Rendimentos Isentos e Não Tributáveis</td>
	<td class=campo align="center" nowrap width=110><b>Valores em reais</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		01. Parcela Isenta dos Proventos de Aposentadoria, Reserva, Reforma e Pensão (65 anos ou mais)</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%if isnull(rs("4_01")) then response.write "0,00" else response.write formatnumber(rs("4_01"),2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		02. Diárias e Ajuda de Custo</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%if isnull(rs("4_02")) then response.write "0,00" else response.write formatnumber(rs("4_02"),2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		03. Pensão, Proventos de Aposentadoria ou Reforma por Moléstia Grave e Aposentadoria ou Reforma por Acidente em Serviço</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%if isnull(rs("4_03")) then response.write "0,00" else response.write formatnumber(rs("4_03"),2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		04. Lucro e Dividendo Apurado a partir de 1996 pago por PJ (Lucro Real, Presumido ou Arbitrado)</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%if isnull(rs("4_04")) then response.write "0,00" else response.write formatnumber(rs("4_04"),2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		05. Valores Pagos ao Titular ou Sócio de Microempresa ou Empresa de Pequeno Porte, exceto Pro-Labore, Aluguéis ou Serviços Prestados</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%if isnull(rs("4_05")) then response.write "0,00" else response.write formatnumber(rs("4_05"),2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		06. Indenização por rescisão de contrato de trabalho, inclusive a título de PDV, e acidente de trabalho</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%if isnull(rs("4_06")) then response.write "0,00" else response.write formatnumber(rs("4_06"),2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		07. Outros</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%if isnull(rs("4_07")) then response.write "0,00" else response.write formatnumber(rs("4_07"),2)%>&nbsp;&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=campo><b>5. Rendimentos sujeitos à Tributação Exclusiva (rendimento líquido)</td>
	<td class=campo align="center" width=110><b>Valores em reais</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		01. Décimo Terceiro Salário</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
<%
if cdbl(rs("5_01"))-dedudep<0 then v501=0 else v501=cdbl(rs("5_01"))-dedudep
%>
		<%if isnull(rs("5_01")) then response.write "0,00" else response.write formatnumber(v501,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		02. Imposto sobre a renda retido na fonte sobre 13º Salário</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%if isnull(rs("5_02")) then response.write "0,00" else response.write formatnumber(cdbl(rs("5_02")),2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		03. Outros</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=campo colspan="2"><b>6. Rendimentos Recebidos Acumuladamente - Art. 12-A da Lei nº 7.713, de 1988 (Sujeitos à Tributação Exclusiva)</td>
</tr>
<tr>
	<td class=campo>
		<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="530">
		<tr><td class="campor"><b>6.1 </b>Número do Processo: </td>
			<td class="campor"><b>Quantidade de Meses</b></td>
			<td class="campor">0.0</b></td>
		</tr>
		<tr><td class="campor" colspan="3">Natureza do Rendimento: </td>
		</tr>
		</table>
	</td>
	<td class=campo align="center" width=110><b>Valores em reais</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		01. Total dos rendimentos tributáveis (inclusive férias e décimo terceiro salário)</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		02. Exclusão: Despesas com a ação judicial</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		03. Dedução: Contribuição previdenciária oficial</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		04. Dedução: Pensão alimentícia (preencher também o quadro 7)</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		05. Imposto sobre a renda retido na fonte</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		06. Rendimentos isentos de pensão, proventos de aposentadoria ou reforma por moléstia grave ou aposentadoria ou reforma por acidente em serviço</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
</table>


<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td colspan=2 class=campo><b>7. Informações Complementares</td>
</tr>
<tr>
<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
<%
if not isnull(rs("3_04")) and cdbl(rs("3_04"))>0 then
response.write "Informações sobre o linha 4 quadro 3:<br>"
sqltp="select CHAPA, 'Pensao'=SUM(f.VALOR) from corporerm.dbo.PFFINANC f inner join base_informe i on i.CODEVENTO=f.CODEVENTO collate database_default " & _
"where f.CHAPA='" & chapa & "' and YEAR(dtpagto)=" & ano & " and i.grupo='3.04' group by f.CHAPA"
rs2.Open sqltp, ,adOpenStatic, adLockReadOnly
totalpensao=rs2("pensao")
rs2.close
sqld="select p.CHAPA, d.NRODEPEND, d.NOME, D.CPF, D.RESPONSAVEL, 'total'=sum(p.VALOR) from corporerm.dbo.PFDEPMOV p  " & _
"inner join corporerm.dbo.PFDEPEND d on d.CHAPA=p.CHAPA and d.NRODEPEND=p.NRODEPEND " & _
"where p.CHAPA='" & chapa & "' and ( (ANOCOMP=" & ano & " AND MESCOMP<=11 AND NROPERIODO=2) OR " & _
"(ANOCOMP=" & ano & "-1 AND MESCOMP=12 AND NROPERIODO=2) ) group by p.CHAPA, d.NRODEPEND, d.NOME, D.CPF, D.RESPONSAVEL " 
rs2.Open sqld, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
Response.write "Nome do pensionista: " & rs2("nome") & " | CPF: " & rs2("cpf") 
if rs2("responsavel")<>"" or not isnull(rs2("responsavel")) then response.write " | " & rs2("responsavel")
response.write " | Total de pensão pago: " & formatnumber(rs2("total"),2)
response.write "<br>"
rs2.movenext
loop
rs2.close
'response.write "<br><br>"
end if
%>
Pagamentos a plano de saúde:<br>
<%
sql1="SELECT o.codigo, o.razaosocial, o.cnpj, o.ans, 'total'=SUM(f.valor) " & _
"FROM assmed_empresa o inner join assmed_empresa_evento e on e.codigo=o.codigo " & _
"inner join corporerm.dbo.PFFINANC f on f.CODEVENTO=e.codevento collate database_default " & _
"where CHAPA='" & chapa & "' and YEAR(f.DTPAGTO)=" & ano & " group by o.codigo, o.razaosocial, o.cnpj, o.ans " 
rs2.Open sql1, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
	response.write "Operadora: " & rs2("cnpj") & " - " & rs2("razaosocial")
	sql2="SELECT h.CHAPA, h.NRODEPEND, 'total'=sum(h.VALOR) " & _
	"FROM CORPORERM.DBO.PFHSTASSMED h inner join assmed_empresa_evento e on e.codevento=h.CODEVENTO collate database_default inner join assmed_empresa o on o.codigo=e.codigo " & _
	"where h.CHAPA='" & chapa & "' and YEAR(h.dtpagto)=" & ano & " and o.codigo='" & rs2("codigo") & "' and h.NRODEPEND=0 group by h.CHAPA, h.NRODEPEND "
	rs3.Open sql2, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then totaltit=rs3("total")
	if isnull(totaltit) or totaltit="" then totaltit=0
	response.write "<br> - Valor pago no ano referente ao titular: R$ " & formatnumber(totaltit,2)
	rs3.close

	sql3="SELECT h.CHAPA, h.NRODEPEND, d.nome, d.cpf, 'total'=sum(h.VALOR) " & _
	"FROM CORPORERM.DBO.PFHSTASSMED h inner join assmed_empresa_evento e on e.codevento=h.CODEVENTO collate database_default " & _
	"inner join assmed_empresa o on o.codigo=e.codigo left join corporerm.dbo.PFDEPEND d on d.CHAPA=h.CHAPA and d.NRODEPEND=h.NRODEPEND " & _
	"where h.CHAPA='" & chapa & "' and YEAR(h.dtpagto)=" & ano & " and o.codigo='" & rs2("codigo") & "' and h.NRODEPEND>0 group by h.CHAPA, h.NRODEPEND, d.NOME, d.CPF "
	rs3.Open sql3, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then
	response.write "<br> - Valor pago no ano referente aos dependentes: "
	response.write "<table><tr><td class=campor>CPF</td><td class=campor>Nome</td><td class=campor>Valor</td></tr>"
	do while not rs3.eof
		totaldep=rs3("total")
		if isnull(totaldep) or totaldep="" then totaldep=0
		response.write "<tr><td class=campor>" & rs3("cpf") & "</td><td class=campor>" & rs3("nome") & "</td><td class=campor>" & formatnumber(totaldep,2) & "</td></tr>"
	rs3.movenext
	loop
	response.write "</table>"
	end if
	rs3.close
	
rs2.movenext
loop
rs2.close
%>
</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td colspan=3 class=campo><b>8. Responsável pelas informações</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		Nome</td>
	<td class="campor" style="border-top:1px solid #000000;border-right:1px solid #000000">DATA</td>
	<td class="campor" style="border-top:1px solid #000000;border-right:1px solid #000000" width=150>Assinatura</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">
		&nbsp;&nbsp;&nbsp;FUND.INSTITUTO DE ENSINO PARA OSASCO</td>
	<td class="campor" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
		&nbsp;&nbsp;&nbsp;<%="28/02/" & ano+1%></td>
	<td class="campor" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
		&nbsp;&nbsp;&nbsp;</td>
</tr>
</table>
<p style="margin-top:0;margin-bottom:0;font-size:7pt">Aprovado pela IN/SRF nº 120/2000 / Dispensa de Assinatura conforme Instrução Normativa 
nº 120 de 28/12/2000
<%
'rs.movenext
if rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página -->
'loop
rs.close
end if 'tipo saida=I

if tiposaida="E" then
rs.Open sql1, ,adOpenStatic, adLockReadOnly
'rs.movefirst
'do while not rs.eof 

cab0="<html><head><meta http-equiv=""CONTENT-TYPE"" content=""text/html; charset=windows-1252"">" & _
"<title>Informe de Rendimentos - Funcionários</title>" & _
"<link rel=""stylesheet"" type=""text/css"" href=""http://rh.unifieo.br/diversos.css"">" & _
"</head><body>"
cab0="<html><style type='text/css'>" & _
"<!--" & _
"td.titulo { font-size:8pt; font-family:tahoma; font-weight:bold; background-color:Silver; color:Black;} " & _
"td.titulop { font-size:10pt; font-family:tahoma; font-weight:bold; background-color:Silver; color:Black;} " & _
"td.campo { font-size:8pt; font-family:tahoma; font-weight:normal; background-color:White; font-size-adjust:inherit; font-stretch:inherit;} " & _
"td.campop { font-size:10pt; font-family:tahoma; font-weight:normal; background-color:White; font-size-adjust:inherit; font-stretch:inherit;} " & _
"td.campor { font-size:9px; font-family:tahoma; font-weight:normal; background-color:White; font-style:inherit; font-variant:normal; font-size-adjust:0; font-stretch:inherit;}" & _
"td.fundor { font-size:9px; font-family:tahoma; font-weight:normal; background-color:Silver; color:Black;} " & _
"p { font-size:10pt; font-family:tahoma; font-weight:normal;} " & _
"-->"&_
"</style><body>"

cab1="<table border=""0"" cellpadding=""2"" cellspacing=""0"" style=""border-collapse: collapse"" width=""650"">" & _
"<tr><td class=campo align=""left""><img src=""http://rh.unifieo.br/images/republica.gif"" border=""0"" width=""70"">" & _
"</td><td class=campo><p style=""font-size: 12 pt;font-family: Tahoma;margin-top: 0; margin-bottom: 0""><b>MINISTÉRIO DA FAZENDA</b></p>" & _
"<p style=""font-size: 10 pt;font-family: Tahoma;margin-top: 0; margin-bottom: 0""><b>Secretaria da Receita Federal</b></p>" & _
"</td><td class=campo align=""center"" width=""50%""><p style=""font-size: 9 pt;font-family: Tahoma;margin-top: 0; margin-bottom: 0"">COMPROVANTE DE RENDIMENTOS PAGOS E DE RETENÇÃO DE IMPOSTO DE RENDA NA FONTE" & _
"<br>Ano-Calendário " & ano & "</td></tr></table>"

cab2="<table border=""0"" cellpadding=""2"" cellspacing=""0"" width=""650"">" & _
"<tr><td colspan=2 class=campo><b>1. Fonte Pagadora Pessoa Jurídica ou Pessoa Física</td></tr>" & _
"<tr><td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000"">Nome Empresarial/Nome</td>" & _
"<td class=""campor"" style=""border-top:1px solid #000000;border-right:1px solid #000000"">CNPJ/CPF</td></tr>" & _
"<tr><td class=""campop"" style=""border-left:1px solid black;border-bottom:1px solid #000000;border-right:1px solid #000000"">" & _
"&nbsp;&nbsp;&nbsp;FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</td><td class=""campop"" style=""border-bottom:1px solid #000000;border-right:1px solid #000000"">" & _
"&nbsp;&nbsp;&nbsp;73.063.166/0001-20</td></tr></table>"

cab3="<table border=""0"" cellpadding=""2"" cellspacing=""0"" style=""border-collapse: separate"" width=""650"">" & _
"<tr><td colspan=2 class=campo><b>2. Pessoa Física Beneficiária dos Rendimentos</td></tr>" & _
"<tr><td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000"">CPF</td>" & _
"<td class=""campor"" style=""border-top:1px solid #000000;border-right:1px solid #000000"">Nome Completo</td></tr>" & _
"<tr><td class=""campop"" style=""border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000"">&nbsp;&nbsp;&nbsp;" & cpf & "</td>" & _
"<td class=""campop"" style=""border-bottom:1px solid #000000;border-right:1px solid #000000"">&nbsp;&nbsp;&nbsp;" & nome & "</td>" & _
"</tr><tr><td class=""campor"" colspan=2 style=""border-left:1px solid #000000;border-right:1px solid #000000"">Natureza do Rendimento</td>" & _
"</tr><tr><td class=""campop"" colspan=2 style=""border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000"">&nbsp;&nbsp;&nbsp;Rendimentos do trabalho assalariado</td>" & _
"</tr></table>"

if isnull(rs("3_01")) then v301="0,00" else v301=formatnumber(rs("3_01"),2)
if isnull(rs("3_02")) then v302="0,00" else v302=formatnumber(rs("3_02"),2)
if isnull(rs("3_03")) then v303="0,00" else v303=formatnumber(rs("3_03"),2)
if isnull(rs("3_04")) then v304="0,00" else v304=formatnumber(rs("3_04"),2)
if isnull(rs("3_05")) then v305="0,00" else v305=formatnumber(rs("3_05"),2)
cab4="<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""border-collapse: collapse"" width=""650"">" & _
"<tr><td class=campo><b>3. Rendimentos Tributáveis, Deduções e Imposto sobre a Renda Retido na Fonte</td>" & _
"<td class=campo align=""center"" width=110><b>Valores em reais</td></tr><tr>" & _
"<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000"">01. Total dos Rendimentos (inclusive Férias)</td>" & _
"<td class=""campor"" align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000"">" & v301 & "&nbsp;&nbsp;" & _
"</td></tr><tr> " & _
"<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000"">" & _
"02. Contribuição Previdenciária Oficial</td>" & _
"<td class=""campor"" align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000"">" & _
"" & v302 & "&nbsp;&nbsp;</td></tr><tr>" & _
"<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000"">" & _
"03. Contribuição à Previdência Privada e ao Fundo de Aposentadoria Programada Individual-FAPI</td>" & _
"<td class=""campor"" align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000"">" & v303 & "&nbsp;&nbsp;</td>" & _
"</tr><tr><td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000"">" & _
"04. Pensão Alimentícia (preencher também o quadro 7)</td>" & _
"<td class=""campor"" align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000"">" & v304 & "&nbsp;&nbsp;</td>" & _
"</tr><tr>" & _
"<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">" & _
"05. Imposto sobre a Renda Retido na Fonte</td>" & _
"<td class=""campor"" align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">" & v305 & "&nbsp;&nbsp;</td>" & _
"</tr></table>"

if isnull(rs("4_01")) then v401="0,00" else v401=formatnumber(rs("4_01"),2)
if isnull(rs("4_02")) then v402="0,00" else v402=formatnumber(rs("4_02"),2)
if isnull(rs("4_03")) then v403="0,00" else v403=formatnumber(rs("4_03"),2)
if isnull(rs("4_04")) then v404="0,00" else v404=formatnumber(rs("4_04"),2)
if isnull(rs("4_05")) then v405="0,00" else v405=formatnumber(rs("4_05"),2)
if isnull(rs("4_06")) then v406="0,00" else v406=formatnumber(rs("4_06"),2)
if isnull(rs("4_07")) then v407="0,00" else v407=formatnumber(rs("4_07"),2)
cab5="<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""border-collapse: collapse"" width=""650"">" & _
"<tr><td class=campo><b>4. Rendimentos Isentos e Não Tributáveis</td><td class=campo align=""center"" nowrap width=110><b>Valores em reais</td></tr><tr>" & _
"<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000"">" & _
"01. Parcela Isenta dos Proventos de Aposentadoria, Reserva, Reforma e Pensão (65 anos ou mais)</td>" & _
"<td class=""campor"" align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000"">" & v401 & "&nbsp;&nbsp;</td>" & _
"</tr><tr>" & _
"<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000"">02. Diárias e Ajuda de Custo</td>" & _
"<td class=""campor"" align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000"">" & v402 & "&nbsp;&nbsp;</td>" & _
"</tr><tr>" & _
"<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000"">03. Pensão, Proventos de Aposentadoria ou Reforma por Moléstia Grave e Aposentadoria ou Reforma por Acidente em Serviço</td>" & _
"<td class=""campor"" align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000"">" & v403 & "&nbsp;&nbsp;</td>" & _
"</tr><tr>" & _
"<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000"">		04. Lucro e Dividendo Apurado a partir de 1996 pago por PJ (Lucro Real, Presumido ou Arbitrado)</td>" & _
"<td class=""campor"" align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000"">" & v404 & "&nbsp;&nbsp;</td>" & _
"</tr><tr>" & _
"<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000"">" & _
"05. Valores Pagos ao Titular ou Sócio de Microempresa ou Empresa de Pequeno Porte, exceto Pro-Labore, Aluguéis ou Serviços Prestados</td>" & _
"<td class=""campor"" align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000"">" & v405 & "&nbsp;&nbsp;</td>" & _
"</tr><tr>" & _
"<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000"">" & _
"06. Indenização por rescisão de contrato de trabalho, inclusive a título de PDV, e acidente de trabalho</td>" & _
"<td class=""campor"" align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000"">" & v406 & "&nbsp;&nbsp;</td>" & _
"</tr><tr>" & _
"<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">		07. Outros</td>" & _
"<td class=""campor"" align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">" & v407 & "&nbsp;&nbsp;</td>" & _
"</tr></table>"

if isnull(rs("5_01")) then v501="0,00" else v501=formatnumber(cdbl(rs("5_01"))-dedudep,2)
if isnull(rs("5_02")) then v502="0,00" else v502=formatnumber(rs("5_02"),2)
cab6="<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""border-collapse: collapse"" width=""650"">" & _
"<tr><td class=campo><b>5. Rendimentos sujeitos à Tributação Exclusiva (rendimento líquido)</td><td class=campo align=""center"" width=110><b>Valores em reais</td>" & _
"</tr><tr>" & _
"<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000"">		01. Décimo Terceiro Salário</td>" & _
"<td class=""campor"" align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000"">" & v501 & "&nbsp;&nbsp;</td>" & _
"</tr><tr>" & _
"<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000"">" & _
"02. Imposto sobre a renda retido na fonte sobre 13º Salário</td>" & _
"<td class=""campor"" align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000"">" & v502 & "&nbsp;&nbsp;</td>" & _
"</tr><tr>" & _
"<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">		03. Outros</td>" & _
"<td class=""campor"" align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">0,00&nbsp;&nbsp;</td>" & _
"</tr></table>"

cab7="<table border=""0"" cellpadding=""3"" cellspacing=""0"" style=""border-collapse: collapse"" width=""650"">" & _
"<tr>	<td class=campo colspan=""2""><b>6. Rendimentos Recebidos Acumuladamente - Art. 12-A da Lei nº 7.713, de 1988 (Sujeitos à Tributação Exclusiva)</td>" & _
"</tr><tr>	<td class=campo>" & _
"		<table border=""1"" cellpadding=""2"" cellspacing=""0"" style=""border-collapse: collapse"" width=""530"">" & _
"		<tr><td class=""campor""><b>6.1 </b>Número do Processo: </td>" & _
"			<td class=""campor""><b>Quantidade de Meses</b></td>" & _
"			<td class=""campor"">0.0</b></td></tr>" & _
"		<tr><td class=""campor"" colspan=""3"">Natureza do Rendimento: </td></tr></table>" & _
"	</td>	<td class=""campo"" align=""center"" width=""110""><b>Valores em reais</td></tr>" & _
"<tr>	<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">" & _
"		01. Total dos rendimentos tributáveis (inclusive férias e décimo terceiro salário)</td>" & _
"	<td class=""campor"" width=110 align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">0,00&nbsp;&nbsp;</td>" & _
"</tr><tr>" & _
"	<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">" & _
"		02. Exclusão: Despesas com a ação judicial</td>" & _
"	<td class=""campor"" width=110 align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">0,00&nbsp;&nbsp;</td>" & _
"</tr><tr>" & _
"	<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">" & _
"		03. Dedução: Contribuição previdenciária oficial</td>" & _
"	<td class=""campor"" width=110 align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">0,00&nbsp;&nbsp;</td>" & _
"</tr><tr>" & _
"	<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">" & _
"		04. Dedução: Pensão alimentícia (preencher também o quadro 7)</td>" & _
"	<td class=""campor"" width=110 align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">0,00&nbsp;&nbsp;</td>" & _
"</tr><tr>" & _
"	<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">" & _
"		05. Imposto sobre a renda retido na fonte</td>" & _
"	<td class=""campor"" width=110 align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">0,00&nbsp;&nbsp;</td>" & _
"</tr><tr>" & _
"	<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">" & _
"		06. Rendimentos isentos de pensão, proventos de aposentadoria ou reforma por moléstia grave ou aposentadoria ou reforma por acidente em serviço</td>" & _
"	<td class=""campor"" width=110 align=""right"" style=""border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">0,00&nbsp;&nbsp;</td>" & _
"</tr></table>"

cab8="<table border=""0"" cellpadding=""2"" cellspacing=""0"" style=""border-collapse: collapse"" width=""650"">" & _
"<tr><td colspan=2 class=campo><b>7. Informações Complementares</td></tr><tr>" & _
"<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000"">"

if not isnull(rs("3_04")) and cdbl(rs("3_04"))>0 then
	cab9="Informações sobre o linha 4 quadro 3:<br>"
	sqltp="select CHAPA, 'Pensao'=SUM(f.VALOR) from corporerm.dbo.PFFINANC f inner join base_informe i on i.CODEVENTO=f.CODEVENTO collate database_default " & _
	"where f.CHAPA='" & chapa & "' and YEAR(dtpagto)=" & ano & " and i.grupo='3.04' group by f.CHAPA"
	rs2.Open sqltp, ,adOpenStatic, adLockReadOnly
		totalpensao=rs2("pensao")
	rs2.close
	sqld="select p.CHAPA, d.NRODEPEND, d.NOME, D.CPF, D.RESPONSAVEL, 'total'=sum(p.VALOR) from corporerm.dbo.PFDEPMOV p " & _
	"inner join corporerm.dbo.PFDEPEND d on d.CHAPA=p.CHAPA and d.NRODEPEND=p.NRODEPEND " & _
	"where p.CHAPA='" & chapa & "' and ( (ANOCOMP=" & ano & " AND MESCOMP<=11 AND NROPERIODO=2) OR " & _
	"(ANOCOMP=" & ano & "-1 AND MESCOMP=12 AND NROPERIODO=2) ) group by p.CHAPA, d.NRODEPEND, d.NOME, D.CPF, D.RESPONSAVEL " 
	rs2.Open sqld, ,adOpenStatic, adLockReadOnly
	do while not rs2.eof
		cab9=cab9 & "Nome do pensionista: " & rs2("nome") & " | CPF: " & rs2("cpf") 
		if rs2("responsavel")<>"" or not isnull(rs2("responsavel")) then cab9=cab9 & " | " & rs2("responsavel")
		cab9=cab9 & " | Total de pensão pago: " & formatnumber(rs2("total"),2) & "<br>"
	rs2.movenext
	loop
	rs2.close
end if

cab10="Pagamentos a plano de saúde:<br>"

sql1="SELECT o.codigo, o.razaosocial, o.cnpj, o.ans, 'total'=SUM(f.valor) " & _
"FROM assmed_empresa o inner join assmed_empresa_evento e on e.codigo=o.codigo " & _
"inner join corporerm.dbo.PFFINANC f on f.CODEVENTO=e.codevento collate database_default " & _
"where CHAPA='" & chapa & "' and YEAR(f.DTPAGTO)=" & ano & " group by o.codigo, o.razaosocial, o.cnpj, o.ans " 
rs2.Open sql1, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
	cab10=cab10 & "Operadora: " & rs2("cnpj") & " - " & rs2("razaosocial")
	sql2="SELECT h.CHAPA, h.NRODEPEND, 'total'=sum(h.VALOR) " & _
	"FROM CORPORERM.DBO.PFHSTASSMED h inner join assmed_empresa_evento e on e.codevento=h.CODEVENTO collate database_default inner join assmed_empresa o on o.codigo=e.codigo " & _
	"where h.CHAPA='" & chapa & "' and YEAR(h.dtpagto)=" & ano & " and o.codigo='" & rs2("codigo") & "' and h.NRODEPEND=0 group by h.CHAPA, h.NRODEPEND "
	rs3.Open sql2, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then totaltit=rs3("total")
	if isnull(totaltit) or totaltit="" then totaltit=0
	cab10=cab10 & "<br> - Valor pago no ano referente ao titular: R$ " & formatnumber(totaltit,2)
	rs3.close

	sql3="SELECT h.CHAPA, h.NRODEPEND, d.nome, d.cpf, 'total'=sum(h.VALOR) " & _
	"FROM CORPORERM.DBO.PFHSTASSMED h inner join assmed_empresa_evento e on e.codevento=h.CODEVENTO collate database_default " & _
	"inner join assmed_empresa o on o.codigo=e.codigo left join corporerm.dbo.PFDEPEND d on d.CHAPA=h.CHAPA and d.NRODEPEND=h.NRODEPEND " & _
	"where h.CHAPA='" & chapa & "' and YEAR(h.dtpagto)=" & ano & " and o.codigo='" & rs2("codigo") & "' and h.NRODEPEND>0 group by h.CHAPA, h.NRODEPEND, d.NOME, d.CPF "
	rs3.Open sql3, ,adOpenStatic, adLockReadOnly
	if rs3.recordcount>0 then
	cab10=cab10 & "<br> - Valor pago no ano referente aos dependentes: "
	cab10=cab10 & "<table><tr><td class=campor>CPF</td><td class=campor>Nome</td><td class=campor>Valor</td></tr>"
	do while not rs3.eof
		totaldep=rs3("total")
		if isnull(totaldep) or totaldep="" then totaldep=0
		cab10=cab10 & "<tr><td class=campor>" & rs3("cpf") & "</td><td class=campor>" & rs3("nome") & "</td><td class=campor>" & formatnumber(totaldep,2) & "</td></tr>"
	rs3.movenext
	loop
	cab10=cab10 & "</table>"
	end if
	rs3.close
rs2.movenext
loop
rs2.close

cab10=cab10 & "</td></tr></table>"

cab11="<table border=""0"" cellpadding=""2"" cellspacing=""0"" style=""border-collapse: collapse"" width=""650"">" & _
"<tr>	<td colspan=3 class=campo><b>8. Responsável pelas informações</td></tr>" & _
"<tr>	<td class=""campor"" style=""border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000"">		Nome</td>" & _
"	<td class=""campor"" style=""border-top:1px solid #000000;border-right:1px solid #000000"">DATA</td>" & _
"	<td class=""campor"" style=""border-top:1px solid #000000;border-right:1px solid #000000"" width=150>Assinatura</td>" & _
"</tr><tr>" & _
"	<td class=""campor"" style=""border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000"">" & _
"		&nbsp;&nbsp;&nbsp;FUND.INSTITUTO DE ENSINO PARA OSASCO</td>" & _
"	<td class=""campor"" style=""border-bottom:1px solid #000000;border-right:1px solid #000000"">" & _
"		&nbsp;&nbsp;&nbsp;28/02/" & ano+1 & "</td>" & _
"	<td class=""campor"" style=""border-bottom:1px solid #000000;border-right:1px solid #000000"">" & _
"		&nbsp;&nbsp;&nbsp;</td>" & _
"</tr></table>" & _
"<p style=""margin-top:0;margin-bottom:0;font-size:7pt"">Aprovado pela IN/SRF nº 120/2000 / Dispensa de Assinatura conforme Instrução Normativa nº 120 de 28/12/2000" 
cab12="</body></html>"

rs.close

Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory. 
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network). 
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM

	email1=chapa & "@unifieo.br"
	if codsituacao="D" then email1=email
	Set Mailer = CreateObject("CDO.Message")
	Mailer.From = "rh@unifieo.br" ' e-mail de quem esta enviando a mensagem 
	Mailer.To = email1 ' e-mail de quem vai receber a mensagem 
	if email<>"" and email<>email1 then Mailer.CC=email
	'Mailer.BCC = emailchefe ' Com Cópia 
	Mailer.ReplyTo = "rh@unifieo.br"
	'Mailer.AttachFile "e:\home\login\dados\arquivo.txt" 'caso queira anexar algum arquivo ao seu e-mail 
	Mailer.Subject = "Informe de Rendimentos - Ano " & ano & "/" & ano+1
	'Mailer.TextBody = "Você tem mensagem" 
	Mailer.HtmlBody=cab0 & cab1 & cab2 & cab3 & cab4 & cab5 & cab6 & cab7 & cab8 & cab9 & cab10 & cab11 & cab12
'==This section provides the configuration information for the remote SMTP server.
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'cdoBasic
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "02379@unifieo.br"
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "eb541627"
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
Mailer.Configuration.Fields.Update

'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.unifieo.br"
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = cdoBasic
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "02379@unifieo.br"
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "*12345678"
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25 
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
'Mailer.Configuration.Fields.Update

'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mail.hover.com"
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1 'cdoBasic
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "edson@benevides.com"
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "queroserbenevides"
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = true
'Mailer.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
'Mailer.Configuration.Fields.Update

'==End remote SMTP server configuration section==
	teste=0
	Mailer.Send 
	Set Mailer = Nothing 
	response.write "<p> Email enviado para " & email1 & " " & email


end if 'tiposaida=E
end if 'request.form<>""

'rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>