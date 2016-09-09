<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a63")="N" or session("a63")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Informe de Rendimentos - Aut�nomo</title>
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

if request.form("B1")="" then
%>
<form method="POST" action="informe.asp" name="form">
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo colspan=2>Sele��es para emiss�o de Informe de Rendimentos</td>
</tr>
<tr>
	<td class="campoa">Nome do Prestador</td>
	<td class="campoa"><select size=1 name="idaut" onChange="javascript:submit()">
		<option value="">Selecione o prestador</option>
		<option value="0" <%if request.form("idaut")=0 then response.write "selected"%>>Todos os aut�nomos</option>
<%
sqltemp="select id_autonomo, nome_autonomo from autonomo order by nome_autonomo "
rs.Open sqltemp, ,adOpenStatic, adLockReadOnly
rs.movefirst:do while not rs.eof 
if request.form("idaut")<>"" then selec=cint(request.form("idaut")) else selec=0
%>
<option value="<%=rs("id_autonomo")%>" <%if rs("id_autonomo")=selec then response.write "selected"%>><%=rs("nome_autonomo")%></option>
<%
rs.movenext:loop
rs.close
%>
	</select>
	</td>
</tr>
<tr>
	<td class="campol">Ano Base</td>
	<td class="campol"><select size=1 name="anobase">
<%
if selec="0" then
	sqltemp="SELECT Year([data_pagamento]) AS anobase FROM autonomo_rpa GROUP BY Year([data_pagamento]) "
else
	sqltemp="SELECT Year([data_pagamento]) AS anobase FROM autonomo_rpa WHERE id_autonomo=" & selec & " GROUP BY Year([data_pagamento]) "
end if
rs.Open sqltemp, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst:do while not rs.eof 
%>
<option value="<%=rs("anobase")%>"><%=rs("anobase")%></option>
<%
rs.movenext:loop
end if
rs.close
%>
	</select>
	</td>
</tr>
<tr>
	<td class=titulo colspan=2>
	<input type="submit" class=button value="Visualizar" name="B1">
	</td>
</tr>
</table>
</form>
<%
end if

if request.form("B1")<>"" then
if request.form("idaut")=0 then
sql1="SELECT a.id_autonomo, a.nome_autonomo, a.cpf, Year([data_pagamento]) AS anobase, Sum([servico_prestado]+[outros_rendimentos]) AS rendimentos, Sum(r.desconto_inss) AS inss, Sum(r.desconto_ir) AS ir " & _
"FROM autonomo AS a INNER JOIN autonomo_rpa AS r ON a.id_autonomo = r.id_autonomo " & _
"GROUP BY a.id_autonomo, a.nome_autonomo, a.cpf, Year([data_pagamento]) " & _
"HAVING Year([data_pagamento])=" & request.form("anobase") & " order by a.nome_autonomo "
else
sql1="SELECT a.id_autonomo, a.nome_autonomo, a.cpf, Year([data_pagamento]) AS anobase, Sum([servico_prestado]+[outros_rendimentos]) AS rendimentos, Sum(r.desconto_inss) AS inss, Sum(r.desconto_ir) AS ir " & _
"FROM autonomo AS a INNER JOIN autonomo_rpa AS r ON a.id_autonomo = r.id_autonomo " & _
"GROUP BY a.id_autonomo, a.nome_autonomo, a.cpf, Year([data_pagamento]) " & _
"HAVING a.id_autonomo=" & request.form("idaut") & " AND Year([data_pagamento])=" & request.form("anobase") & " "
end if
rs.Open sql1, ,adOpenStatic, adLockReadOnly
rs.movefirst
do while not rs.eof 
cpf=rs("cpf")
if cpf<>"" then cpf=left(cpf,3) & "." & mid(cpf,4,3) & "." & mid(cpf,7,3) & "-" & right(cpf,2)
%>
<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=campo align="left"><img src="../images/republica.gif" border="0" width="70">
	</td>
	<td class=campo><p style="font-size: 12 pt;font-family: Tahoma;margin-top: 0; margin-bottom: 0"><b>MINIST�RIO DA FAZENDA</b></p>
			<p style="font-size: 10 pt;font-family: Tahoma;margin-top: 0; margin-bottom: 0"><b>Secretaria da Receita Federal</b></p>
	</td>
	<td class=campo align="center" width=50%><p style="font-size: 9 pt;font-family: Tahoma;margin-top: 0; margin-bottom: 0">COMPROVANTE DE RENDIMENTOS PAGOS E DE RETEN��O DE IMPOSTO DE RENDA NA FONTE
	<br>Ano-Calend�rio <%=rs("anobase")%>
	</td>
</tr>
</table>

<br>
<table border="0" cellpadding="2" cellspacing="0" width="650">
<tr>
	<td colspan=2 class=campo><b>1. FONTE PAGADORA PESSOA JUR�DICA OU PESSOA F�SICA</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		Nome Empresarial/Nome</td>
	<td class="campor" style="border-top:1px solid #000000;border-right:1px solid #000000">CNPJ/CPF</td>
</tr>
<tr>
	<td class="campop" style="border-left:1px solid black;border-bottom:1px solid #000000;border-right:1px solid #000000">
		&nbsp;&nbsp;&nbsp;FUNDA��O INSTITUTO DE ENSINO PARA OSASCO</td>
	<td class="campop" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
		&nbsp;&nbsp;&nbsp;73.063.166/0001-20</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: separate" width="650">
<tr>
	<td colspan=2 class=campo><b>2. PESSOA F�SICA BENEFICI�RIA DOS RENDIMENTOS</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		CPF</td>
	<td class="campor" style="border-top:1px solid #000000;border-right:1px solid #000000">
		Nome Completo</td>
</tr>
<tr>
	<td class="campop" style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">
		&nbsp;&nbsp;&nbsp;<%=cpf%></td>
	<td class="campop" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
		&nbsp;&nbsp;&nbsp;<%=rs("nome_autonomo")%></td>
</tr>
<tr>
	<td class="campor" colspan=2 style="border-left:1px solid #000000;border-right:1px solid #000000">
		Natureza do Rendimento</td>
</tr>
<tr>
	<td class="campop" colspan=2 style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">
		&nbsp;&nbsp;&nbsp;Rendimentos do trabalho sem v�nculo empregat�cio</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=campo><b>3. RENDIMENTOS TRIBUT�VEIS, DEDU��ES E IMPOSTO RETIDO NA FONTE</td>
	<td class=campo align="center" width=110><b>VALORES EM REAIS</td>
</tr>
<tr>
	<td class="campo" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		01. Total dos Rendimentos (inclusive F�rias)</td>
	<td class="campo" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%=formatnumber(rs("rendimentos"),2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campo" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		02. Contribui��o Previdenci�ria Oficial</td>
	<td class="campo" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%=formatnumber(rs("inss"),2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campo" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		03. <font style="font-size:9pt">Contribui��o � Previd�ncia Privada e ao Fundo de Aposentadoria Programada Individual-FAPI</td>
	<td class="campo" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campo" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		04. Pens�o Aliment�cia (informar o benefici�rio no quadro 7)</td>
	<td class="campo" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campo" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		05. Imposto de Renda Retido</td>
	<td class="campo" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%=formatnumber(rs("ir"),2)%>&nbsp;&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=campo><b>4. RENDIMENTOS ISENTOS E N�O TRIBUT�VEIS</td>
	<td class=campo align="center" nowrap width=110><b>VALORES EM REAIS</td>
</tr>
<tr>
	<td class="campo" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		01. <font style="font-size:8pt">Parcela Isenta dos Proventos de Aposentadoria, Reserva, Reforma e Pens�o (65 anos ou mais)</td>
	<td class="campo" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campo" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		02. Di�rias e Ajuda de Custo</td>
	<td class="campo" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campo" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		03. <font style="font-size:8pt">Pens�o, Proventos de Aposentadoria ou Reforma por Mol�stia Grave e Aposentadoria ou Reforma por Acidente em Servi�o</td>
	<td class="campo" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campo" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		04. <font style="font-size:8pt">Lucro e Dividendo Apurado a partir de 1996 pago por PJ (Lucro Real, Presumido ou Arbitrado)</td>
	<td class="campo" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campo" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		05. <font style="font-size:8pt">Valores Pagos ao Titular ou S�cio de Microempresa ou Empresa de Pequeno Porte, exceto Pro-Labore, Alugu�is ou Servi�os Prestados</td>
	<td class="campo" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campo" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		06. <font style="font-size:9pt">Indeniza��o por rescis�o de contrato de trabalho, inclusive a t�tulo de PDV, e acidente de trabalho</td>
	<td class="campo" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campo" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		07. Outros (especificar):</td>
	<td class="campo" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=campo><b>5. RENDIMENTOS SUJEITOS � TRIBUTA��O EXCLUSIVA (RENDIMENTO L�QUIDO)</td>
	<td class=campo align="center" width=110><b>VALORES EM REAIS</td>
</tr>
<tr>
	<td class="campo" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		01. D�cimo Terceiro Sal�rio</td>
	<td class="campo" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campo" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		02. Outros</td>
	<td class="campo" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="3" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td class=campo colspan="2"><b>6. Rendimentos Recebidos Acumuladamente - Art. 12-A da Lei n� 7.713, de 1988 (Sujeitos � Tributa��o Exclusiva)</td>
</tr>
<tr>
	<td class=campo>
		<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="530">
		<tr><td class="campor"><b>6.1 </b>N�mero do Processo: </td>
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
		01. Total dos rendimentos tribut�veis (inclusive f�rias e d�cimo terceiro sal�rio)</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		02. Exclus�o: Despesas com a a��o judicial</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		03. Dedu��o: Contribui��o previdenci�ria oficial</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		04. Dedu��o: Pens�o aliment�cia (preencher tamb�m o quadro 7)</td>
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
		06. Rendimentos isentos de pens�o, proventos de aposentadoria ou reforma por mol�stia grave ou aposentadoria ou reforma por acidente em servi�o</td>
	<td class="campor" align="right" style="border-top:1px solid #000000;border-right:1px solid #000000;border-bottom:1px solid #000000">
		<%=formatnumber(0,2)%>&nbsp;&nbsp;</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td colspan=2 class=campo><b>7. INFORMA��ES COMPLEMENTARES</td>
</tr>
<tr>
	<td class="campop" colspan=2 height=50 style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000;border-top:1px solid #000000">
	</td>
</tr>
</table>

<table border="0" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="650">
<tr>
	<td colspan=3 class=campo><b>8. RESPONS�VEL PELAS INFORMA��ES</td>
</tr>
<tr>
	<td class="campor" style="border-left:1px solid #000000;border-top:1px solid #000000;border-right:1px solid #000000">
		Nome</td>
	<td class="campor" style="border-top:1px solid #000000;border-right:1px solid #000000">DATA</td>
	<td class="campor" style="border-top:1px solid #000000;border-right:1px solid #000000" width=150>Assinatura</td>
</tr>
<tr>
	<td class="campop" style="border-left:1px solid #000000;border-bottom:1px solid #000000;border-right:1px solid #000000">
		&nbsp;&nbsp;&nbsp;ROGERIO MATEUS DOS SANTOS ARAUJO</td>
	<td class="campop" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
		&nbsp;&nbsp;&nbsp;<%="30/01/" & request.form("anobase")+1%></td>
	<td class="campop" style="border-bottom:1px solid #000000;border-right:1px solid #000000">
		&nbsp;&nbsp;&nbsp;</td>
</tr>
</table>
<p style="margin-top:0;margin-bottom:0;font-size:7pt">Aprovado pela IN/SRF n� 120/2000

<%
rs.movenext
if rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a p�gina -->
loop
rs.close

end if

'rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>