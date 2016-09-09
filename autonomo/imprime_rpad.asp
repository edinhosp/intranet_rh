<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a52")="N" or session("a52")="" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Impressão de RPA</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")

sessao=session.sessionid
id_lanc=request("codigo")

sql="SELECT a.id_autonomo, a.nome_autonomo, r.descricao_servico, r.data_emissao, r.id_lanc, " & _
"r.data_pagamento, r.servico_prestado, r.desc_outros_rend, r.outros_rendimentos, r.desconto_ir, " & _
"r.desconto_inss, r.desconto_iss, r.descricao_outros, r.outros_descontos, r.valor_liquido, a.cpf, " & _
"a.nit, a.rg, a.orgao_rg, a.ccm, r.inss_outra_empresa, a.bancocod,a.banconome,a.conta,a.agencia " & _
"FROM autonomo_rpa AS r INNER JOIN autonomo AS a ON r.id_autonomo = a.id_autonomo " & _
"WHERE r.id_lanc=" & id_lanc
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
rs.Open sql, ,adOpenStatic, adLockReadOnly

rs.movefirst
do while not rs.eof
for a=1 to 2
%>
<hr>
<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width='650'>
<tr>
	<td class="campop" align="center"><b>RECIBO DE PAGAMENTO A AUTÔNOMO - RPA Nº <%=rs("id_lanc")%></b></td>
</tr>
</table>

<table border='1' cellpadding='2' cellspacing='0' style='border-collapse: collapse' bordercolor="#000000" width='650'>
<tr>
	<td class=campo align="center">NOME OU RAZÃO SOCIAL DA EMPRESA</td>
	<td class=campo align="center">MATRÍCULA (CNPJ OU INSS)</td>
</tr>
<tr>
	<td class=campo align="left"><b>FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</b></td>
	<td class=campo align="center"><b>73.063.166/0001-20</b></td>
</tr>
</table>

<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width='650'>
<tr><td class=campo align="right">RECEBI DA EMPRESA ACIMA IDENTIFICADA, PELA PRESTAÇÃO DOS SERVIÇOS</td></tr>
</table>

<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width='650'>
<tr>
	<td width=20 class=campo align="left">DE&nbsp;</td>
	<td width=400 class=campo style="border-bottom: 1px solid #000000"> <%=rs("descricao_servico")%></td>
	<td width=130 class=campo align="right">&nbsp;A IMPORTÂNCIA DE R$ </td>
	<td width=100 class=campo style="border-bottom: 1px solid #000000" align="right"><b><%=formatnumber(rs("valor_liquido"),2)%></b></td>
</tr>
</table>

<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width='650'>
<tr>
	<td width=600 class=campo style="border-bottom: 1px solid #000000">(<%=extenso2(rs("valor_liquido"))%>)</td>
	<td width=50 class=campo align="right">&nbsp;CONFORME</td>
</tr>
</table>

<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width='650'>
<tr><td class=campo align="left">DISCRIMINATIVO ABAIXO:</td></tr>
</table>

<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width='650'>
<tr>
	<td valign=top>
<!-- inicio quadro esquerda -->	
<%
base=cdbl(rs("servico_prestado"))+rs("outros_rendimentos")
%>
<table border='1' cellpadding='2' cellspacing='0' style='border-collapse: collapse' bordercolor="#000000" width='320'>
<tr>
	<td class=campo align="center">SALÁRIO-BASE</td>
	<td class=campo align="center">TAXA</td>
	<td class=campo align="center">VALOR PARA INSS</td>
</tr>
<tr>
	<td class=campo align="center"><%=formatnumber(base,2)%></td>
	<td class=campo align="center">20%</td>
	<td class=campo align="center"><%=formatnumber(base*0.2,2)%></td>
</tr>
<tr>
	<td class=campo align="left" colspan=2>INSS descontado em outra empresa</td>
	<td class=campo align="center"><%=formatnumber(rs("inss_outra_empresa"),2)%></td>
</tr>
</table>

<br>
<table border='1' cellpadding='2' cellspacing='0' style='border-collapse: collapse' bordercolor="#000000" width='320'>
<tr><td class=campo align="center">CARRETEIRO (VR. BASE P/CÁLCULO DO INSS)</td></tr>
<tr><td class=campo align="center">APLICAR 11,71% SOBRE O VALOR DO FRETE PAGO<br>&nbsp;</td></tr>
</table>
<%
cpf=rs("cpf")
if cpf<>"" then cpf=left(cpf,3) & "." & mid(cpf,4,3) & "." & mid(cpf,7,3) & "-" & right(cpf,2)
%>
<br>
<table border='1' cellpadding='2' cellspacing='0' style='border-collapse: collapse' bordercolor="#000000" width='320'>
<tr><td class=campo align="center" colspan=4>NÚMERO DE INSCRIÇÃO</td></tr>
<tr>
	<td class=campo align="left" colspan=2>1-nº INSS/PIS:</td>
	<td class=campo align="center" colspan=2><%=rs("nit")%></td>
</tr>
<tr>
	<td class=campo width=70 align="left">2-nº CPF:</td>
	<td class=campo width=95 align="center"><%=cpf%></td>
	<td class=campo width=80 align="left">3-nº CCM:</td>
	<td class=campo width=80 align="center"><%=rs("ccm")%></td>
</tr>
</table>

<br>
<table border='1' cellpadding='2' cellspacing='0' style='border-collapse: collapse' bordercolor="#000000" width='320'>
<tr><td class=campo align="center" colspan=2>DOCUMENTO DE IDENTIDADE</td></tr>
<tr>
	<td class=campo align="center">NÚMERO</td>
	<td class=campo align="center">ORGÃO EMISSÃO</td>
</tr>
<tr>
	<td class=campo align="center">&nbsp;<%=rs("rg")%></td>
	<td class=campo align="center">&nbsp;<%=rs("orgao_rg")%></td>
</tr>
</table>

<br>
<table border='1' cellpadding='2' cellspacing='0' style='border-collapse: collapse' bordercolor="#000000" width='320'>
<tr>
	<td class=campo align="center">LOCALIDADE</td>
	<td class=campo align="center">DATA</td>
</tr>
<tr>
	<td class=campo align="center">OSASCO</td>
	<td class=campo align="center">&nbsp;<%=rs("data_pagamento")%></td>
</tr>
</table>

<!-- final quadro esquerda -->	
	</td>
	<td valign=top>
<!-- inicio quadro direita -->	
<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width='320'>
<tr><td class=campo align="left"><b>ESPECIFICAÇÃO:</b></td></tr>
</table>

<table border='0' cellpadding='2' cellspacing='0' style='border-collapse: collapse' width='320'>
<tr>
	<td class="campop">I</td>
	<td class="campop" colspan=3>VALOR DO SERVIÇO PRESTADO....</td>
	<td class="campop">R$</td>
	<td class="campop" align="right" style="border-bottom: 1px solid #000000"><%=formatnumber(rs("servico_prestado"),2)%></td>
</tr>
<tr>
	<td class="campop">II</td>
	<td class="campop" colspan=3>&nbsp;<%=rs("desc_outros_rend")%></td>
	<td class="campop">R$</td>
	<td class="campop" align="right" style="border-bottom: 1px solid #000000"><%=formatnumber(rs("outros_rendimentos"),2)%></td>
</tr>
<tr>
	<td class="campop" align="right" colspan=4>Soma&nbsp;</td>
	<td class="campop">R$</td>
	<td class="campop" align="right"><%=formatnumber(base,2)%></td>
</tr>
<tr>
	<td class=campo align="left" colspan=6>&nbsp;<b>DESCONTOS:</b></td>
</tr>
<tr>
	<td class="campop" align="left">III</td>
	<td class="campop" align="left">IMP.RENDA FONTE</td>
	<td class="campop" align="left">R$</td>
	<td class="campop" align="right" style="border-bottom: 1px solid #000000"><%=formatnumber(rs("desconto_ir"),2)%></td>
	<td class="campop" align="left" colspan=2>&nbsp;</td>
</tr>
<tr>
	<td class="campop" align="left">IV</td>
	<td class="campop" align="left">ISS-FONTE</td>
	<td class="campop" align="left">R$</td>
	<td class="campop" align="right" style="border-bottom: 1px solid #000000"><%=formatnumber(rs("desconto_iss"),2)%></td>
	<td class="campop" align="left" colspan=2>&nbsp;</td>
</tr>
<tr>
	<td class="campop" align="left">V</td>
	<td class="campop" align="left">INSS</td>
	<td class="campop" align="left">R$</td>
	<td class="campop" align="right" style="border-bottom: 1px solid #000000"><%=formatnumber(rs("desconto_inss"),2)%></td>
	<td class="campop" align="left" colspan=2>&nbsp;</td>
</tr>
<%
descontos=cdbl(rs("desconto_inss"))+rs("desconto_ir")+rs("desconto_iss")+rs("outros_descontos")
%>
<tr>
	<td class="campop" align="left">VI</td>
	<td class="campop" align="left">&nbsp;<%=rs("descricao_outros")%></td>
	<td class="campop" align="left">R$</td>
	<td class="campop" align="right" style="border-bottom: 1px solid #000000"><%=formatnumber(rs("outros_descontos"),2)%></td>
	<td class="campop" align="left">R$</td>
	<td class="campop" align="right" style="border-bottom: 1px solid #000000"><%=formatnumber(descontos,2)%></td>
</tr>
<tr>
	<td class="campop" align="right" colspan=4>Valor Líquido&nbsp;</td>
	<td class="campop">R$</td>
	<td class="campop" align="right"><%=formatnumber(rs("valor_liquido"),2)%></td>
</tr>
</table>

<br>
<table border='0' cellpadding='2' cellspacing='0' style="border: 1px solid #000000" width='320'>
<tr><td class=campo align="center">ASSINATURA<br>&nbsp;<br>&nbsp;
</td></tr>
</table>

<br>
<table border='1' cellpadding='2' cellspacing='0' style="border-collapse: collapse" bordercolor="#000000" width='320'>
<tr><td class="campor" align="center">NOME COMPLETO</td></tr>
<tr><td class="campop" align="center"><b><%=rs("nome_autonomo")%></b></td></tr>
</table>

<!-- final quadro direita -->	
	</td>
</tr>
<tr>
	<td valign=top style="border-top: 1px solid #000000" class="campor"><%=rs("data_emissao")%>&nbsp;-&nbsp;&nbsp;
<%
'if rs("bancocod")<>"" then response.write rs("bancocod") & "/"
'if rs("agencia")<>"" then response.write rs("agencia") & "/"
'if rs("conta")<>"" then response.write rs("conta") & ""
%>
	</td>
	<td valign=top style="border-top: 1px solid #000000" class="campor" align="right">
	<% 
	if a=1 then response.write "Mod. 07/99 - FIEO - 1ª via: empresa"
	if a=2 then response.write "Mod. 07/99 - FIEO - 2ª via: trabalhador autônomo"
	%>
	</td>
</tr>
	
</table>
<%
next
'response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página --> 
rs.movenext
loop

sql = "UPDATE autonomo_rpa SET emitiu_rpa = -1 WHERE id_lanc=" & id_lanc
conexao.execute sql

rs.close
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>