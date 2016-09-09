<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Recadastramento - Unimed</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
set rs2=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
if request("chapa")<>"" then unico=" and m.chapa='" & request("chapa") & "' " else unico=""

sql1="select m.chapa, f.nome, m.codigo, f.dtnascimento, f.sexo, f.estcivil, f.cpf, rg=f.cartidentidade,  " & _
"orgao_emissor=case when f.orgemissorident is null then '' else f.orgemissorident end+'/'+case when f.ufcartident is null then '' else f.ufcartident end,  " & _
"data_exped=f.dtemissaoident, pispasep, mae, rua, numero, complemento, bairro, estado, cidade, cep, admissao, " & _
"funcao, secao, email, telefone1, telefone2, telefone3, codsecao " & _
"from assmed_mudanca m inner join qry_funcionarios f on f.chapa collate database_default=m.chapa " & _
"where m.empresa='V' and getdate() between m.ivigencia and m.fvigencia " & _
unico & _
"order by codsecao, nome "
rs.CursorLocation = adUseClient
rs.Open sql1, conexao ,adOpenStatic, adLockReadOnly
rs.activeconnection=nothing
do while not rs.eof
select case left(rs("codsecao"),2)
	case "01"
		filial="NARCISO"
	case "03"
		filial="VILA YARA"
	case "04"
		filial="JD.WILSON"
end select
if rs("sexo")="F" then 
	sf="X":sm="&nbsp;&nbsp;"
else
	sm="X":sf="&nbsp;&nbsp;"
end if
corborda="black"
%>

<div align="center">
<center>
<table border="0" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td colspan=2 class="campop" style="font-size:16px" align="center"><b>Atualização Cadastral Obrigatória - Plano Odontológico</b></td></tr>
<tr><td colspan=2 class="campop" style="font-size:12px" valign="middle" height="25"><b>Dados da Empregadora</b></td></tr>
<tr>
	<td class=campo style="border:1px solid">
	Razão Social: FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO
	</td>
	<td class=campo style="border:1px solid">
	Filial: <%=filial%>
	</td>
</tr>
<tr><td colspan=2 class="campop" style="font-size:12px" valign="middle" height="25"><b>Dados Gerais do Associado Titular</b></td></tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td colspan=4 class=campo>Nome: <%=fquadro3(rs("nome"),60)%></td></tr>
<tr>
	<td class=campo>Data Nasc: <%=fquadro3(fdata3(rs("dtnascimento")),10)%></td>
	<Td class=campo>Sexo: F (<%=sf%>) M (<%=sm%>)</td>
	<td class=campo>Estado Civil: <%=fquadro3(rs("estcivil"),20)%></td>
	<td class=campo>CPF/MF: <%=fquadro3(textopuro(rs("cpf"),2),11)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>RG: <%=fquadro3(rs("rg"),15)%></td>
	<td class=campo>Orgão Emissor: <%=fquadro3(rs("orgao_emissor"),8)%></td>
	<td class=campo>Pais Emissor: <%=fquadro3(espaco3("",6),6)%></td>
	<td class=campo>Data Exped.: <%=fquadro3(fdata3(rs("data_exped")),10)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>PIS/PASEP: <%=fquadro3(rs("pispasep"),11)%></td>
	<td class=campo>CNS/SUS: <%=fquadro3(espaco3("",12),12)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>Nome da Mãe: <%=fquadro3(rs("mae"),60)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>Endereço: <%=fquadro3(rs("rua"),40)%></td>
	<td class=campo>Nº: <%=fquadro3(rs("numero"),10)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><%if rs("complemento")="" or isnull(rs("complemento")) then complemento=espaco3("",15) else complemento=rs("complemento")%>
	<td class=campo>Compl.: <%=fquadro3(complemento,15)%></td>
	<td class=campo>UF: <%=fquadro3(rs("estado"),2)%></td>
	<td class=campo>Cidade: <%=fquadro3(rs("cidade"),15)%></td>
	<td class=campo>Bairro: <%=fquadro3(rs("bairro"),15)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>CEP: <%=fquadro3(rs("CEP"),8)%></td>
	<td class=campo>Data Admissão: <%=fquadro3(fdata3(rs("admissao")),10)%></td>
	<td class=campo>Nº Matricula: <%=fquadro3(rs("chapa"),5)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class="campor">Cargo: <%=fquadro3(rs("funcao"),30)%></td>
	<td class="campor">Depto.: <%=fquadro3(rs("secao"),30)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><%if rs("email")="" or isnull(rs("email")) then email=espaco3("",30) else email=rs("email")%>
	<td class=campo>E-mail: <%=fquadro3(email,40)%></td>
	<td class=campo>Cód. Associado.: <%=fquadro3(textopuro(rs("codigo"),2),14)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>Tel. Res.: (<%=fquadro3(espaco3("",2),2)%>) <%=fquadro3(textopuro(rs("telefone1"),2),8)%></td>
	<td class=campo>Tel. Cel.: (<%=fquadro3(espaco3("",2),2)%>) <%=fquadro3(textopuro(rs("telefone2"),2),8)%></td>
	<td class=campo>Tel. Com.: (<%=fquadro3(espaco3("",2),2)%>) <%=fquadro3(textopuro(rs("telefone3"),2),8)%></td>
</tr>
</table>
<table border="0" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr><td class="campop" style="font-size:12px;border:0px" valign="middle" height="25"><b>Dados Dependentes / Agregados</b></td></tr>
</table>
<%
sql2="select d.chapa, d.dependente, d.sexo, d.nascimento, d.parentesco, d.mae, d.cpf, m.codigo from assmed_dep d " & _
"inner join assmed_dep_mudanca m on m.id_dep=d.id_dep where m.empresa='V' and d.chapa='" & rs("chapa")& "' " & _
"and getdate() between m.ivigencia and m.fvigencia"
rs2.CursorLocation = adUseClient
rs2.Open sql2, conexao ,adOpenStatic, adLockReadOnly
totaldep=rs2.recordcount
rs2.activeconnection=nothing
do while not rs2.eof
if rs2("sexo")="F" then 
	sdf="X":sdm="&nbsp;"
else
	sdm="X":sdf="&nbsp;"
end if
%>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>Nome: <%=fquadro3(rs2("dependente"),40)%></td>
	<td class=campo>Data Nasc.: <%=fquadro3(fdata3(rs2("nascimento")),10)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>CPF: <%=fquadro3(textopuro(rs2("cpf"),2),11)%></td>
	<Td class=campo>Sexo: F (<%=sdf%>) M (<%=sdm%>)</td>
	<td class=campo>Universitário: S (&nbsp;&nbsp;) N (&nbsp;&nbsp;)</td>
	<td class=campo>Parentesco: <%=fquadro3(rs2("parentesco"),15)%></td>
	<td class=campo>Agregado: S (&nbsp;&nbsp;) N (&nbsp;&nbsp;)</td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>Nome da Mãe: <%=fquadro3(rs2("mae"),60)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>RG: <%=fquadro3(espaco3("",15),15)%></td>
	<td class=campo>Órgão Emissor: <%=fquadro3(espaco3("",6),6)%></td>
	<td class=campo>País Emissor: <%=fquadro3(espaco3("",8),8)%></td>
	<td class=campo>Data Exped.: <%=fquadro3(espaco3("",10),10)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>Endereço: <%=fquadro3(rs("rua"),40)%></td>
	<td class=campo>Nº: <%=fquadro3(rs("numero"),10)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>Compl.: <%=fquadro3(complemento,15)%></td>
	<td class=campo>UF: <%=fquadro3(rs("estado"),2)%></td>
	<td class=campo>Cidade: <%=fquadro3(rs("cidade"),15)%></td>
	<td class=campo>Bairro: <%=fquadro3(rs("bairro"),15)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>Atividade Principal: <%=fquadro3(espaco3("",30),30)%></td>
	<td class=campo>Tel. Cel.: (<%=fquadro3(espaco3("",2),2)%>) <%=fquadro3(espaco3("",8),8)%></td>
</tr>
<tr><td class=campo colspan=2 style="border:2px solid"></td></tr>
</table>
<%
rs2.movenext
loop
rs2.close
%>
<%
if totaldep<4 then
	for d=1 to (4-totaldep)
%>	
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>Nome: <%=space(10)%></td>
	<td class=campo>Data Nasc.: <%=space(10)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>CPF: <%=space(10)%></td>
	<Td class=campo>Sexo: F (&nbsp;&nbsp;) M (&nbsp;&nbsp;)</td>
	<td class=campo>Universitário: S (&nbsp;&nbsp;) N (&nbsp;&nbsp;)</td>
	<td class=campo>Parentesco: <%=space(10)%></td>
	<td class=campo>Agregado: S (&nbsp;&nbsp;) N (&nbsp;&nbsp;)</td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>Nome da Mãe: <%=space(10)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>RG: <%=space(10)%></td>
	<td class=campo>Órgão Emissor: <%=space(10)%></td>
	<td class=campo>País Emissor: <%=space(10)%></td>
	<td class=campo>Data Exped.: <%=space(10)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>Endereço: <%=space(10)%></td>
	<td class=campo>Nº: <%=space(10)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>Compl.: <%=space(10)%></td>
	<td class=campo>UF: <%=space(10)%></td>
	<td class=campo>Cidade: <%=space(10)%></td>
	<td class=campo>Bairro: <%=space(10)%></td>
</tr>
</table>
<table border="1" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo>Atividade Principal: <%=space(10)%></td>
	<td class=campo>Tel. Cel.: <%=space(10)%></td>
</tr>
<tr><td class=campo colspan=2 style="border:2px solid"></td></tr>
</table>
<%	
	next
end if
%>
<table border="0" bordercolor="<%=corborda%>" cellpadding="2" cellspacing="0" style="border-collapse: collapse" width="690">
<tr>
	<td class=campo colspan=2>
	Você, beneficiário do plano odontológico, deve providenciar o mais breve possível a atualização dos seus dados cadastrais, que se 
	encontram incompletos junto ao RH.
	<br>Este documento totalmente preenchido e assinado, deverá ser entregue no departamento de RH até ____/____/_____
	<br>O atraso na entrega deste documento poderá acarretar atraso na liberação do tratamento odontológico.
	<br>
	<br>
	Data: _____/_____/_____
	</td>
	<td class=campo rowspan=2 align="center">
		<img src="../images/ans_metlife.jpg" width="20">
	</td>
</tr>
<tr>
	<td class=campo>
	<br><br>Assinatura Titular _______________________________________
	</td>
	<td class="campop" valign=Bottom align="right">
	Central de Atendimento MetLife - 0800 638 5433
	<p style="font-size:7pt">Todos os campos são obrigatórios. (*) obrigatório ao menos 1 (um).</p>
	</td>
</tr>
</table>

</center>
</div>

<%
if rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>"
rs.movenext
loop
rs.close
set rs=nothing
set rs2=nothing

conexao.close
set conexao=nothing
%>
</body>
</html>