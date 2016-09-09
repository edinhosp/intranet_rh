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
<title>Termo de Rescisão do Contrato de Trabalho</title>
<script language="javascript" type="text/javascript">
<!--
/****************************************************
     Author: Eric King
     Url: http://redrival.com/eak/index.shtml
     This script is free to use as long as this info is left in
     Featured on Dynamic Drive script library (http://www.dynamicdrive.com)
****************************************************/
var win=null;
function NewWindow(mypage,myname,w,h,scroll,pos){
if(pos=="random"){LeftPosition=(screen.width)?Math.floor(Math.random()*(screen.width-w)):100;TopPosition=(screen.height)?Math.floor(Math.random()*((screen.height-h)-75)):100;}
if(pos=="center"){LeftPosition=(screen.width)?(screen.width-w)/2:100;TopPosition=(screen.height)?(screen.height-h)/2:100;}
else if((pos!="center" && pos!="random") || pos==null){LeftPosition=0;TopPosition=20}
settings='width='+w+',height='+h+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scroll+',location=no,directories=no,status=yes,menubar=no,toolbar=no,resizable=no';
win=window.open(mypage,myname,settings);}
// -->
</script>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<script language="JavaScript"> <!--
// Verifica se campos obrigatorios do formulario foram preenchidos
function nome1() { form.chapa.value=form.nome.value;form.submit(); }
function chapa1() { form.nome.value=form.chapa.value;form.submit(); }
--></script>
<%
dim conexao, rs, rs2
dim mes(12)
mes(1)="Janeiro":mes(2)="Fevereiro":mes(3)="Março":mes(4)="Abril":mes(5)="Maio":mes(6)="Junho"
mes(7)="Julho":mes(8)="Agosto":mes(9)="Setembro":mes(10)="Outubro":mes(11)="Novembro":mes(12)="Dezembro"
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

if request.form("B1")="" or request.form("id")="" then
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção de funcionário para Termo de Rescisão
<form method="POST" action="trct2011.asp" name="form">
<%
sqla="select f.chapa, f.nome from corporerm.dbo.pfunc f where f.codsituacao='D' or f.datademissao is not null "
sqla="select top 150 f.chapa, f.nome, datademissao from corporerm.dbo.pfunc f where ((f.codsituacao='D' or f.datademissao is not null) and f.chapa<'10000') order by datademissao desc, nome"
rs.Open sqla, ,adOpenStatic, adLockReadOnly
%>
<input type="text" name="chapa" size="5" maxlength="5" onchange="chapa1()" value="<%=request.form("chapa")%>">
<select name="nome" class=a onchange="nome1()" style="font-family:'Courier New';font-size:8pt;">
	<option value="0">Selecione o funcionário</option>
	<option value="00000" <%if request.form("chapa")="00000" then response.write "selected"%>>Por data de pagamento</option>
<%
rs.movefirst:do while not rs.eof
if request.form("chapa")=rs("chapa") then temps="selected" else temps=""
%>
	<option  value="<%=rs("chapa")%>" <%=temps%> ><%=rs("nome") & string(46-len(rs("nome"))," ") & " - " & rs("datademissao")%></option>
<%
rs.movenext:loop
rs.close
%>
</select>
<br><br>
<table border="1" cellpadding="2" cellspacing="2" style="border-collapse: collapse">
<tr>
	<td class=titulo></td>
	<td class=titulo>Ano</td>
	<td class=titulo>Mês</td>
	<td class=titulo>Período</td>
</tr>
<%
repete=0
if request.form("chapa")<>"00000" then
sqlp="select anocomp, mescomp, nroperiodo, descricao from corporerm.dbo.pfperff ff  " & _
"inner join corporerm.dbo.pfunc f on f.chapa=ff.chapa " & _
"where ff.chapa='" & request.form("chapa") & "' and nroperiodo not in (2) " & _
"and convert(datetime,convert(nvarchar,anocomp)+'/'+convert(nvarchar,mescomp)+'/'+'01')>=convert(datetime,convert(nvarchar,year(datademissao))+'/'+convert(nvarchar,month(datademissao))+'/'+'01') "
rs.Open sqlp, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
	do while not rs.eof
	id=rs("anocomp") & "|" & rs("mescomp") & "|" & rs("nroperiodo")
	'rs("nroperiodo") & rs("dtvencimento")
	if repete=0 and uano=rs("anocomp") and umes=rs("mescomp") then
		response.write "<tr><td class=campo><input type=""radio"" name=""id"" value=" & uano & "|" & umes & "|99></td>"
		response.write "<td class=campo>" & rs("anocomp") & "</td>"
		response.write "<td class=campo>" & rs("mescomp") & "</td>"
		response.write "<td class=campo>Junta periodos</td></tr>"
		repete=1
	end if
%>
<tr>
	<td class=campo><input type="radio" name="id" value="<%=id%>"></td>
	<td class=campo><%=rs("anocomp")%></td>
	<td class=campo><%=rs("mescomp")%></td>
	<td class=campo><%=rs("descricao")%></td>
</tr>
<%
	uano=rs("anocomp"):umes=rs("mescomp")
	rs.movenext:loop
end if
rs.close

else

sqlp="SELECT dtpagtorescisao, Count(CHAPA) AS Total FROM corporerm.dbo.Pfunc GROUP BY dtpagtorescisao HAVING dtpagtorescisao>getdate()-30 ORDER BY dtpagtorescisao DESC"
rs.Open sqlp, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
	do while not rs.eof
%>
<tr>
	<td class=campo><input type="radio" name="id" value="<%=rs("dtpagtorescisao")%>"></td>
	<td class=campo>(<%=rs("total")%> termo<%if rs("total")>1 then response.write "s"%>)</td>
	<td class=campo>&nbsp;</td>
	<td class=campo><%=rs("dtpagtorescisao")%></td>
</tr>
<%
	rs.movenext:loop
end if
rs.close

end if
%>
</table>
<br>
<input type="submit" value="Pesquisar" name="B1" class="button"></p>
</form>
<%
end if

if request.form("B1")<>"" and request.form("id")<>"" then
chapa=request.form("chapa")

if chapa="00000" then
	dtpagtorescisao=request.form("id")
else
	pos=0:dim varpos(2)
	for a=1 to len(request.form("id"))
		letra=mid(request.form("id"),a,1)
		if letra="|" then 
			pos=pos+1
		else
			varpos(pos)=varpos(pos) & letra
		end if
	next
	anocomp=varpos(0)
	mescomp=varpos(1)
	nroperiodo=varpos(2)
end if
contador=0

if chapa="00000" then
	sqlsel="select distinct chapa from corporerm.dbo.pfunc where dtpagtorescisao='" & dtaccess(dtpagtorescisao) & "' "
	rs.Open sqlsel, ,adOpenStatic, adLockReadOnly
	do while not rs.eof
		redim preserve ch(contador):ch(contador)=rs("chapa")
		'redim preserve dv(contador):dv(contador)=rs("dtvencimento")
		'redim preserve dv(contador):dv(contador)=rs("dtvencimento")
		'redim preserve np(contador):np(contador)=rs("nroperiodo")
	rs.movenext
	contador=contador+1
	loop
	rs.close
else
	redim ch(0),ac(0),mc(0),np(0),dr(0)
	ch(0)=chapa
	ac(0)=anocomp
	mc(0)=mescomp
	np(0)=nroperiodo
end if

for b=0 to ubound(ch)
sql1="select f.chapa, f.nome, f.codsecao, f.funcao, f.carteiratrab, f.seriecarttrab, f.ufcarttrab, f.cpf, f.dtnascimento, f.mae, f.pispasep, f.rua, f.numero, f.complemento, f.bairro, f.cidade, f.estado, f.cep, " & _
"s.cgc, ruae=s.rua, numeroe=s.numero, complementoe=s.complemento, bairroe=s.bairro, cidadee=s.cidade, estadoe=s.estado, cepe=s.cep, s.cnaerais, f.salario, f.codrecebimento, f.demissao, codnivelsal, motivodemissao, " & _
"f.admissao, dtavisoprevio, codcategoria, codsindicato " & _
"from qry_funcionarios f, corporerm.dbo.psecao s where s.codigo=f.codsecao and f.chapa='" & ch(b) & "' "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
session("chapa")=rs("chapa")
session("chapanome")=rs("nome")
dr(b)=rs("demissao"):calculo=dr(b)

'sql2="select dtiniperaquis, dtfimperaquis, dtinigozo, dtfimgozo, diasabono, nrofaltas " & _
'"from corporerm.dbo.pfhstfer_old where chapa='" & ch(b) & "' and nroperiodo=" & np(b) & " and dtfimperaquis='" & dtaccess(dv(b)) & "' "
'rs2.Open sql2, ,adOpenStatic, adLockReadOnly
'rs2.close

sql21="select top 1 dtmudanca from corporerm.dbo.pfhstsal where chapa='" & ch(b) & "' and dtmudanca<='" & dtaccess(calculo) & "' order by dtmudanca desc"
sql2="select salario from corporerm.dbo.pfhstsal where chapa='" & ch(b) & "' and dtmudanca in (" & sql21 & ") "
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
salario=rs2("salario")
rs2.close
%>
<div align="center">
<center>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr><td class=titulop align="center" valign="center" height=30><font size="+1">TERMO DE RESCISÃO DO CONTRATO DE TRABALHO</td></tr>
<tr><td class=titulop align="center" valign="center">IDENTIFICAÇÃO DO EMPREGADOR</td></tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=campo>01 CNPJ/CEI<br><font size="2"><%=rs("cgc")%></td>
	<td class=campo>02 Razão Social/Nome<br><font size="2">FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</td>
</tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=campo>03 Endereço (logradouro, nº, andar, apartamento)<br><font size="2"><%=rs("ruae") & " " & rs("numeroe") & " " & rs("complementoe")%></td>
	<td class=campo>04 Bairro<br><font size="2"><%=rs("bairroe")%></td>
</tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=campo>05 Município<br><font size="2"><%=rs("cidadee")%></td>
	<td class=campo>06 UF<br><font size="2"><%=rs("estadoe")%></td>
	<td class=campo>07 CEP<br><font size="2"><%=rs("cepe")%></td>
	<td class=campo>08 CNAE<br><font size="2"><%=rs("cnaerais")%></td>
	<td class=campo>09 CNPJ/CEI Tomador/Obra<br><font size="2">&nbsp;</td>
</tr>
<tr><td class=titulop align="center" colspan=5 valign="center">IDENTIFICAÇÃO DO TRABALHADOR</td></tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=campo>10 PIS/PASEP<br><font size="2"><%=rs("pispasep")%></td>
	<td class=campo>11 Nome<br><font size="2"><b><%=rs("nome")%></b></td>
</tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=campo>12 Endereço (logradouro, nº, andar, apartamento)<br><font size="2"><%=rs("rua") & " " & rs("numero") & " " & rs("complemento")%></td>
	<td class=campo>13 Bairro<br><font size="2"><%=rs("bairro")%></td>
</tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=campo>14 Município<br><font size="2"><%=rs("cidade")%></td>
	<td class=campo>15 UF<br><font size="2"><%=rs("estado")%></td>
	<td class=campo>16 CEP<br><font size="2"><%=rs("cep")%></td>
	<td class=campo>17 Carteira de Trabalho (nº, série, UF)<br><font size="2"><%=rs("carteiratrab") & " / " & rs("seriecarttrab") & " / " & rs("ufcarttrab")%></td>
</tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=campo>18 CPF<br><font size="2"><%=rs("cpf")%></td>
	<td class=campo>19 Data de nascimento<br><font size="2"><%=rs("dtnascimento")%></td>
	<td class=campo>20 Nome da mãe<br><font size="2"><%=rs("mae")%></td>
</tr>
<tr><td class=titulop align="center" colspan=3 valign="center">DADOS DO CONTRATO</td></tr>
</table>
<%
select case rs("codnivelsal")
	case "0"
		tipocontrato="2. Contrato de trabalho por prazo determinado c/cláusula assecuratória de rescisão"
	case "1"
		tipocontrato="1. Contrato de trabalho por prazo indeterminado"
	case else
		tipocontrato="1. Contrato de trabalho por prazo indeterminado"
end select
sql2="select descricao from corporerm.dbo.pmotdemissao where codcliente='" & rs("motivodemissao") & "'"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount=0 then motivo="02" else motivo=rs2("descricao")
rs2.close
if rs("motivodemissao")="" or isnull(rs("motivodemissao")) then motivodemissao="02" else motivodemissao=rs("motivodemissao")
select case motivodemissao
	case "02"
		codafast="I1"
	case "04"
		codafast="J"
	case "03"
		codafast="H"
	case "23"
		codafast="S2"
	case "06"
		codafast="I3"
	case else
		codafast="XXXXXXXX"
end select
sql2="select cnpj, codentidade, nome from corporerm.dbo.psindic where codigo='" & rs("codsindicato") & "'"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
cnpj_sind=rs2("cnpj")
cod_sind=rs2("codentidade")
nome_sind=rs2("nome")
rs2.close
%>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=campo>21 Tipo de Contrato<br><font size="2"><%=tipocontrato%></td>
	<td class=campo>22 Causa do Afastamento<br><font size="2"><%=motivo%></td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=campo>23 Remuneração Mês Anterior Afast.<br><font size="2"><%=formatnumber(salario,2)%></td>
	<td class=campo>24 Data de Admissão<br><font size="2"><%=rs("admissao")%></td>
	<td class=campo valign=top>25 Data do Aviso Prévio<br><font size="2"><%=rs("dtavisoprevio")%></td>
	<td class=campo valign=top>26 Data de afastamento<br><font size="2"><%=rs("demissao")%></td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=campo>27 Cód. Afastamento<br><font size="2"> &nbsp;<%=codafast%></td>
	<td class=campo>28 Pensão Alimentícia (%) (TRCT)<br><font size="2"><input type="text" class="form_input10" value="" size="4"></td>
	<td class=campo>29 Pensão Alimentícia (%) (Saque FGTS)<br><font size="2"><input type="text" class="form_input10" value="" size="4"></td>
	<td class=campo>30 Categoria do Trabalhador<br><font size="2"> <%=rs("codcategoria")%></td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=campo>31 Código Sindical<br><font size="2"><%=cod_sind%></td>
	<td class=campo>32 CNPJ e Nome da Entidade Sindical Laboral<br><font size="2"><%=cnpj_sind & " <span style='font-size:8pt'>" & nome_sind & "</span>"%></td>
</tr>
<tr><td class=titulop align="center" colspan=2 valign="center">DISCRIMINAÇÃO DAS VERBAS RESCISÓRIAS</td></tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr><td class=campo align="left" colspan=6 valign="center"><b>VERBAS RESCISÓRIAS</td></tr>
<tr>
	<td class=campo><b>Rubrica</td><td class=campo><b>Valor</td>
	<td class=campo><b>Rubrica</td><td class=campo><b>Valor</td>
	<td class=campo><b>Rubrica</td><td class=campo><b>Valor</td>
</tr>
<%
if np(b)="99" then filtro1="" else filtro1=" and nroperiodo=" & np(b)
sql2="select f.codevento, f.ref, f.valor, e.descricao, e.provdescbase, t.rubrica, r.agrupar, r.descricao desc_rubrica " & _
"from corporerm.dbo.pffinanc f inner join corporerm.dbo.pevento e on e.codigo=f.codevento " & _
"left join trct_eventos t on t.codigo=e.codigo " & _
"inner join trct_rubricas r on r.rubrica=t.rubrica " & _
"where f.chapa='" & ch(b) & "' " & filtro1 & " and mescomp=" & mc(b) & " and anocomp=" & ac(b) & " " & _
"and e.provdescbase<>'B' and valor>0 " & _
"order by e.provdescbase desc, agrupar, t.rubrica, codevento " 
'"where f.chapa='" & ch(b) & "' and nroperiodo=" & np(b) & " and mescomp=" & mc(b) & " and anocomp=" & ac(b) & " " & _

rs2.Open sql2, ,adOpenStatic, adLockReadOnly
totprov=0:totdesc=0:totalizou=0
contacols=0
do while not rs2.eof
valor=cdbl(rs2("valor"))
if rs2("provdescbase")="P" then totprov=totprov+valor
if rs2("provdescbase")="D" then totdesc=totdesc+valor
if contacols=0 then response.write "<tr height=30>"
if ulttipo<>rs2("provdescbase") and rs2.absoluteposition>1 then 
	if contacols<3 then for a=contacols to 2:response.write "<td class=campo></td><td class=campo></td>":next
	response.write "</tr>":contacols=0
	response.write "<tr height=30><td class=campo></td><td class=campo></td><td class=campo></td><td class=campo></td>"
	response.write "<td class=fundo><b>TOTAL RESCISÓRIO BRUTO</td>"
	response.write "<td class=fundo align="right"><b>" & formatnumber(totprov,2) & "</td></tr>"
	response.write "<tr><td class=campo align="left" colspan=6 valign="center"><b>DEDUÇÕES</td></tr>"
	response.write "<tr>"
	response.write "<td class=campo><b>Desconto</td><td class=campo><b>Valor</td>"
	response.write "<td class=campo><b>Desconto</td><td class=campo><b>Valor</td>"
	response.write "<td class=campo><b>Desconto</td><td class=campo><b>Valor</td>"
	response.write "</tr><tr height=30>"
	totalizou=1
end if
desc_rubrica=replace(rs2("desc_rubrica"),"[right4]", right(rs2("descricao"),4))
if isnull(rs2("ref")) then ref=0.00 else ref=rs2("ref")
desc_rubrica=replace(desc_rubrica,"[ref]", ref)
t=rs2("codevento")

if rs2("rubrica")="95" or rs2("rubrica")="115" then desc_rubrica=desc_rubrica & " - " & rs2("descricao")
if t="509" or t="510" then desc_rubrica=desc_rubrica & " - Média"
if t="221" or t="227" then desc_rubrica=desc_rubrica & " - " & rs2("ref") & " dias"

if rs2("codevento")="087" then desc_rubrica=desc_rubrica & " " & rs2("ref") & " dias"
if rs2("codevento")="073" then desc_rubrica=desc_rubrica & " " & rs2("ref") & " aulas"

if t="089" or t="088" _
	then desc_rubrica=desc_rubrica & " " & rs2("ref") & " horas"

teste=rs2.fields(5)
%>
	<td class=campo valign=top><%=rs2("rubrica") & " " & desc_rubrica%></td>
	<td class=campo valign=top align="right">&nbsp;<%=formatnumber(rs2("valor"),2)%>&nbsp;</td>
<%
contacols=contacols+1
if contacols=3 then contacols=0:response.write "</tr>"
ulttipo=rs2("provdescbase")
rs2.movenext
loop
if contacols<3 then for a=contacols to 2:response.write "<td class=campo></td><td class=campo></td>":next

if totalizou=0 then 
	'if contacols<3 then for a=contacols to 2:response.write "<td class=campo></td><td class=campo></td>":next
	response.write "</tr>":contacols=0
	response.write "<tr height=30><td class=campo></td><td class=campo></td><td class=campo></td><td class=campo></td>"
	response.write "<td class=fundo><b>TOTAL RESCISÓRIO BRUTO</td>"
	response.write "<td class=fundo align="right"><b>" & formatnumber(totprov,2) & "</td></tr>"
	response.write "<tr><td class=campo align="left" colspan=6 valign="center"><b>DEDUÇÕES</td></tr>"
	response.write "<tr>"
	response.write "<td class=campo><b>Desconto</td><td class=campo><b>Valor</td>"
	response.write "<td class=campo><b>Desconto</td><td class=campo><b>Valor</td>"
	response.write "<td class=campo><b>Desconto</td><td class=campo><b>Valor</td>"
	response.write "</tr><tr height=30>"
end if

response.write "<tr height=30><td class=campo></td><td class=campo></td><td class=campo></td><td class=campo></td>"
response.write "<td class=fundo><b>TOTAL DAS DEDUÇÕES</td>"
response.write "<td class=fundo align="right"><b>" & formatnumber(totdesc,2) & "</td></tr>"

response.write "<tr height=30><td class=campo></td><td class=campo></td><td class=campo></td><td class=campo></td>"
response.write "<td class=fundo><b>VALOR RESCISÓRIO LÍQUIDO</td>"
response.write "<td class=fundo align="right"><b>" & formatnumber(totprov-totdesc,2) & "</td></tr>"

response.write "<tr><td colspan=6 style=""border-bottom:1px solid"" height=1></td></tr></table><DIV style=""page-break-after:always""></DIV>"
response.write "<table border=""1"" bordercolor=""#000000"" cellpadding=""2"" cellspacing=""0"" style=""border-collapse: collapse;border-bottom:0px transparent"" width=""650"">"
response.write "<tr height=30><td class=campo width=""26%""></td><td class=campo width=""7%""></td>"
response.write "              <td class=campo width=""26%""></td><td class=campo width=""7%""></td>"
response.write "<td class=fundo><b>VALOR RESCISÓRIO LÍQUIDO</td>"
response.write "<td class=fundo align="right"><b>" & formatnumber(totprov-totdesc,2) & "</td></tr>"

'--------------------------------------
%>
</table>

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr><td class=titulop align="center" colspan=2 valign="center">FORMALIZAÇÃO DA RESCISÃO</td></tr>
<tr>
	<td class=campo valign=top width=50% height=40>150 Local e data do recebimento<br><font size="2">&nbsp;</td>
	<td class=campo valign=top >151 Carimbo e assinatura do empregador ou preposto<br><font size="2">&nbsp;</td>
</tr>
<tr>
	<td class=campo valign=top height=40>152 Assinatura do trabalhador<br><font size="2">&nbsp;</td>
	<td class=campo valign=top >153 Assinatura do responsável legal do trabalhador<br><font size="2">&nbsp;</td>
</tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=campo valign=top width=50% height=40><p style="margin-top:0px;margin-bottom:4px;font-size:8pt">154 HOMOLOGAÇÃO</p>
	<span style="font-size:8pt;">Foi prestada, gratuitamente, assistência ao trabalhador, nos termos do art. 477, § 1ì, da Consolidação das Leis do
	Trabalho - CLT, sendo comprovado, neste ato, o efetivo pagamento das verbas rescisórias acima especificadas.
	<br><br><br>
	________________________________________________<br>
	Local e data
	</span>
	&nbsp;</td>
	<td class=campo valign=top width=25%>155 Digital do trabalhador<br><font size="2">&nbsp;</td>
	<td class=campo valign=top >156 Digital do responsável legal<br><font size="2">&nbsp;</td>
</tr>
<tr>
	<td class=campo style="border-top:1px transparent">
	<br><br>
	________________________________________________<br>
	Carimbo e assinatura do assistente
	</td>
	<td class=campo colspan=2 valign=top style="border-bottom:0px transparent">158 Recepção pelo Banco (data e carimbo)</td>
</tr>
</table>


<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;" width="650">
<tr>
	<td class=campo valign=top width=50% height=80 style="border:1px solid">157 Identificação do orgão homologador<br><font size="2">&nbsp;</td>
	<td class=campo valign=top width=50% style="border-bottom:1px solid;border-right:1px solid"><br><font size="2">&nbsp;</td>
</tr>
</table>	
	
<table border="1" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-top:0px transparent" width="650">
<tr><td class=titulop align="center" colspan=2 valign="center" align="center">
A ASSISTÊNCIA NO ATO DE RESCISÃO CONTRATUAL É GRATUITA.<br>
<span style="font-size:8pt">
Pode o trabalhador iniciar ação judicial quanto aos créditos resultantes das relações de trabalho até o limite de dois anos após a extinção do contrato de
trabalho (Inc. XXIX, Art. 7º da Constituição Federal/1988).
</span>
</td></tr>
</table>


<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=10><tr><td></td></tr></table>

<table>
<%
rs2.movefirst
totprov=0:totdesc=0
teste=0
if teste=1 then

do while not rs2.eof
valor=cdbl(rs2("valor"))
if rs2("provdescbase")="P" then impprov=formatnumber(valor,2) else impprov="&nbsp;"
if rs2("provdescbase")="D" then impdesc=formatnumber(valor,2) else impdesc="&nbsp;"
if rs2("provdescbase")="P" then totprov=totprov+valor
if rs2("provdescbase")="D" then totdesc=totdesc+valor
if isnull(rs2("ref")) then 
		ref2="&nbsp;" 
	elseif cdbl(rs2("ref"))=0 then 
		ref2="&nbsp;" 
	else 
		ref2=rs2("ref") 
end if
%>
<tr>
	<td class="campop" style="border-left:1px solid #000000;" align="center" height=20><%=rs2("codevento")%></td>
	<td class="campop" style="border-left:1px solid #000000;" align="right"><%=ref2%>&nbsp;&nbsp;</td>
	<td class="campop" style="border-left:1px solid #000000;">&nbsp;<%=rs2("descricao")%></td>
	<td class="campop" style="border-left:1px solid #000000;" align="right"><%=impprov%>&nbsp;&nbsp;</td>
	<td class="campop" style="border-left:1px solid #000000;border-right:1px solid #000000" align="right"><%=impdesc%>&nbsp;&nbsp;</td>
</tr>
<%
rs2.movenext:loop
rs2.close
liquido=totprov-totdesc

%>
</table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650" height=10><tr><td></td></tr></table>

<%
end if

if b<ubound(ch) then response.write "<DIV style=""page-break-after:always""></DIV>"
rs.close

next

set rs=nothing
%>
<%
set rs=nothing
set rs2=nothing
end if ' 

conexao.close
set conexao=nothing
%>
</body>
</html>