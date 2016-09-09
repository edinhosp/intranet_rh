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
<link rel="stylesheet" type="text/css" href="../trct2012.css">
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
conexao.Open application("consqlteste")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rs3=server.createobject ("ADODB.Recordset")
Set rs3.ActiveConnection = conexao

if request.form("B1")="" or request.form("id")="" then
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção de funcionário para Termo de Rescisão
<form method="POST" action="trct2012b.asp" name="form">
<%
sqla="select f.chapa, f.nome from corporerm_teste.dbo.pfunc f where f.codsituacao='D' or f.datademissao is not null "
sqla="select top 150 f.chapa, f.nome, datademissao from corporerm_teste.dbo.pfunc f where ((f.codsituacao='D' or f.datademissao is not null) and f.chapa<'10000') order by datademissao desc, nome"
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
sqlp="select anocomp, mescomp, nroperiodo, descricao from corporerm_teste.dbo.pfperff ff  " & _
"inner join corporerm_teste.dbo.pfunc f on f.chapa=ff.chapa " & _
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

sqlp="SELECT dtpagtorescisao, Count(CHAPA) AS Total FROM corporerm_teste.dbo.Pfunc GROUP BY dtpagtorescisao HAVING dtpagtorescisao>getdate()-30 ORDER BY dtpagtorescisao DESC"
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
<p>Impressão de valores zerados: <select name="zero">
<option value="Z">Não imprimir 0,00</option>
<option value="V">Imprimir 0,00</option>
</select>
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
	sqlsel="select distinct chapa from corporerm_teste.dbo.pfunc where dtpagtorescisao='" & dtaccess(dtpagtorescisao) & "' "
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
"f.admissao, dtavisoprevio, codcategoria, codsindicato, tipodemissao " & _
"from corporerm_teste.dbo.qry_funcionarios f, corporerm_teste.dbo.psecao s where s.codigo=f.codsecao and f.chapa='" & ch(b) & "' "
rs.Open sql1, ,adOpenStatic, adLockReadOnly
session("chapa")=rs("chapa")
session("chapanome")=rs("nome")
dr(b)=rs("demissao"):calculo=dr(b)
tpquit=datediff("m",rs("admissao"),rs("demissao"))
if tpquit>12 then termo="H" else termo="Q"

'sql2="select dtiniperaquis, dtfimperaquis, dtinigozo, dtfimgozo, diasabono, nrofaltas " & _
'"from corporerm_teste.dbo.pfhstfer where chapa='" & ch(b) & "' and nroperiodo=" & np(b) & " and dtfimperaquis='" & dtaccess(dv(b)) & "' "
'rs2.Open sql2, ,adOpenStatic, adLockReadOnly
'rs2.close

sql21="select top 1 dtmudanca from corporerm_teste.dbo.pfhstsal where chapa='" & ch(b) & "' and dtmudanca<='" & dtaccess(calculo) & "' and salario>0 order by dtmudanca desc"
sql2="select salario from corporerm_teste.dbo.pfhstsal where chapa='" & ch(b) & "' and dtmudanca in (" & sql21 & ") and salario>0"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
salario=rs2("salario")
rs2.close
%>
<div align="center">
<center>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr><td class=rtitulo1 align="center" valign="center" height=30><b>TERMO DE RESCISÃO DO CONTRATO DE TRABALHO</td></tr>
<tr><td height="5pt"></td></tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr><td class=rtitulo2 colspan="2" align="center" valign="center" style="border: 1px solid #000000"><b>IDENTIFICAÇÃO DO EMPREGADOR</td></tr>
<tr>
	<td class=rcampo1>&nbsp;01 CNPJ/CEI<br><span class=campo1>&nbsp;<%=rs("cgc")%></td>
	<td class=rcampo1>&nbsp;02 Razão Social/Nome<br><span class=campo1>&nbsp;FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</td>
</tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=rcampo1>&nbsp;03 Endereço (logradouro, nº, andar, apartamento)<br><span class=campo1>&nbsp;<%=rs("ruae") & " " & rs("numeroe") & " " & rs("complementoe")%></td>
	<td class=rcampo1>&nbsp;04 Bairro<br><span class=campo1>&nbsp;<%=rs("bairroe")%></td>
</tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=rcampo1>&nbsp;05 Município<br><span class=campo1>&nbsp;<%=rs("cidadee")%></td>
	<td class=rcampo1>&nbsp;06 UF<br><span class=campo1>&nbsp;<%=rs("estadoe")%></td>
	<td class=rcampo1>&nbsp;07 CEP<br><span class=campo1>&nbsp;<%=rs("cepe")%></td>
	<td class=rcampo1>&nbsp;08 CNAE<br><span class=campo1>&nbsp;<%=rs("cnaerais")%></td>
	<td class=rcampo1>&nbsp;09 CNPJ/CEI Tomador/Obra<br><span class=campo1>&nbsp;&nbsp;</td>
</tr>
<tr><td class=rtitulo2 align="center" colspan=5 valign="center"><b>IDENTIFICAÇÃO DO TRABALHADOR</td></tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=rcampo1 width=130>&nbsp;10 PIS/PASEP<br><span class=campo1>&nbsp;<%=rs("pispasep")%></td>
	<td class=rcampo1>&nbsp;11 Nome<br><span class=campo1>&nbsp;<%=rs("nome")%></b></td>
</tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=rcampo1>&nbsp;12 Endereço (logradouro, nº, andar, apartamento)<br><span class=campo1>&nbsp;<%=rs("rua") & " " & rs("numero") & " " & rs("complemento")%></td>
	<td class=rcampo1>&nbsp;13 Bairro<br><span class=campo1>&nbsp;<%=rs("bairro")%></td>
</tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=rcampo1>&nbsp;14 Município<br>&nbsp;<span class=campo1><%=rs("cidade")%></td>
	<td class=rcampo1>&nbsp;15 UF<br>&nbsp;<span class=campo1><%=rs("estado")%></td>
	<td class=rcampo1>&nbsp;16 CEP<br>&nbsp;<span class=campo1><%=rs("cep")%></td>
	<td class=rcampo1 width=150>&nbsp;17 Carteira de Trabalho (nº, série, UF)<br>&nbsp;<span class=campo1><%=rs("carteiratrab") & " / " & rs("seriecarttrab") & " / " & rs("ufcarttrab")%></td>
	<td class=rcampo1>&nbsp;18 CPF<br>&nbsp;<span class=campo1><%=rs("cpf")%></td>
</tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=rcampo1 width=130>&nbsp;19 Data de nascimento<br>&nbsp;<span class=campo1><%=dataform(rs("dtnascimento"))%></td>
	<td class=rcampo1>&nbsp;20 Nome da mãe<br>&nbsp;<span class=campo1><%=rs("mae")%></td>
</tr>
<tr><td class=rtitulo2 align="center" colspan=3 valign="center"><b>DADOS DO CONTRATO</td></tr>
</table>
<%
select case rs("codnivelsal")
	case "0"
		tipocontrato="2. Contrato de trabalho por prazo determinado com cláusula assecuratória de direito recíproco de rescisão antecipada"
	case "1"
		tipocontrato="1. Contrato de trabalho por prazo indeterminado"
	case else
		tipocontrato="1. Contrato de trabalho por prazo indeterminado"
end select
sql2="select descricao from corporerm_teste.dbo.pmotdemissao where codcliente='" & rs("motivodemissao") & "'"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
if rs2.recordcount=0 then motivo="02" else motivo=rs2("descricao")
rs2.close
if rs("motivodemissao")="" or isnull(rs("motivodemissao")) then motivodemissao="02" else motivodemissao=rs("motivodemissao")
select case motivodemissao
	case "02"
		cod_="SJ2":desc_="Despedida sem justa causa, pelo empregador"
	case "03"
		cod_="JC2":desc_="Despedida com justa causa, pelo empregador"
	case "04"
		cod_="SJ1":desc_="Rescisão contratual a pedido do empregado"
	case "23"
		cod_="FT1":desc_="Rescisão do contrato de trabalho por falecimento do empregado"
	case "06"
		cod_="PD0":desc_="Extinção normal do contrato de trabalho por prazo determinado"
	case "termino1"
		cod_="RA1":desc_="Rescisão antecipada, pelo empregado, do contrato de trabalho por prazo determinado"
	case "termino2"
		cod_="RA2":desc_="Rescisão antecipada, pelo empregador, do contrato de trabalho por prazo determinado"
	case else
		cod_="XX":desc_="XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
end select
sql2="select cnpj, codentidade, nome from corporerm_teste.dbo.psindic where codigo='" & rs("codsindicato") & "'"
rs2.Open sql2, ,adOpenStatic, adLockReadOnly
cnpj_sind=rs2("cnpj")
cod_sind=rs2("codentidade")
nome_sind=rs2("nome")
rs2.close
%>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr><td class=rcampo2>&nbsp;21 Tipo de Contrato<br>&nbsp;<span class=campo1><%=tipocontrato%></td></tr>
<tr><td class=rcampo2>&nbsp;22 Causa do Afastamento<br>&nbsp;<span class=campo1><%=desc_%></td></tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=rcampo1>&nbsp;23 Remuneração Mês Anterior<br>&nbsp;<span class=campo1><%=formatnumber(salario,2)%></td>
	<td class=rcampo1>&nbsp;24 Data de Admissão<br>&nbsp;<span class=campo1><%=dataform(rs("admissao"))%></td>
	<td class=rcampo1>&nbsp;25 Data do Aviso Prévio<br>&nbsp;<span class=campo1><%=dataform(rs("dtavisoprevio"))%></td>
	<td class=rcampo1>&nbsp;26 Data de afastamento<br>&nbsp;<span class=campo1><%=dataform(rs("demissao"))%></td>
	<td class=rcampo1>&nbsp;27 Cód. Afastamento<br>&nbsp;<span class=campo1><%=cod_%></td>
</tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=rcampo1>&nbsp;28 Pensão Alimentícia (%)<br>&nbsp;<span class=campo1><input type="text" class="form_input10" value="" size="4"></td>
	<td class=rcampo1>&nbsp;29 Pensão Alimentícia (%) (FGTS)<br>&nbsp;<span class=campo1><input type="text" class="form_input10" value="" size="4"></td>
	<td class=rcampo1>&nbsp;30 Categoria do Trabalhador<br>&nbsp;<span class=campo1><%=numzero(rs("codcategoria"),2)%></td>
</tr>
</table>
<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=rcampo1>&nbsp;31 Código Sindical<br>&nbsp;<span class=campo1><%=cod_sind%></td>
	<td class=rcampo1>&nbsp;32 CNPJ e Nome da Entidade Sindical Laboral<br>&nbsp;<span class=campo1><%=cnpj_sind & " <span style='font-size:10pt'>" & nome_sind & "</span>"%></td>
</tr>
<tr><td class=rtitulo2 align="center" colspan=2 valign="center"><b>DISCRIMINAÇÃO DAS VERBAS RESCISÓRIAS</td></tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr><td class=rtitulo3 align="left" colspan=6 valign="center"><b>&nbsp;VERBAS RESCISÓRIAS</td></tr>
<tr>
	<td class=rtitulo3><b>&nbsp;Rubrica</td><td class=rtitulo3><b>&nbsp;Valor</td>
	<td class=rtitulo3><b>&nbsp;Rubrica</td><td class=rtitulo3><b>&nbsp;Valor</td>
	<td class=rtitulo3><b>&nbsp;Rubrica</td><td class=rtitulo3><b>&nbsp;Valor</td>
</tr>
<%
if np(b)="99" then filtro1="" else filtro1=" and nroperiodo=" & np(b)
sql2="select distinct f.codevento, f.ref, f.valor, f.descricao, f.provdescbase, r.rubrica, r.agrupar, r.desc_rubrica " & _
"from (select r.agrupar, r.rubrica, r.descricao desc_rubrica, r.obrig, e.codigo, e.descricao, e.provdescbase " & _
"	from trct_rubricas r left join trct_eventos e on e.rubrica=r.rubrica where r.obrig=1) r " & _
"left join (select codevento, DESCRICAO, REF, valor, e.provdescbase " & _
"	from corporerm_teste.dbo.PFFINANC f inner join corporerm_teste.dbo.PEVENTO e on e.CODIGO=f.codevento " & _
"	where f.chapa='" & ch(b) & "' " & filtro1 & " and mescomp=" & mc(b) & " and anocomp=" & ac(b) & " " & _
"	and e.provdescbase<>'B' and valor>0 ) f on f.CODEVENTO=r.codigo " & _
"union " & _
"select f.codevento, f.ref, f.valor, e.descricao, e.provdescbase, t.rubrica, " & _
"r.agrupar, r.descricao desc_rubrica " & _
"from corporerm_teste.dbo.pffinanc f " & _
"	inner join corporerm_teste.dbo.pevento e on e.codigo=f.codevento " & _
"	left join trct_eventos t on t.codigo=e.codigo " & _
"	inner join trct_rubricas r on r.rubrica=t.rubrica " & _
"where f.chapa='" & ch(b) & "' " & filtro1 & " and mescomp=" & mc(b) & " and anocomp=" & ac(b) & " " & _
"and e.provdescbase<>'B' and valor>0 and obrig=0 " 

sql3="order by agrupar, r.rubrica, codevento desc "
sql4=sql2 & sql3
'response.write sql4
rs2.Open sql4, ,adOpenStatic, adLockReadOnly
totprov=0:totdesc=0:totalizou=0
contacols=0:contaprov=0:contaded=0:t95=1:t115=1

do while not rs2.eof
colsprov=36:colsded=15
apendice=""
if ultrubrica=rs2("rubrica") and isnull(rs2("valor")) then rs2.movenext
if isnull(rs2("valor")) then valor=0 else valor=cdbl(rs2("valor"))
if rs2("provdescbase")="P" then totprov=totprov+valor
if rs2("provdescbase")="D" then totdesc=totdesc+valor

if contacols=0 then response.write "<tr height=30>"
if ulttipo<>rs2("agrupar") and rs2.absoluteposition>1 then 
	if contacols<3 then for a=contacols to 2:contaprov=contaprov+1:response.write "<td class=rcampo3></td><td class=rcampo4></td>":next
	response.write "</tr>":contacols=0
	if contaprov<colsprov then
		linhatemp=(colsprov-contaprov)/3
		for l=1 to linhatemp
			response.write "<tr><td class=rcampo3></td><td class=rcampo4></td><td class=rcampo3></td><td class=rcampo4></td><td class=rcampo3></td><td class=rcampo4></td></tr>"
		next
	end if
	response.write "<tr height=30><td class=rcampo3></td><td class=rcampo4></td><td class=rcampo3></td><td class=rcampo4></td>"
	response.write "<td class=rtitulo2><b>" & "TOTAL BRUTO</td>"
	response.write "<td class=rtitulo2 align="right"><b>" & formatnumber(totprov,2) & "</td></tr>"
	response.write "<tr><td class=rtitulo3 align="left" colspan=6 valign="center"><b>&nbsp;DEDUÇÕES</td></tr>"
	response.write "<tr>"
	response.write "<td class=rtitulo3><b>&nbsp;Desconto</td><td class=rtitulo3><b>&nbsp;Valor</td>"
	response.write "<td class=rtitulo3><b>&nbsp;Desconto</td><td class=rtitulo3><b>&nbsp;Valor</td>"
	response.write "<td class=rtitulo3><b>&nbsp;Desconto</td><td class=rtitulo3><b>&nbsp;Valor</td>"
	response.write "</tr><tr height=30>"
	totalizou=1
end if
if isnull(rs2("descricao")) then descricao="" else descricao=rs2("descricao")
desc_rubrica=rs2("desc_rubrica")
desc_rubrica=replace(desc_rubrica,"[right4]", right(descricao,4))
if isnull(rs2("ref")) then ref=0.00 else ref=rs2("ref")
desc_rubrica=replace(desc_rubrica,"[ref]", ref)
t=rs2("codevento")

if rs2("rubrica")="95" or rs2("rubrica")="115" then desc_rubrica=descricao
if t="509" or t="510" then desc_rubrica=desc_rubrica & " - Média"
if t="221" or t="227" then desc_rubrica=desc_rubrica & " - " & ref & " dias"
if rs2("rubrica")="64" and valor>0 then desc_rubrica=replace(desc_rubrica,"[ano13]", year(rs("dtdemissao"))) else desc_rubrica=replace(desc_rubrica,"[ano13]", "____")

if t="087" then desc_rubrica=desc_rubrica & " " & ref & " dias"
if t="073" then desc_rubrica=desc_rubrica & " " & ref & " aulas"

if t="089" or t="088" then desc_rubrica=desc_rubrica & " " & ref & " horas"

if rs2("rubrica")="66" and valor>0 then
	sqlf="select dtvencferias, saldoferias from corporerm_teste.dbo.pfunc where chapa='" & rs("chapa") & "' "
	rs3.Open sqlf, ,adOpenStatic, adLockReadOnly
		dtvencf=rs3("dtvencferias")
		saldoferias=rs3("saldoferias")
	rs3.close
	dtinicf=dateadd("yyyy",-1,dtvencf)
	dtinicf=dateadd("d",1,dtinicf)
	stringferias=" (" & ref & ") " & dtinicf & " a " & dtvencf
else
	stringferias="__/__/__ a __/__/__"
	saldoferias=0
end if
'if rs2("rubrica")="66" then response.write ref & "///" & saldoferias
if rs2("rubrica")="66" and valor>0 and cdbl(saldoferias)>0 and cdbl(ref)<>cdbl(saldoferias) then
	dtinicf2=dateadd("yyyy",1,dtinicf)
	dtvencf2=dateadd("yyyy",1,dtvencf)
	resto=cdbl(ref)-cdbl(saldoferias)
	stringferias=" (" & saldoferias & ") " & dtinicf & " a " & dtvencf & " <br>&nbsp;&nbsp;(" & resto & ") " & dtinicf2 & " a " & dtvencf2 
end if

desc_rubrica=replace(desc_rubrica, "[perfim]", stringferias)
if rs2("rubrica")="95" then apendice="." & t95:t95=t95+1
if rs2("rubrica")="115" then apendice="." & t115:t115=t115+1

teste=rs2.fields(5)
if request.form("zero")="Z" and valor=0 then 
	impressao="&nbsp;"
else
	impressao=formatnumber(valor,2)
end if
%>
	<td class=rcampo3 valign=top>&nbsp;<%=rs2("rubrica")& apendice & " " & desc_rubrica%></td>
	<td class=rcampo4 valign=top align="right">&nbsp;<%=impressao%>&nbsp;</td>
<%
if rs2("agrupar")=2 then contaded=contaded+1
if rs2("agrupar")=1 then contaprov=contaprov+1
contacols=contacols+1
if contacols=3 then contacols=0:response.write "</tr>"
ulttipo=rs2("agrupar")
ultrubrica=rs2("rubrica")
rs2.movenext
loop
if contacols<3 then for a=contacols to 2:contaprov=contaprov+1:response.write "<td class=rcampo3></td><td class=rcampo4></td>":next

if totalizou=0 then 
	'if contacols<3 then for a=contacols to 2:response.write "<td class=campo></td><td class=campo></td>":next
	response.write "</tr>":contacols=0
	if contaprov<colsprov then
		linhatemp=(colsprov-contaprov)/3
		for l=1 to linhatemp
			response.write "<tr><td class=rcampo3></td><td class=rcampo4></td><td class=rcampo3></td><td class=rcampo4></td><td class=rcampo3></td><td class=rcampo4></td></tr>"
		next
	end if
	response.write "<tr height=30><td class=rcampo3></td><td class=rcampo4></td><td class=rcampo3></td><td class=rcampo4></td>"
	response.write "<td class=rtitulo2><b>" & "TOTAL BRUTO</td>"
	response.write "<td class=rtitulo2 align="right"><b>" & formatnumber(totprov,2) & "</td></tr>"
	response.write "<tr><td class=rtitulo3 align="left" colspan=6 valign="center"><b>&nbsp;DEDUÇÕES</td></tr>"
	response.write "<tr>"
	response.write "<td class=rtitulo3><b>&nbsp;Desconto</td><td class=rtitulo3><b>&nbsp;Valor</td>"
	response.write "<td class=rtitulo3><b>&nbsp;Desconto</td><td class=rtitulo3><b>&nbsp;Valor</td>"
	response.write "<td class=rtitulo3><b>&nbsp;Desconto</td><td class=rtitulo3><b>&nbsp;Valor</td>"
	response.write "</tr><tr height=30>"
end if

	if contaded<colsded then
		linhatemp=(colsded-contaded)/3
		for l=1 to linhatemp
			response.write "<tr><td class=rcampo3></td><td class=rcampo4></td><td class=rcampo3></td><td class=rcampo4></td><td class=rcampo3></td><td class=rcampo4></td></tr>"
		next
	end if
response.write "<tr height=30><td class=rcampo3></td><td class=rcampo4></td><td class=rcampo3></td><td class=rcampo4></td>"
response.write "<td class=rtitulo2><b>" & "TOTAL DEDUÇÕES</td>"
response.write "<td class=rtitulo2 align="right"><b>" & formatnumber(totdesc,2) & "</td></tr>"

response.write "<tr height=30><td class=rcampo3></td><td class=rcampo4></td><td class=rcampo3></td><td class=rcampo4></td>"
response.write "<td class=rtitulo2><b>VALOR LÍQUIDO</td>"
response.write "<td class=rtitulo2 align="right"><b>" & formatnumber(totprov-totdesc,2) & "</td></tr>"

response.write "<tr><td colspan=6 style=""border-bottom:1px solid"" height=1></td></tr>"

'--------------------------------------
%>
</table>
<DIV style="page-break-after:always"></DIV>

<%
if termo="Q" then
	titulo="TERMO DE QUITAÇÃO DE RESCISÃO DE CONTRATO DE TRABALHO"
else
	titulo="TERMO DE HOMOLOGAÇÃO DE RESCISÃO DE CONTRATO DE TRABALHO"
end if
%>

<table border="0" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr><td class=rtitulo1 align="center" valign="center" height=30><b><%=titulo%></td></tr>
<tr><td height="5pt"></td></tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr><td class=rtitulo2 colspan="2" align="left" valign="center" style="border: 1px solid #000000"><b>&nbsp;EMPREGADOR</td></tr>
<tr>
	<td class=rcampo1>&nbsp;01 CNPJ/CEI<br><span class=campo1>&nbsp;<%=rs("cgc")%></td>
	<td class=rcampo1>&nbsp;02 Razão Social/Nome<br><span class=campo1>&nbsp;FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</td>
</tr>
<tr><td class=rtitulo2 align="left" colspan=2 valign="center"><b>&nbsp;TRABALHADOR</td></tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=rcampo1 width=130>&nbsp;10 PIS/PASEP<br><span class=campo1>&nbsp;<%=rs("pispasep")%></td>
	<td class=rcampo1>&nbsp;11 Nome<br><span class=campo1>&nbsp;<%=rs("nome")%></b></td>
</tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=rcampo1 width=150>&nbsp;17 Carteira de Trabalho (nº, série, UF)<br>&nbsp;<span class=campo1><%=rs("carteiratrab") & " / " & rs("seriecarttrab") & " / " & rs("ufcarttrab")%></td>
	<td class=rcampo1>&nbsp;18 CPF<br>&nbsp;<span class=campo1><%=rs("cpf")%></td>
	<td class=rcampo1 width=130>&nbsp;19 Data de nascimento<br>&nbsp;<span class=campo1><%=dataform(rs("dtnascimento"))%></td>
	<td class=rcampo1>&nbsp;20 Nome da mãe<br>&nbsp;<span class=campo1><%=rs("mae")%></td>
</tr>
<tr><td class=rtitulo2 align="left" colspan=4 valign="center"><b>&nbsp;CONTRATO</td></tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr><td class=rcampo2>&nbsp;22 Causa do Afastamento<br>&nbsp;<span class=campo1><%=desc_%></td></tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=rcampo1>&nbsp;24 Data de Admissão<br>&nbsp;<span class=campo1><%=dataform(rs("admissao"))%></td>
	<td class=rcampo1>&nbsp;25 Data do Aviso Prévio<br>&nbsp;<span class=campo1><%=dataform(rs("dtavisoprevio"))%></td>
	<td class=rcampo1>&nbsp;26 Data de afastamento<br>&nbsp;<span class=campo1><%=dataform(rs("demissao"))%></td>
	<td class=rcampo1>&nbsp;27 Cód. Afastamento<br>&nbsp;<span class=campo1><%=cod_%></td>
	<td class=rcampo1>&nbsp;29 Pensão Alimentícia (%) (FGTS)<br>&nbsp;<span class=campo1><input type="text" class="form_input10" value="" size="4"></td>
</tr>
</table>

<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-bottom:1px solid" width="650">
<tr>
	<td class=rcampo1>&nbsp;30 Categoria do Trabalhador<br>&nbsp;<span class=campo1><%=numzero(rs("codcategoria"),2)%></td>
</tr>
</table>

<%if termo="H" then%>
<table border="1" bordercolor="#000000" cellpadding="0" cellspacing="0" style="border-collapse: collapse;border-top:0px transparent" width="650">
<tr>
	<td class=rcampo2>&nbsp;31 Código Sindical<br>&nbsp;<span class=campo1><%=cod_sind%></td>
	<td class=rcampo2>&nbsp;32 CNPJ e Nome da Entidade Sindical Laboral<br>&nbsp;<span class=campo1><%=cnpj_sind & " <span style='font-size:10pt'>" & nome_sind & "</span>"%></td>
</tr>
</table>
<%end if%>
<br>

<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=rcampo>
<%if termo="H" then%>
	Foi prestada, gratuitamente, assistência na rescisão do contrato de trabalho, nos termos do artigo n.º 477,
	§ 1º, da Consolidação das Leis do Trabalho (CLT), sendo comprovado neste ato o efetivo pagamento das verbas
	rescisórias especificadas no corpo do TRCT, no valor líquido de R$ <%=formatnumber(totprov-totdesc,2)%>, o qual, devidamente rubricado pelas partes, é parte integrante do
	presente Termo de Homologação.
	<br><br>
	As partes assistidas no presente ato de rescisão contratual foram identificadas como legítimas conforme previsto
	na Instrução Normativa/SRT n.º 15/2010.
	<br><br>
	Fica ressalvado o direito do trabalhador pleitear judicialmente os direitos informados no campo 155, abaixo.
	<br><br>
	____________________________/_____, _____ de _____________________________ de ________.
<%else%>
	Foi realizada a rescisão do contrato de trabalho do trabalho acima qualificado, nos termos do artigo n.º 477 
	da Consolidação das Leis do Trabalho (CLT). A assistência à rescisão prevista no § 1º do art. n.º 477 da CLT 
	não é devida, tendo em vista a duração do contrato de trabalho não ser superior a um ano de serviço e não existir 
	previsão de assistência à rescisão contratual em Acordo ou Convenção Coletiva de Trabalho da categoria a qual
	pertence o trabalhador.
	<br><br>
	No dia _____/_____/_______ foi realizado, nos termos do art. 23 da Instrução Normativa/SRT n.º 15/2010, o efetivo
	pagamento das verbas rescisórias especificadas no corpo do TRCT, no valor líquido de R$ <%=formatnumber(totprov-totdesc,2)%>, o qual, devidamente rubricado pelas partes, é
	parte integrante do presente Termo de Quitação.
	<br><br><br><br>
	____________________________/_____, _____ de _____________________________ de ________.
<%end if%>	
	</td>
</tr>
</table>


<%if termo="H" then%>
<br><br><br>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=rcampo2 valign=top style="border-top:1px solid" width=40%>150 Assinatura do Empregador ou Preposto</td>
	<td class=rcampo2 valign=top >&nbsp;</td>
</tr>
</table>
<br>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=rcampo2 valign=top style="border-top:1px solid" width=40%>151 Assinatura do Trabalhador</td>
	<td class=rcampo2></td>
	<td class=rcampo2 valign=top style="border-top:1px solid" width=40%>152 Assinatura do Responsável Legal do Trabalhador</td>
	<td class=rcampo2></td>
</tr>
</table>
<br>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=rcampo3 valign=top style="border-top:1px solid" width=40%>153 Carimbo e Assinatura do Assistente</td>
	<td class=rcampo3></td>
	<td class=rcampo3 valign=top style="border-top:1px solid" width=40%>154 Nome do Orgão Homologador</td>
	<td class=rcampo3></td>
</tr>
</table>
<%else%>
<br><br><br><br>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=rcampo2 valign=top style="border-top:1px solid" width=55%>150 Assinatura do Empregador ou Preposto</td>
	<td class=rcampo2 valign=top >&nbsp;</td>
</tr>
</table>
<br><br>
<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-bottom:0px transparent" width="650">
<tr>
	<td class=rcampo2 valign=top style="border-top:1px solid" width=50%>151 Assinatura do Trabalhador</td>
	<td class=rcampo2></td>
	<td class=rcampo2 valign=top style="border-top:1px solid" width=45%>152 Assinatura do Responsável Legal do Trabalhador</td>
</tr>
</table>
<%end if%>


<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;" width="650">
<tr>
<%if termo="H" then%>
	<td class=rrodape1 valign=top style="border:1px solid">155 Ressalvas</td>
<%else%>
	<td class=rrodape1 valign=top>&nbsp;</td>
<%end if%>
</tr>
<tr>
	<td class=rrodape2 valign=top style="border:1px solid;height:15pt">156 Informações à CAIXA:</td>
</tr>
</table>


<table border="0" bordercolor="#000000" cellpadding="2" cellspacing="0" style="border-collapse: collapse;border-top:0px transparent" width="650">
<tr><td class=rtitulo4 align="center" colspan=2 valign="center" align="center">
<b>A ASSISTÊNCIA NO ATO DE RESCISÃO CONTRATUAL É GRATUITA.<br>
<span style="font-size:8pt">
Pode o trabalhador iniciar ação judicial quanto aos créditos resultantes das relações de trabalho até o limite de 
dois anos após a extinção do contrato de trabalho (Inc. XXIX, Art. 7º da Constituição Federal/1988).
</span>
</td></tr>
</table>

<%
rs2.movefirst
totprov=0:totdesc=0
teste=0
if teste=1 then
%>
<table>
<%
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