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
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Recibo de Pagamento</title>
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
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rse=server.createobject ("ADODB.Recordset")
Set rse.ActiveConnection = conexao

if request.form<>"" then
	mes=request("mes"):ano=request("ano")
	periodo=request("periodo")
	temp=0
else
	temp=1
end if

if temp=1 then
datform=dateserial(year(now),month(now)-0,1)
mesform=month(datform)
anoform=year(datform)
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção de mês para impressão
<form method="POST" action="holerith2.asp" name="form">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width=300>
<tr>
	<td class=grupo align="center" >Mês</td>
	<td class=grupo align="center" >Ano</td>
	<td class=grupo align="center" >Periodo</td>
</tr>
<tr>
	<td class=fundo><input type="text" class="a" name="mes"  value="<%=mesform%>" size="2" maxlength="2"></td>
	<td class=fundo><input type="text" class="a" name="ano" value="<%=anoform%>" size="4" maxlength="4"></td>
	<td class=fundo><input type="text" class="a" name="periodo" value="2" size="4" maxlength="4"></td>
</tr>
<tr>
	<td class=grupo>Por ordem de:</td>
	<td class=grupo colspan=2>Ultimo impresso</td>
</tr>
<tr>
	<td class=fundo><input type="radio" name="ordem" value="nome" <%if session("ultimohtp")="nome" then response.write "checked"%> > nome<br>
	<input type="radio" name="ordem" value="chapa" <%if session("ultimohtp")="chapa" then response.write "checked"%> > chapa</td>
	<td class=fundo colspan=2><input type="text" class="a" name="ultimo"  value="<%=session("ultimohol")%>" size="20" > </td>
</tr>
<tr>
	<td class=fundo colspan=3 align="center"><input type="submit" value="Visualizar para imprimir" name="B1" class="button"></td>
</tr>
</table>
</form>
<!-- fim formulario -->
<%
elseif temp=0 then
'if request.form<>"" then
if mes-1<=0 then
	anoant=ano-1
	mesant=12
else
	anoant=ano
	mesant=mes-1
end if
if mes=8 then mesant2=6 else mesant2=mesant
sessao=session.sessionid:sessao=session("usuariomaster")
sql2="delete from temp_hol2 where sessao='" & sessao & "'"
sql2="delete from temp_hol2 "
conexao.execute sql2:
sql1="INSERT INTO temp_hol2 ( sessao, chapa, codevento, provdescbase, descricao, r1, v1 ) " & _
"SELECT '" & sessao & "', ff.CHAPA, ff.CODEVENTO, [e].PROVDESCBASE, [e].DESCRICAO, ff.REF, ff.VALOR " & _
"FROM (corporerm.dbo.PFFINANC AS ff INNER JOIN corporerm.dbo.PEVENTO AS e ON ff.CODEVENTO = [e].CODIGO) INNER JOIN corporerm.dbo.PFUNC AS f ON ff.CHAPA = f.CHAPA " & _
"WHERE ff.ANOCOMP=" & ano & " AND ff.MESCOMP=" & mes & " and ff.nroperiodo=" & periodo & " and ff.valor>0 "
conexao.execute sql1
sql1="INSERT INTO temp_hol2 ( sessao, chapa, codevento, provdescbase, descricao, r2, v2 ) " & _
"SELECT '" & sessao & "', ff.CHAPA, ff.CODEVENTO, [e].PROVDESCBASE, [e].DESCRICAO, ff.REF, ff.VALOR " & _
"FROM (corporerm.dbo.PFFINANC AS ff INNER JOIN corporerm.dbo.PEVENTO AS e ON ff.CODEVENTO = [e].CODIGO) INNER JOIN corporerm.dbo.PFUNC AS f ON ff.CHAPA = f.CHAPA " & _
"WHERE ff.ANOCOMP=" & anoant & " AND ff.MESCOMP=" & mesant & " and ff.nroperiodo=" & periodo & " AND f.CODSINDICATO<>'03'  and ff.valor>0 "
conexao.execute sql1
sql1="INSERT INTO temp_hol2 ( sessao, chapa, codevento, provdescbase, descricao, r2, v2 ) " & _
"SELECT '" & sessao & "', ff.CHAPA, ff.CODEVENTO, [e].PROVDESCBASE, [e].DESCRICAO, ff.REF, ff.VALOR " & _
"FROM (corporerm.dbo.PFFINANC AS ff INNER JOIN corporerm.dbo.PEVENTO AS e ON ff.CODEVENTO = [e].CODIGO) INNER JOIN corporerm.dbo.PFUNC AS f ON ff.CHAPA = f.CHAPA " & _
"WHERE ff.ANOCOMP=" & anoant & " AND ff.MESCOMP=" & mesant2 & " and ff.nroperiodo=" & periodo & " AND f.CODSINDICATO='03'  and ff.valor>0 "
conexao.execute sql1

sqla="SELECT top 100 f.NOME, f.CHAPA, f.CODSECAO, f.CODAGENCIAPAGTO, f.CONTAPAGAMENTO, " & _
"s.DESCRICAO AS secao, c.NOME AS funcao, s.CGC, c.CBO2002, p.ANOCOMP, p.MESCOMP, p.NROPERIODO " & _
"FROM corporerm.dbo.PFUNC AS f, corporerm.dbo.PSECAO AS s, corporerm.dbo.PFUNCAO AS c, corporerm.dbo.PFPERFF AS p " & _
"WHERE f.CODSECAO=s.CODIGO AND c.CODIGO=f.codfuncao AND p.CHAPA=f.chapa " & _
"AND p.ANOCOMP=" & ano & " AND p.MESCOMP=" & mes & " and p.nroperiodo=2 "

if request.form("ordem")="chapa" then 
	sqla=sqla & " and f.chapa>'" & request.form("ultimo") & "' "
	sql1=sqla & " order by f.chapa "
else
	sqla=sqla & " and f.nome>'" & request.form("ultimo") & "' "
	sql1=sqla & " order by f.nome "
end if
rs.Open sql1, ,adOpenStatic, adLockReadOnly
session("chapa")=rs("chapa")

if mes-1<=0 then
	anoant=ano-1
	mesant=12
else
	anoant=ano
	mesant=mes-1
end if
imprime=0:paginai=0
rs.movefirst
do while not rs.eof
tvatu=0:tvant=0:tdatu=0:tdant=0
%>
<!-- table pagina -->
<%
if imprime=0 then
%>
<table border="0" width=660 height="1000">
<tr><td valign="top" class=campo height=500 style="border-top:#000000 dotted 2">
<%
else
%>
<tr><td valign="top" class=campo height=500 style="border-top:#000000 dotted 2;border-bottom:#000000 dotted 2">
<%
end if
%>
<!-- inicio holerith -->
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="650">
	<tr>
		<td><b>FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</b></td>
		<td width=250 align="right"><b>Recibo de Pagamento de Salário</b></td>
	</tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="650">
	<tr>
		<td width=150><%=rs("cgc")%></td>
		<td ><%=rs("funcao")%></td>
		<td ><%=rs("secao")%></td>
	</tr>
</table>
<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="650">
	<tr>
		<td class="campor">Código</td>
		<td class="campor">Nome do Funcionário</td>
		<td class="campor">CBO</td>
		<td class="campor">Seção</td>
	</tr>
	<tr>
		<td><%=rs("chapa")%></td>
		<td><b><%=rs("nome")%></b></td>
		<td><%=rs("cbo2002")%></td>
		<td><%=rs("codsecao")%></td>
	</tr>
</table>

<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="690">
	<tr>
		<td class=campo align="center" rowspan=2>Cód.</td>
		<td class=campo align="center" rowspan=2>Descrição</td>
		<td class=campo align="center" colspan=3 style="border-right:#000000 solid 2"><b><%=mesant&"/"&anoant%></td>
		<td class=campo align="center" colspan=3 style="border-right:#000000 solid 2"><b><%=mes&"/"&ano%></td>
		<td class=campo align="center" rowspan=2>Var.</td>
	</tr>
	<tr>
		<td class="campor" align="center">Referência</td>
		<td class="campor" align="center">Vencimentos</td>
		<td class="campor" align="center" style="border-right:#000000 solid 2">Descontos</td>
		<td class="campor" align="center">Referência</td>
		<td class="campor" align="center">Vencimentos</td>
		<td class="campor" align="center" style="border-right:#000000 solid 2">Descontos</td>
	</tr>
<%
'bases - não estou utilizando.....
sqlb="SELECT p.BASEINSS, p.BASEINSS13, p.BASEIRRF, p.BASEIRRF13, p.INSSCAIXA, " & _
"p.BASEFGTS, p.BASEFGTS13, p.salariodecalculo " & _
"FROM PFPERFF AS p " & _
"WHERE p.ANOCOMP=" & ano & " AND p.MESCOMP=" & mes & " AND p.chapa='" & rs("chapa") & "' "

sqle="SELECT sessao, chapa, provdescbase, codevento, descricao, " & _
"Sum(r2) AS rant, Sum(v2) AS vant, Sum(r1) AS ratu, Sum(v1) AS vatu " & _
"FROM temp_hol2 " & _
"GROUP BY sessao, chapa, provdescbase, codevento, descricao " & _
"HAVING sessao='" & sessao & "' AND chapa='" & rs("chapa") & "' AND provdescbase<>'B' " & _
"ORDER BY provdescbase DESC , codevento "

rse.Open sqle, ,adOpenStatic, adLockReadOnly

if rse.recordcount>0 then
linhah=0:'datapagamento=rse("dtpagto")
rse.movefirst
do while not rse.eof
datu="&nbsp":dant="&nbsp":vatu="&nbsp":vant="&nbsp"
if rse("rant")=0 or isnull(rse("rant")) then rant="&nbsp;" else rant=formatnumber(rse("rant"),2)
if rse("ratu")=0 or isnull(rse("ratu")) then ratu="&nbsp;" else ratu=formatnumber(rse("ratu"),2)
if rse("provdescbase")="D" then 
	if isnull(rse("vatu")) then datu="&nbsp" else datu=formatnumber(rse("vatu"),2)
	if isnull(rse("vant")) then dant="&nbsp" else dant=formatnumber(rse("vant"),2)
	if isnull(rse("vatu")) then tdatu=tdatu else tdatu=tdatu+cdbl(rse("vatu"))
	if isnull(rse("vant")) then tdant=tdant else tdant=tdant+cdbl(rse("vant"))
	if datu="&nbsp" then v1=0 else v1=cdbl(rse("vatu"))
	if dant="&nbsp" then v2=0 else v2=cdbl(rse("vant"))
else 
	'descontos="&nbsp;"
end if
if rse("provdescbase")="P" then 
	if isnull(rse("vatu")) then vatu="&nbsp" else vatu=formatnumber(rse("vatu"),2)
	if isnull(rse("vant")) then vant="&nbsp" else vant=formatnumber(rse("vant"),2)
	if isnull(rse("vatu")) then tvatu=tvatu else tvatu=tvatu+cdbl(rse("vatu"))
	if isnull(rse("vant")) then tvant=tvant else tvant=tvant+cdbl(rse("vant"))
	if vatu="&nbsp" then v1=0 else v1=cdbl(rse("vatu"))
	if vant="&nbsp" then v2=0 else v2=cdbl(rse("vant"))
else 
	'vencimentos="&nbsp;"
end if
variacao=v1-v2
percvar="":if variacao>0 and rse("provdescbase")="P" and v1>0 and v2>0 then percvar=formatpercent(variacao/v2,2) else percvar="&nbsp;"
if variacao=0 then variacao="&nbsp" else variacao=formatnumber(variacao,2)

%>
	<tr>
		<td class=campo><%=rse("codevento")%></td>
		<td class=campo><%=rse("descricao")%></td>
		<td align="right" class=campo><%=rant%>&nbsp;</td>
		<td align="right" class=campo><%=vant%>&nbsp;</td>
		<td align="right" class=campo style="border-right:#000000 solid 2"><%=dant%>&nbsp;</td>
		<td align="right" class=campo><%=ratu%>&nbsp;</td>
		<td align="right" class=campo><%=vatu%>&nbsp;</td>
		<td align="right" class=campo style="border-right:#000000 solid 2"><%=datu%>&nbsp;</td>
		<td align="right" class="campor" nowrap><%=variacao%>&nbsp;<%=percvar%></td>
	</tr>
<%
linhah=linhah+1
rse.movenext
loop
	if linhah<15 then
		for a=1 to (15-linhah)
			response.write "<tr>"
			response.write "<td class=campo>&nbsp;</td>"
			response.write "<td class=campo>&nbsp;</td>"
			response.write "<td class=campo align=""right"">&nbsp;</td>"
			response.write "<td class=campo align=""right"">&nbsp;</td>"
			response.write "<td class=campo align=""right"" style='border-right:#000000 solid 2'>&nbsp;</td>"
			response.write "<td class=campo align=""right"">&nbsp;</td>"
			response.write "<td class=campo align=""right"">&nbsp;</td>"
			response.write "<td class=campo align=""right"" style='border-right:#000000 solid 2'>&nbsp;</td>"
			response.write "<td class=""campor"" align=""right"">&nbsp;</td>"
			response.write "</tr>"
		next
	end if

tlatu=tvatu-tdatu
tlant=tvant-tdant
%>
	<tr>
		<td colspan=3 rowspan=3 valign=top>
		Depósito na Agência: <%=rs("codagenciapagto")%><br>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		Conta: <%=rs("contapagamento")%><br>
		Data de crédito: <%=datapagamento%>
		</td>
		<td class="campor" align="center">Total de Vencimentos</td>
		<td class="campor" align="center" style="border-right:#000000 solid 2">Total de Descontos</td>
		<td rowspan=3></td>
		<td class="campor" align="center">Total de Vencimentos</td>
		<td class="campor" align="center" style="border-right:#000000 solid 2">Total de Descontos</td>
		<td class="campor">&nbsp;</td>
	</tr>
	<tr>
		<td align="right"><%=formatnumber(tvant,2)%>&nbsp;</td>
		<td align="right" style="border-right:#000000 solid 2"><%=formatnumber(tdant,2)%>&nbsp;</td>
		<td align="right"><%=formatnumber(tvatu,2)%>&nbsp;</td>
		<td align="right" style="border-right:#000000 solid 2"><%=formatnumber(tdatu,2)%>&nbsp;</td>
		<td class="campor">&nbsp;</td>
	</tr>
	<tr>
		<td class="campor" align="center" valign="center" >Valor Líquido <img src="../images/arrow.gif" border="0" width="13" height="10" alt=""></td>
		<td align="right" style="border-right:#000000 solid 2"><%=formatnumber(tlant,2)%>&nbsp;</td>
		<td class="campor" align="center" valign="center" >Valor Líquido <img src="../images/arrow.gif" border="0" width="13" height="10" alt=""></td>
		<td align="right" style="border-right:#000000 solid 2"><%=formatnumber(tlatu,2)%>&nbsp;</td>
		<td class="campor">&nbsp;</td>
	</tr>
</table>

<%
else 'rse.recordcount
	for a=1 to 15
		response.write "<tr>"
		response.write "<td width=50>&nbsp;</td>"
		response.write "<td width=300>&nbsp;</td>"
		response.write "<td width=70 align=""right"">&nbsp;</td>"
		response.write "<td align=""right"">&nbsp;</td>"
		response.write "<td align=""right"">&nbsp;</td>"
		response.write "<td align=""right"">&nbsp;</td>"
		response.write "<td align=""right"">&nbsp;</td>"
		response.write "<td align=""right"">&nbsp;</td>"
		response.write "<td align=""right"">&nbsp;</td>"
		response.write "</tr>"
	next
end if
rse.close

'if not isnull(rs("salariodecalculo")) then 
'	if cdbl(rs("salariodecalculo"))>0 then salariobase=formatnumber(rs("salariodecalculo"),2) else salariobase="&nbsp;"
'end if
'sqlbase="select max(c.limitesuperior) as baseinss from pcalcvlr c, ptabcalc t " & _
'"where t.iniciovigencia=c.iniciovigencia and t.codigo=c.codtabcalc " & _
'"and c.codtabcalc='01' and '" & dtaccess(dateserial(ano,mes,1)) & "' between t.iniciovigencia and t.finalvigencia "
'rse.Open sqlbase, ,adOpenStatic, adLockReadOnly
'baseinss=cdbl(rse("baseinss"))
'baseinssh=cdbl(rs("baseinss"))+cdbl(rs("baseinss13"))
'if baseinssh>baseinss then basei=baseinss else basei=baseinssh
'basei=formatnumber(basei,2)
'basefgts=cdbl(rs("basefgts"))+cdbl(rs("basefgts13"))
'fgtsmes=int(basefgts*8)/100
'basefgts=formatnumber(basefgts,2)
'fgtsmes=formatnumber(fgtsmes,2)
'baseirrf=cdbl(rs("baseirrf"))+cdbl(rs("baseirrf13"))-cdbl(rs("insscaixa"))
'sqldep="select valor from pvalfix " & _
'"where '" & dtaccess(dateserial(ano,mes,1)) & "' between iniciovigencia and finalvigencia and codigo='04'"
'rse.close
'rse.Open sqldep, ,adOpenStatic, adLockReadOnly
'valordep=cdbl(rse("valor"))
'sqlqt="select nrodependirrf as ndep " & _
'"from pfhstndp d, (select max(dtmudanca) as mdata from pfhstndp where chapa='" & rs("chapa") & "' and dtmudanca<='" & dtaccess(dateserial(ano,mes,1)) & "') t " & _
'"where chapa='" & rs("chapa") & "' and dtmudanca=t.mdata"
'rse.close
'rse.Open sqlqt, ,adOpenStatic, adLockReadOnly
'if rse.recordcount=0 then
'	ndep=0
'else
'	ndep=cdbl(rse("ndep"))
'end if
'rse.close
'deducao=valordep * ndep
'baseirrf=baseirrf-deducao
'baseirrf=formatnumber(baseirrf,2)

%>
<!-- bases -->
<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="650">
	<tr>
		<td class="campor">Salário Base</td>
		<td class="campor">Sal. Contr. INSS</td>
		<td class="campor">Base Cálc. FGTS</td>
		<td class="campor">F.G.T.S. do mês</td>
		<td class="campor">Base Cálc. IRRF</td>
		<td class="campor">Faixa IRRF</td>
	</tr>
	<tr>
		<td align="center">&nbsp;<%=salariobase%></td>
		<td align="center">&nbsp;<%=basei%></td>
		<td align="center">&nbsp;<%=basefgts%></td>
		<td align="center">&nbsp;<%=fgtsmes%></td>
		<td align="center">&nbsp;<%=baseirrf%></td>
		<td align="center">&nbsp;<%paginai=paginai+1:resto=paginai mod 2: if resto=0 then response.write int(paginai/2)%></td>
	</tr>
</table>
<!-- fim holerith -->

<%
if imprime=0 then
%>
</td></tr>
<!-- table pagina -->
<%
else
%>
</td></tr></table>
<%
if rs.absoluteposition<rs.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página -->
end if
if imprime=0 then imprime=1 else imprime=0

if request.form("ordem")="chapa" then
	session("ultimohol")=rs("chapa") 
	session("ultimohtp")="chapa"
else 
	session("ultimohol")=rs("nome")
	session("ultimohtp")="nome"
end if

rs.movenext
loop

rs.close
set rs=nothing
set rse=nothing
end if ' temps

conexao.close
set conexao=nothing
%>
</body>
</html>