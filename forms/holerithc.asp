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
set rse=server.createobject ("ADODB.Recordset")
Set rse.ActiveConnection = conexao
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao

if request("codigo")<>"" or request.form<>"" then
	if request.form="" then temp=request("codigo"):mes=request("mes"):ano=request("ano"):periodo=request("periodo"):sqlc=" and p.nroperiodo=" & periodo 
	if request("codigo")="" then temp=request.form("coddoc"):mes=request.form("mes"):ano=request.form("ano")
	
sqla="SELECT f.NOME, f.CHAPA, f.CODSECAO, f.CODAGENCIAPAGTO, f.CONTAPAGAMENTO, " & _
"s.DESCRICAO AS secao, c.NOME AS funcao, s.CGC, c.CBO2002, p.ANOCOMP, p.MESCOMP, " & _
"p.NROPERIODO, p.BASEINSS, p.BASEINSS13, p.BASEIRRF, p.BASEIRRF13, p.INSSCAIXA, " & _
"p.BASEFGTS, p.BASEFGTS13, p.salariodecalculo " & _
"FROM corporerm.dbo.PFUNC AS f, corporerm.dbo.PSECAO AS s, corporerm.dbo.PFUNCAO AS c, (select * from corporerm.dbo.pfperff union all select * from corporerm.dbo.pfperffcompl) AS p " & _
"WHERE f.CODSECAO=s.CODIGO AND c.CODIGO=f.codfuncao AND p.CHAPA=f.chapa " & _
"AND p.ANOCOMP=" & ano & " AND p.MESCOMP=" & mes & " and nroperiodo=2 " & _
"and f.chapa in (select distinct chapa from corporerm.dbo.pffinanc f " & _
"	inner join g2cursoeve e on e.sal=f.codevento collate database_default " & _
"	where coddoc='" & request.form("coddoc") & "' and MESCOMP=" & mes & " and ANOCOMP=" & ano & ") " 

	sql1=sqla & " order by f.nome "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	session("chapa")=rs("chapa")
end if
%>
<%
dataform=dateserial(year(now),month(now)-1,1)
mesform=month(dataform)
anoform=year(dataform)

IF request.form="" then
%>
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleção de funcionário
<form method="POST" action="holerithc.asp" name="form">
  <p style="margin-top: 0; margin-bottom: 0">
  Curso <input type="text" name="coddoc" size="20" class="a" value="<%=session("coddoc")%>"><br>
  Mês <input type="text" class="a" name="mes"  value="<%=mesform%>" size="2" maxlength="2">
  Ano  <input type="text" class="a" name="ano" value="<%=anoform%>" size="4" maxlength="4">
      
  <input type="submit" value="Pesquisar" name="B1" class="button"></p>
</form>
<!-- fim formulario -->
<%
end if 'request.form=""

if request.form<>"" then
'if request.form<>"" then
session("coddoc")=request.form("coddoc")

if rs.recordcount>0 then
do while not rs.eof
tliquido=0
tvencimentos=0
tdescontos=0
sqle="SELECT ff.CHAPA, ff.ANOCOMP, ff.MESCOMP, ff.NROPERIODO, e.PROVDESCBASE, ff.CODEVENTO, e.DESCRICAO, ff.REF, ff.VALOR, ff.dtpagto " & _
"FROM (select * from corporerm.dbo.pffinanc union all select * from corporerm.dbo.pffinanccompl) AS ff INNER JOIN corporerm.dbo.PEVENTO AS e ON ff.CODEVENTO = e.CODIGO " & _
"WHERE ff.ANOCOMP=" & ano & " AND ff.MESCOMP=" & mes & " AND " & _
"e.PROVDESCBASE<>'B' AND ff.VALOR<>0 and ff.nroperiodo=" & rs("nroperiodo") & " " &  _
"AND ff.chapa='" & rs("chapa") & "' " & _
"ORDER BY ff.NROPERIODO, e.PROVDESCBASE DESC , ff.CODEVENTO "
rse.Open sqle, ,adOpenStatic, adLockReadOnly

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
		<td width=80 align="right"><%=numzero(mes,2) & "/" & ano%></td>
	</tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="650">
	<tr>
		<td width=150>&nbsp;</td>
		<td ><%=rs("secao")%></td>
		<td width=80>&nbsp;</td>
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

<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="650">
	<tr>
		<td class="campor" align="center">Cód.</td>
		<td class="campor" align="center">Descrição</td>
		<td class="campor" align="center">Referência</td>
		<td class="campor" align="center">Vencimentos</td>
		<td class="campor" align="center">Descontos</td>
	</tr>
<%
especial=0:divisor=2.2005
if rse.recordcount>0 then
linhah=0:datapagamento=rse("dtpagto")
rse.movefirst
do while not rse.eof
if rse("provdescbase")="D" then 
	if especial=1 and rse("codevento")="098" then
		valord=((60294.62)/divisor)*0.275-548.42
	else
		valord=rse("valor")
	end if
	descontos=formatnumber(valord,2)
	tdescontos=tdescontos+cdbl(valord)
else 
	descontos="&nbsp;"
end if
if rse("provdescbase")="P" then 
	if especial=1 then valorp=cdbl(rse("valor"))/divisor else valorp=rse("valor")
	vencimentos=formatnumber(valorp,2)
	tvencimentos=tvencimentos+cdbl(valorp)
else 
	vencimentos="&nbsp;"
end if
if rse("ref")="" or isnull(rse("ref")) then ref="" else ref=formatnumber(rse("ref"),2)
%>
	<tr>
		<td>&nbsp;<%=rse("codevento")%></td>
		<td>&nbsp;<%=rse("descricao")%></td>
		<td align="right"><%=ref%>&nbsp;</td>
		<td align="right"><%=vencimentos%>&nbsp;</td>
		<td align="right"><%=descontos%>&nbsp;</td>
	</tr>
<%
linhah=linhah+1
rse.movenext
loop
	if linhah<15 then
		for a=1 to (15-linhah)
			response.write "<tr>"
			response.write "<td>&nbsp;</td>"
			response.write "<td>&nbsp;</td>"
			response.write "<td align="right">&nbsp;</td>"
			response.write "<td align="right">&nbsp;</td>"
			response.write "<td align="right">&nbsp;</td>"
			response.write "</tr>"
		next
	end if

tliquido=tvencimentos-tdescontos
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
		<td class="campor" align="center">Total de Descontos</td>
	</tr>
	<tr>
		<td align="right"><%=formatnumber(tvencimentos,2)%>&nbsp;</td>
		<td align="right"><%=formatnumber(tdescontos,2)%>&nbsp;</td>
	</tr>
	<tr>
		<td class="campor" align="center" valign="center" >Valor Líquido <img src="../images/arrow.gif" border="0" width="13" height="10" alt=""></td>
		<td align="right"><%=formatnumber(tliquido,2)%>&nbsp;</td>
	</tr>
</table>
<%
else 'rse.recordcount
	for a=1 to 15
		response.write "<tr>"
		response.write "<td width=50>&nbsp;</td>"
		response.write "<td width=300>&nbsp;</td>"
		response.write "<td width=70 align="right">&nbsp;</td>"
		response.write "<td align="right">&nbsp;</td>"
		response.write "<td align="right">&nbsp;</td>"
		response.write "</tr>"
	next
end if
rse.close

if not isnull(rs("salariodecalculo")) then 
	if cdbl(rs("salariodecalculo"))>0 then salariobase=formatnumber(rs("salariodecalculo"),2) else salariobase="&nbsp;"
end if
sqlbase="select max(c.limitesuperior) as baseinss from corporerm.dbo.pcalcvlr c, corporerm.dbo.ptabcalc t " & _
"where t.iniciovigencia=c.iniciovigencia and t.codigo=c.codtabcalc " & _
"and c.codtabcalc='01' and '" & dtaccess(dateserial(ano,mes,1)) & "' between t.iniciovigencia and t.finalvigencia "
rse.Open sqlbase, ,adOpenStatic, adLockReadOnly
baseinss=cdbl(rse("baseinss"))
baseinssh=cdbl(rs("baseinss"))+cdbl(rs("baseinss13"))
if baseinssh>baseinss then basei=baseinss else basei=baseinssh
basei=formatnumber(basei,2)
basefgts=cdbl(rs("basefgts"))+cdbl(rs("basefgts13"))
fgtsmes=int(basefgts*8)/100
if especial=1 then basefgts=basefgts/divisor
basefgts=formatnumber(basefgts,2)
if especial=1 then fgtsmes=fgtsmes/divisor
fgtsmes=formatnumber(fgtsmes,2)
baseirrf=cdbl(rs("baseirrf"))+cdbl(rs("baseirrf13"))
if especial=1 then baseirrf=baseirrf/divisor
baseirrf=baseirrf-cdbl(rs("insscaixa"))
sqldep="select valor from corporerm.dbo.pvalfix " & _
"where '" & dtaccess(dateserial(ano,mes,1)) & "' between iniciovigencia and finalvigencia and codigo='04'"
rse.close
rse.Open sqldep, ,adOpenStatic, adLockReadOnly
valordep=cdbl(rse("valor"))
sqlqt="select nrodependirrf as ndep " & _
"from corporerm.dbo.pfhstndp d, (select max(dtmudanca) as mdata from corporerm.dbo.pfhstndp where chapa='" & rs("chapa") & "' and dtmudanca<='" & dtaccess(dateserial(ano,mes,1)) & "') t " & _
"where chapa='" & rs("chapa") & "' and dtmudanca=t.mdata"
rse.close
rse.Open sqlqt, ,adOpenStatic, adLockReadOnly
if rse.recordcount=0 then
	ndep=0
else
	ndep=cdbl(rse("ndep"))
end if
rse.close
deducao=valordep * ndep
baseirrf=baseirrf-deducao
baseirrf=formatnumber(baseirrf,2)

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
		<td align="center">&nbsp;</td>
	</tr>
</table>
<DIV style="page-break-after:always"></DIV>

<!-- fim holerith -->
<%
rs.movenext
loop


set rse=nothing

rs.close
set rs=nothing

end if
end if ' temps
conexao.close
set conexao=nothing

%>
</body>
</html>