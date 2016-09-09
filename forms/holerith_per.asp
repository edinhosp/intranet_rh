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

if request.form<>"" then
	chapa=request.form("chapa")
	if len(chapa)<5 then chapa=numzero(chapa,5)
	mes=request.form("mes"):ano=request.form("ano")
	mes2=request.form("mes2"):ano2=request.form("ano2")
	tiposaida=request.form("saida")
	
sqla="SELECT f.NOME, f.CHAPA, f.CODSECAO, f.CODAGENCIAPAGTO, f.CONTAPAGAMENTO, " & _
"s.DESCRICAO AS secao, c.NOME AS funcao, s.CGC, c.CBO2002, p.ANOCOMP, p.MESCOMP, " & _
"p.NROPERIODO, p.BASEINSS, p.BASEINSS13, p.BASEIRRF, p.BASEIRRF13, p.INSSCAIXA, " & _
"p.BASEFGTS, p.BASEFGTS13, p.salariodecalculo " & _
"FROM corporerm.dbo.PFUNC AS f, corporerm.dbo.PSECAO AS s, corporerm.dbo.PFUNCAO AS c, (select * from corporerm.dbo.pfperff union all select * from corporerm.dbo.pfperffcompl) AS p " & _
"WHERE f.CODSECAO=s.CODIGO AND c.CODIGO=f.codfuncao AND p.CHAPA=f.chapa " & _
"and CONVERT(smalldatetime,convert(nvarchar,anocomp)+case when mescomp<10 then '0'+convert(nvarchar,mescomp) else convert(nvarchar,mescomp) end+'01') " & _
"between CONVERT(smalldatetime,convert(nvarchar," & ano & ")+case when " & mes & "<10 then '0'+convert(nvarchar," & mes & ") else convert(nvarchar," & mes & ") end+'01') " & _
"and CONVERT(smalldatetime,convert(nvarchar," & ano2 & ")+case when " & mes2 & "<10 then '0'+convert(nvarchar," & mes2 & ") else convert(nvarchar," & mes2 & ") end+'01') " & _
"and f.CHAPA='" & chapa & "' "
	sql1=sqla & " order by f.nome, p.anocomp, p.mescomp "
	set rs=server.createobject ("ADODB.Recordset")
	Set rs.ActiveConnection = conexao
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	session("chapa")=rs("chapa")
	temp=0
	'if rs.recordcount>1 then temp=2
else
	temp=1
end if
%>
<%
if temp=1 then
datform=dateserial(year(now),month(now)-1,1)
mesform=month(datform)
anoform=year(datform)
if request.form("saida")="" then tiposaida="I" else tiposaida=request.form("saida")
%>
<form method="POST" action="holerith_per.asp" name="form">
<table border="1" cellpadding="2" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=titulo colspan=3>Seleções para emissão de Holerith por Período</td>
</tr>
<tr>
	<td class="campoa">Chapa: <input type="text" name="chapa" SIZE="5" value="<%=session("chapa")%>"></td>
	<td class="campoa">De: Mês <input type="text" class="a" name="mes"  value="<%=mesform+1%>" size="2" maxlength="2">
  Ano  <input type="text" class="a" name="ano" value="<%=anoform-1%>" size="4" maxlength="4"></td>
	<td class="campoa">Até: Mês <input type="text" class="a" name="mes2"  value="<%=mesform%>" size="2" maxlength="2">
  Ano  <input type="text" class="a" name="ano2" value="<%=anoform%>" size="4" maxlength="4"></td>
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
<!-- fim formulario -->
<%
elseif temp=0 then
'if request.form<>"" then
session("chapa")=rs("chapa")
session("chapanome")=rs("nome")

if tiposaida="I" then
'------------------------------------------------------
rs.movefirst
do while not rs.eof
'------------------------------------------------------
%>
<!-- inicio holerith -->

<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="630px">
	<tr><td><b>FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</b></td>
		<td width=250 align="right"><b>Recibo de Pagamento de Salário</b></td></tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="630px">
	<tr><td width=150><%=rs("cgc")%></td>
		<td ><%=rs("funcao")%></td>
		<td width=80 align="right"><%=numzero(rs("mescomp"),2) & "/" & rs("anocomp")%></td>
	</tr>
</table>
<table border="0" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width="630px">
	<tr><td width=150>&nbsp;</td>
		<td ><%=rs("secao")%></td>
		<td width=80>&nbsp;</td>
	</tr>
</table>

<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="630px">
	<tr><td class="campor">Código</td>
		<td class="campor">Nome do Funcionário</td>
		<td class="campor">CBO</td>
		<td class="campor">Seção</td>
	</tr>
	<tr><td><%=rs("chapa")%></td>
		<td><b><%=rs("nome")%></b></td>
		<td><%=rs("cbo2002")%></td>
		<td><%=rs("codsecao")%></td>
	</tr>
</table>

<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="630px">
	<tr><td class="campor" align="center">Cód.</td>
		<td class="campor" align="center">Descrição</td>
		<td class="campor" align="center">Referência</td>
		<td class="campor" align="center">Vencimentos</td>
		<td class="campor" align="center">Descontos</td>
	</tr>
<%
tvencimentos=0:tdescontos=0
salariobase=0:basei=0:basefgts=0:fgtsmes=0:baseirrf=0:baseinss=0

sqle="SELECT ff.CHAPA, ff.ANOCOMP, ff.MESCOMP, ff.NROPERIODO, e.PROVDESCBASE, ff.CODEVENTO, e.DESCRICAO, ff.REF, ff.VALOR, ff.dtpagto " & _
"FROM (select * from corporerm.dbo.pffinanc union all select * from corporerm.dbo.pffinanccompl) AS ff INNER JOIN corporerm.dbo.PEVENTO AS e ON ff.CODEVENTO = e.CODIGO " & _
"WHERE ff.CHAPA='" & rs("chapa") & "' AND ff.ANOCOMP=" & rs("anocomp") & " AND ff.MESCOMP=" & rs("mescomp") & " AND " & _
"e.PROVDESCBASE<>'B' AND ff.VALOR<>0 and ff.nroperiodo=" & rs("nroperiodo") & " " &  _
"ORDER BY ff.NROPERIODO, e.PROVDESCBASE DESC , ff.CODEVENTO "
set rse=server.createobject ("ADODB.Recordset")
Set rse.ActiveConnection = conexao
rse.Open sqle, ,adOpenStatic, adLockReadOnly

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
		<td class="campo">&nbsp;<%=rse("codevento")%></td>
		<td class="campo">&nbsp;<%=rse("descricao")%></td>
		<td class="campo" align="right"><%=ref%>&nbsp;</td>
		<td class="campo" align="right"><%=vencimentos%>&nbsp;</td>
		<td class="campo" align="right"><%=descontos%>&nbsp;</td>
	</tr>
<%
linhah=linhah+1
rse.movenext
loop
	if linhah<15 then
		for a=1 to (15-linhah)
			response.write "<tr><td>&nbsp;</td><td>&nbsp;</td><td align=""right"">&nbsp;</td><td align=""right"">&nbsp;</td><td align=""right"">&nbsp;</td></tr>"
		next
	end if

tliquido=tvencimentos-tdescontos
%>
	<tr>
		<td colspan="3" rowspan="3" valign="top" class="campo" >
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
		<td class="campo" align="right"><%=formatnumber(tvencimentos,2)%>&nbsp;</td>
		<td class="campo" align="right"><%=formatnumber(tdescontos,2)%>&nbsp;</td>
	</tr>
	<tr>
		<td class="campor" align="center" valign="center" >Valor Líquido <img src="../images/arrow.gif" border="0" width="13" height="10" alt=""></td>
		<td class="campo" align="right"><%=formatnumber(tliquido,2)%>&nbsp;</td>
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
		response.write "</tr>"
	next
end if
rse.close

if not isnull(rs("salariodecalculo")) then 
	if cdbl(rs("salariodecalculo"))>0 then salariobase=formatnumber(rs("salariodecalculo"),2) else salariobase="&nbsp;"
end if
sqlbase="select max(c.limitesuperior) as baseinss from corporerm.dbo.pcalcvlr c, corporerm.dbo.ptabcalc t " & _
"where t.iniciovigencia=c.iniciovigencia and t.codigo=c.codtabcalc " & _
"and c.codtabcalc='01' and '" & dtaccess(dateserial(rs("anocomp"),rs("mescomp"),1)) & "' between t.iniciovigencia and t.finalvigencia "
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
"where '" & dtaccess(dateserial(year(datapagamento),month(datapagamento),1)) & "' between iniciovigencia and finalvigencia and codigo='04'"
rse.close
rse.Open sqldep, ,adOpenStatic, adLockReadOnly
if rse.recordcount=0 then valordep=0 else valordep=cdbl(rse("valor"))
rse.close
sqlqt="select nrodependirrf as ndep " & _
"from corporerm.dbo.pfhstndp d, (select max(dtmudanca) as mdata from corporerm.dbo.pfhstndp where chapa='" & rs("chapa") & "' and dtmudanca<='" & dtaccess(dateserial(rs("anocomp"),rs("mescomp"),1)) & "') t " & _
"where chapa='" & rs("chapa") & "' and dtmudanca=t.mdata"
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
<table border="0" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="630px">
	<tr><td class="campo" style="border-right:1px solid #CCCCCC">Salário Base</td>
		<td class="campo" style="border-right:1px solid #CCCCCC">Sal. Contr. INSS</td>
		<td class="campo" style="border-right:1px solid #CCCCCC">Base Cálc. FGTS</td>
		<td class="campo" style="border-right:1px solid #CCCCCC">F.G.T.S. do mês</td>
		<td class="campo" style="border-right:1px solid #CCCCCC">Base Cálc. IRRF</td>
		<td class="campo" style="border-right:1px solid #CCCCCC">Faixa IRRF</td>
	</tr>
	<tr><td class="campo" align="center" style="border-right:1px solid #CCCCCC">&nbsp;<%=salariobase%></td>
		<td class="campo" align="center" style="border-right:1px solid #CCCCCC">&nbsp;<%=basei%></td>
		<td class="campo" align="center" style="border-right:1px solid #CCCCCC">&nbsp;<%=basefgts%></td>
		<td class="campo" align="center" style="border-right:1px solid #CCCCCC">&nbsp;<%=fgtsmes%></td>
		<td class="campo" align="center" style="border-right:1px solid #CCCCCC">&nbsp;<%=baseirrf%></td>
		<td class="campo" align="center" style="border-right:1px solid #CCCCCC">&nbsp;</td>
	</tr>
</table>
<br>
<br>
<DIV style="page-break-after:always"></DIV>

<!-- fim holerith -->
<%

'------------------------------------------------------
rs.movenext
loop
'------------------------------------------------------
end if 'tipo saida

if tiposaida="E" then

ini0="<html><head><meta http-equiv=""CONTENT-TYPE"" content=""text/html; charset=windows-1252"">" & _
"<title>Comprovante de Pagamento</title>" & _
"<link rel=""stylesheet"" type=""text/css"" href=""http://rh.unifieo.br/diversos.css"">" & _
"</head><body>"
ini0="<html><style type='text/css'>" & _
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

matrizcab=cint(rs.recordcount*5+1)
redim cab(matrizcab)
ncab=0
'------------------------------------------------------
rs.movefirst
do while not rs.eof
'------------------------------------------------------

cab(ncab)="<table border=""0"" cellpadding=""1"" cellspacing=""0"" style=""border-collapse: collapse"" width=""630px""> " & _
"<tr><td><b>Fundação Instituto de Ensino para Osasco</b></td><td width=250 align=""right""><b>Recibo de Pagamento de Salário</b></td></tr> " & _
"</table> " & _
"<table border=""0"" cellpadding=""1"" cellspacing=""0"" style=""border-collapse: collapse"" width=""630px""> " & _
"<tr><td width=150>" & rs("cgc") & "</td><td >" & rs("funcao") & "</td>" & _
"<td width=80 align=""right"">" & numzero(rs("mescomp"),2) & "/" & rs("anocomp") & "</td> " & _
"</tr></table> " & _
"<table border=""0"" cellpadding=""1"" cellspacing=""0"" style=""border-collapse: collapse"" width=""630px""> " & _
"<tr><td width=150>&nbsp;</td> " & _
"<td >" & rs("secao") & "</td><td width=80>&nbsp;</td></tr></table>" & _
"<table border=""0"" cellpadding=""1"" cellspacing=""1"" style=""border-collapse: collapse"" width=""630px""> " & _
"<tr><td class=""campor"">Código</td><td class=""campor"">Nome do Funcionário</td><td class=""campor"">CBO</td><td class=""campor"">Seção</td></tr> " & _
"<tr><td>" & rs("chapa") & "</td><td><b>" & rs("nome") & "</b></td><td>" & rs("cbo2002") & "</td><td>" & rs("codsecao") & "</td></tr></table>" & _
"<table border=""1"" bordercolor=""#CCCCCC"" cellpadding=""1"" cellspacing=""1"" style=""border-collapse: collapse"" width=""630px"">" & _
"<tr><td class=""campor"" align=""center"">Cód.</td><td class=""campor"" align=""center"">Descrição</td>" & _
"<td class=""campor"" align=""center"">Referência</td><td class=""campor"" align=""center"">Vencimentos</td>" & _
"<td class=""campor"" align=""center"">Descontos</td></tr>"
ncab=ncab+1

tvencimentos=0:tdescontos=0
salariobase=0:basei=0:basefgts=0:fgtsmes=0:baseirrf=0:baseinss=0

sqle="SELECT ff.CHAPA, ff.ANOCOMP, ff.MESCOMP, ff.NROPERIODO, e.PROVDESCBASE, ff.CODEVENTO, e.DESCRICAO, ff.REF, ff.VALOR, ff.dtpagto " & _
"FROM (select * from corporerm.dbo.pffinanc union all select * from corporerm.dbo.pffinanccompl) AS ff INNER JOIN corporerm.dbo.PEVENTO AS e ON ff.CODEVENTO = e.CODIGO " & _
"WHERE ff.CHAPA='" & rs("chapa") & "' AND ff.ANOCOMP=" & rs("anocomp") & " AND ff.MESCOMP=" & rs("mescomp") & " AND " & _
"e.PROVDESCBASE<>'B' AND ff.VALOR<>0 and ff.nroperiodo=" & rs("nroperiodo") & " " &  _
"ORDER BY ff.NROPERIODO, e.PROVDESCBASE DESC , ff.CODEVENTO "
set rse=server.createobject ("ADODB.Recordset")
Set rse.ActiveConnection = conexao
rse.Open sqle, ,adOpenStatic, adLockReadOnly

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

	cab(ncab)=cab(ncab) & "<tr><td class=""campo"">&nbsp;" & rse("codevento") & "</td><td class=""campo"">&nbsp;" & rse("descricao") & "</td>" & _
	"<td class=""campo"" align=""right"">" & ref & "&nbsp;</td><td class=""campo"" align=""right"">" & vencimentos & "&nbsp;</td>" & _
	"<td class=""campo"" align=""right"">" & descontos & "&nbsp;</td></tr>"
	linhah=linhah+1
	rse.movenext
	loop
	ncab=ncab+1
	if linhah<15 then
		for a=1 to (15-linhah)
			cab(ncab)=cab(ncab) & "<tr><td>&nbsp;</td><td>&nbsp;</td><td align=""right"">&nbsp;</td><td align=""right"">&nbsp;</td><td align=""right"">&nbsp;</td></tr>"
		next
		ncab=ncab+1
	end if

	tliquido=tvencimentos-tdescontos

	cab(ncab)="<tr><td colspan=""3"" rowspan=""3"" valign=""top"" class=""campo"" > " & _
	"Depósito na Agência: " & rs("codagenciapagto") & "<br> " & _
	"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; " & _
	"Conta: " & rs("contapagamento") & "<br>Data de crédito: " & datapagamento & "" & _
	"</td><td class=""campor"" align=""center"">Total de Vencimentos</td><td class=""campor"" align=""center"">Total de Descontos</td> " & _
	"</tr><tr><td class=""campo"" align=""right"">" & formatnumber(tvencimentos,2) & "&nbsp;</td> " & _
	"<td class=""campo"" align=""right"">" & formatnumber(tdescontos,2) & "&nbsp;</td></tr><tr> " & _
	"<td class=""campor"" align=""center"" valign=""center"" >Valor Líquido <img src=""http://rh.unifieo.br/images/arrow.gif"" border=""0"" width=""13"" height=""10"" alt=""""> </td>" & _
	"<td class=""campo"" align=""right"">" & formatnumber(tliquido,2) & "&nbsp;</td></tr></table>"
	ncab=ncab+1
	
else 'rse.recordcount
	cab(ncab)=""
	ncab=ncab+1
	for a=1 to 15
		cab(ncab)=cab(ncab) & "<tr><td width=50>&nbsp;</td><td width=300>&nbsp;</td><td width=70 align=""right"">&nbsp;</td><td align=""right"">&nbsp;</td><td align=""right"">&nbsp;</td></tr>"
	next
	ncab=ncab+1
end if
rse.close

if not isnull(rs("salariodecalculo")) then 
	if cdbl(rs("salariodecalculo"))>0 then salariobase=formatnumber(rs("salariodecalculo"),2) else salariobase="&nbsp;"
end if
sqlbase="select max(c.limitesuperior) as baseinss from corporerm.dbo.pcalcvlr c, corporerm.dbo.ptabcalc t " & _
"where t.iniciovigencia=c.iniciovigencia and t.codigo=c.codtabcalc " & _
"and c.codtabcalc='01' and '" & dtaccess(dateserial(rs("anocomp"),rs("mescomp"),1)) & "' between t.iniciovigencia and t.finalvigencia "
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
"where '" & dtaccess(dateserial(year(datapagamento),month(datapagamento),1)) & "' between iniciovigencia and finalvigencia and codigo='04'"
rse.close
rse.Open sqldep, ,adOpenStatic, adLockReadOnly
if rse.recordcount=0 then valordep=0 else valordep=cdbl(rse("valor"))
rse.close
sqlqt="select nrodependirrf as ndep " & _
"from corporerm.dbo.pfhstndp d, (select max(dtmudanca) as mdata from corporerm.dbo.pfhstndp where chapa='" & rs("chapa") & "' and dtmudanca<='" & dtaccess(dateserial(rs("anocomp"),rs("mescomp"),1)) & "') t " & _
"where chapa='" & rs("chapa") & "' and dtmudanca=t.mdata"
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

<!-- bases -->
cab(ncab)="<table border=""0"" cellpadding=""1"" cellspacing=""1"" style=""border-collapse: collapse"" width=""630px""> " & _
"<tr><td class=""campo"" style=""border-right:1px solid #CCCCCC"">Salário Base</td> " & _
"<td class=""campo"" style=""border-right:1px solid #CCCCCC"">Sal. Contr. INSS</td> " & _
"<td class=""campo"" style=""border-right:1px solid #CCCCCC"">Base Cálc. FGTS</td> " & _
"<td class=""campo"" style=""border-right:1px solid #CCCCCC"">F.G.T.S. do mês</td> " & _
"<td class=""campo"" style=""border-right:1px solid #CCCCCC"">Base Cálc. IRRF</td> " & _
"<td class=""campo"" style=""border-right:1px solid #CCCCCC"">Faixa IRRF</td></tr> " & _
"<tr><td class=""campo"" align=""center"" style=""border-right:1px solid #CCCCCC"">&nbsp;" & salariobase & "</td>" & _
"<td class=""campo"" align=""center"" style=""border-right:1px solid #CCCCCC"">&nbsp;" & basei & "</td>" & _
"<td class=""campo"" align=""center"" style=""border-right:1px solid #CCCCCC"">&nbsp;" & basefgts & "</td>" & _
"<td class=""campo"" align=""center"" style=""border-right:1px solid #CCCCCC"">&nbsp;" & fgtsmes & "</td>" & _
"<td class=""campo"" align=""center"" style=""border-right:1px solid #CCCCCC"">&nbsp;" & baseirrf & "</td>" & _
"<td class=""campo"" align=""center"" style=""border-right:1px solid #CCCCCC"">&nbsp;</td></tr></table>" & _
"<hr><br><br><DIV style=""page-break-after:always""></DIV>"
ncab=ncab+1

'------------------------------------------------------
rs.movenext
loop
'------------------------------------------------------
fim99="</body></html>"

Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory. 
Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network). 
Const cdoAnonymous = 0 'Do not authenticate
Const cdoBasic = 1 'basic (clear-text) authentication
Const cdoNTLM = 2 'NTLM

	sqlemail="select chapa, email, codsituacao from qry_funcionarios where chapa='" & chapa & "'"
	rse.Open sqlemail, ,adOpenStatic, adLockReadOnly
	email=rse("email")
	codsituacao=rse("codsituacao")
	rse.close
	email1=chapa & "@unifieo.br"
	
	if codsituacao="D" then email1=email
	if left(email1,1)>"0" then email1=email
	Set Mailer = CreateObject("CDO.Message")
	Mailer.From = "rh@unifieo.br" ' e-mail de quem esta enviando a mensagem 
	Mailer.To = email1 ' e-mail de quem vai receber a mensagem 
	if email<>"" and email<>email1 then Mailer.CC=email
	'Mailer.BCC = emailchefe ' Com Cópia 
	Mailer.ReplyTo = "rh@unifieo.br"
	'Mailer.AttachFile "e:\home\login\dados\arquivo.txt" 'caso queira anexar algum arquivo ao seu e-mail 
	Mailer.Subject = "Comprovante de Pagamento - " & mes & "/" & ano & " até " & mes2 & "/" & ano2
	'Mailer.TextBody = "Você tem mensagem" 
	Mailer.HtmlBody=ini0
	for a=0 to ncab
		Mailer.HtmlBody=Mailer.HtmlBody & cab(a)
	next
	Mailer.HtmlBody=Mailer.HtmlBody & fim99
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
'==End remote SMTP server configuration section==
	teste=0
	Mailer.Send 
	Set Mailer = Nothing 
	response.write "<p> Comprovante de Pagamento - " & mes & "/" & ano & " até " & mes2 & "/" & ano2
	response.write "<p> Enviado para " & email1 & " " & email

end if 'tipo saida

rs.close
set rs=nothing
set rse=nothing

end if ' temps
conexao.close
set conexao=nothing

%>
</body>
</html>