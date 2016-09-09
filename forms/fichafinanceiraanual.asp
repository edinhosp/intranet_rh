<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=False
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a76")="N" or session("a76")="" then response.write "<script language='JavaScript' type='text/javascript'>self.close();</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Ficha Financeira Anual</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2, sal_anterior(10)
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rsq=server.createobject ("ADODB.Recordset")
Set rsq.ActiveConnection = conexao
set rse=server.createobject ("ADODB.Recordset")
Set rse.ActiveConnection = conexao

dim sbase(12), scinss(12), bcfgts(12), vfgts(12), bcirrf(12), agencia(12), conta(12)

if request.form<>"" then
	ano=request("ano")
	temp=0
else
	temp=1
end if

if temp=1 then
datform=dateserial(year(now),month(now)-0,1)
anoform=year(datform)-1
%>
<form method="POST" action="fichafinanceiraanual.asp" name="formff">
<input type="hidden" name="chapa" value="<%=chapa%>">
<input type="hidden" name="nomefunc" value="<%=nomefunc%>">
<p style="margin-top: 0; margin-bottom: 0" class=titulo>Seleções para Ficha Financeira
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width=300>
<tr>
	<td class=grupo align="center" >Ano</td>
	<td class=grupo>Por ordem de:</td>
	<td class=grupo colspan=2>Ultimo impresso</td>
</tr>
<tr>
	<td class=fundo><input type="text" class="a" name="ano" value="<%=anoform%>" size="4" maxlength="4"></td>
	<td class=fundo><input type="radio" name="ordem" value="nome" <%if session("ultimohtp")="nome" then response.write "checked"%> > nome<br>
	<input type="radio" name="ordem" value="chapa" <%if session("ultimohtp")="chapa" then response.write "checked"%> > chapa</td>
	<td class=fundo colspan=2><input type="text" class="a" name="ultimo"  value="<%=session("ultimohol")%>" size="20" > </td>
</tr>
<tr>
	<td class=fundo colspan=3 align="center"><input type="submit" value="Visualizar para imprimir" name="B1" class="button"></td>
</tr>
</table>

</form>
<%
elseif temp=0 then
	sessao=session("usuariomaster")
sqlf="select top 50 f.chapa, f.NOME, f.CODSECAO, f.SECAO, f.Funcao, f.CBO2002, f.cgc, f.CODBANCOPAGTO, f.CODAGENCIAPAGTO, f.CONTAPAGAMENTO " & _
"from qry_funcionarios f where f.chapa in (select distinct chapa from corporerm.dbo.PFFINANC where ANOCOMP=" & ano & ") "

if request.form("ordem")="chapa" then 
	sqlf=sqlf & " and f.chapa>'" & request.form("ultimo") & "' "
	sqlf=sqlf & " order by f.chapa "
else
	sqlf=sqlf & " and f.nome>'" & request.form("ultimo") & "' "
	sqlf=sqlf & " order by f.nome "
end if

rs2.Open sqlf, ,adOpenStatic, adLockReadOnly
do while not rs2.eof
totalano=0.00:inicio=1
	totalv1=0:totalv2=0:totalv3=0:totalv4=0:totalv5=0:totalv6=0:totalv7=0:totalv8=0:totalv9=0:totalv10=0:totalv11=0:totalv12=0:totalv=0
	totall1=0:totall2=0:totall3=0:totall4=0:totall5=0:totall6=0:totall7=0:totall8=0:totall9=0:totall10=0:totall11=0:totall12=0:totall=0
	totalp1=0:totalp2=0:totalp3=0:totalp4=0:totalp5=0:totalp6=0:totalp7=0:totalp8=0:totalp9=0:totalp10=0:totalp11=0:totalp12=0:totalp=0
	totald1=0:totald2=0:totald3=0:totald4=0:totald5=0:totald6=0:totald7=0:totald8=0:totald9=0:totald10=0:totald11=0:totald12=0:totald=0
%>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=999>
<th class=titulo colspan=5>Ficha Financeira - Anual</th>
<tr>
	<td class="campor">Empregador<br><b>FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</td>
	<td class="campor">CNPJ<br><b><%=rs2("cgc")%></td>
	<td class="campor">Seção<br><b><%=rs2("codsecao")%> - <%=rs2("secao")%></td>
	<td class="campor">Ano<br><b><%=ano%></td>
</tr>
</table>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=999>
<tr>
	<td class="campor">Chapa      <br><b><%=rs2("chapa")%></td>
	<td class="campor">Funcionário<br><b><%=rs2("nome")%></td>
	<td class="campor">Função     <br><b><%=rs2("funcao")%></td>
	<td class="campor">CBO        <br><b><%=rs2("cbo2002")%></td>
</tr>
</table>
<br>
<table border="1" bordercolor="#CCCCCC" cellpadding="1" cellspacing="0" style="border-collapse: collapse" width=999>
<tr>
	<td class=fundor rowspan=2>Cod.</td>	<td class=fundor rowspan=2>Descr.Evento</td>	<td class=fundor rowspan=2>Tipo</td>
	<td class=titulor colspan=2 align="center">Jan.</td>	<td class=titulor colspan=2 align="center">Fev.</td>
	<td class=titulor colspan=2 align="center">Mar.</td>	<td class=titulor colspan=2 align="center">Abr.</td>
	<td class=titulor colspan=2 align="center">Mai.</td>	<td class=titulor colspan=2 align="center">Jun.</td>
	<td class=titulor colspan=2 align="center">Jul.</td>	<td class=titulor colspan=2 align="center">Ago.</td>
	<td class=titulor colspan=2 align="center">Set.</td>	<td class=titulor colspan=2 align="center">Out.</td>
	<td class=titulor colspan=2 align="center">Nov.</td>	<td class=titulor colspan=2 align="center">Dez.</td>
	<td class=fundor rowspan=2>Total Ano</td>
</tr>
<tr>
	<td class=fundor>Ref.</td>	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>	<td class=fundor>Valor</td>
	<td class=fundor>Ref.</td>	<td class=fundor>Valor</td>
</tr>
<%

sqlv="SELECT ff.NROPERIODO, ff.CHAPA, ff.ANOCOMP, e.PROVDESCBASE, ff.CODEVENTO, e.DESCRICAO, " & _
"'r01'=sum(case ff.mescomp when 01 then case when ff.ref is null then 0 else ff.ref end else 0 end), 'v01'=sum(case ff.mescomp when 01 then ff.valor else 0 end), " & _
"'r02'=sum(case ff.mescomp when 02 then case when ff.ref is null then 0 else ff.ref end else 0 end), 'v02'=sum(case ff.mescomp when 02 then ff.valor else 0 end), " & _
"'r03'=sum(case ff.mescomp when 03 then case when ff.ref is null then 0 else ff.ref end else 0 end), 'v03'=sum(case ff.mescomp when 03 then ff.valor else 0 end), " & _
"'r04'=sum(case ff.mescomp when 04 then case when ff.ref is null then 0 else ff.ref end else 0 end), 'v04'=sum(case ff.mescomp when 04 then ff.valor else 0 end), " & _
"'r05'=sum(case ff.mescomp when 05 then case when ff.ref is null then 0 else ff.ref end else 0 end), 'v05'=sum(case ff.mescomp when 05 then ff.valor else 0 end), " & _
"'r06'=sum(case ff.mescomp when 06 then case when ff.ref is null then 0 else ff.ref end else 0 end), 'v06'=sum(case ff.mescomp when 06 then ff.valor else 0 end), " & _
"'r07'=sum(case ff.mescomp when 07 then case when ff.ref is null then 0 else ff.ref end else 0 end), 'v07'=sum(case ff.mescomp when 07 then ff.valor else 0 end), " & _
"'r08'=sum(case ff.mescomp when 08 then case when ff.ref is null then 0 else ff.ref end else 0 end), 'v08'=sum(case ff.mescomp when 08 then ff.valor else 0 end), " & _
"'r09'=sum(case ff.mescomp when 09 then case when ff.ref is null then 0 else ff.ref end else 0 end), 'v09'=sum(case ff.mescomp when 09 then ff.valor else 0 end), " & _
"'r10'=sum(case ff.mescomp when 10 then case when ff.ref is null then 0 else ff.ref end else 0 end), 'v10'=sum(case ff.mescomp when 10 then ff.valor else 0 end), " & _
"'r11'=sum(case ff.mescomp when 11 then case when ff.ref is null then 0 else ff.ref end else 0 end), 'v11'=sum(case ff.mescomp when 11 then ff.valor else 0 end), " & _
"'r12'=sum(case ff.mescomp when 12 then case when ff.ref is null then 0 else ff.ref end else 0 end), 'v12'=sum(case ff.mescomp when 12 then ff.valor else 0 end) " & _
"FROM (select * from corporerm.dbo.PFFINANC where chapa='" & rs2("chapa") & "' AND ANOCOMP=" & ano & " union all " & _
"select * from corporerm.dbo.PFFINANCCOMPL where chapa='" & rs2("chapa") & "' AND ANOCOMP=" & ano & ") ff " & _
"INNER JOIN corporerm.dbo.PEVENTO e ON ff.CODEVENTO = e.CODIGO " & _
"WHERE ff.CHAPA='" & rs2("chapa") & "' AND ff.ANOCOMP=" & ano & " and provdescbase<>'B' " & _
"GROUP BY ff.NROPERIODO, ff.CHAPA, ff.ANOCOMP, e.PROVDESCBASE, ff.CODEVENTO, e.DESCRICAO " & _
"order by nroperiodo, provdescbase desc, codevento "
rs.Open sqlv, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
'lastper=rs("nroperiodo")
'lasttipo=rs("provdescbase")
rs.movefirst
do while not rs.eof

if ((lasttipo<>rs("provdescbase")) or (lastper<>rs("nroperiodo"))) and lasttipo<>"B" then
	avar=" > " & lasttipo & " " & len(lasttipo) & " " & rs("provdescbase") & " " & len(rs("provdescbase"))
	if inicio=0 then
		response.write "<tr><td class=""campotr"" colspan=3 align=""right"">Total " & lasttipo & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv1,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv2,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv3,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv4,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv5,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv6,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv7,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv8,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv9,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv10,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv11,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv12,2,0,0,0) & "</td>"
		totalv=totalv1+totalv2+totalv3+totalv4+totalv5+totalv6+totalv7+totalv8+totalv9+totalv10+totalv11+totalv12
		response.write "<td class=""campotr"" align=""right"">" & formatnumber(totalv,2,0,0,0) & "</td></tr>"
	end if
	totalv1=0:totalv2=0:totalv3=0:totalv4=0:totalv5=0:totalv6=0:totalv7=0:totalv8=0:totalv9=0:totalv10=0:totalv11=0:totalv12=0:totalv=0
end if

if lastper<>rs("nroperiodo") or lastchapa<>rs2("chapa") then
	if inicio=0 then
		totall1=totalp1-totald1
		totall2=totalp2-totald2
		totall3=totalp3-totald3
		totall4=totalp4-totald4
		totall5=totalp5-totald5
		totall6=totalp6-totald6
		totall7=totalp7-totald7
		totall8=totalp8-totald8
		totall9=totalp9-totald9
		totall10=totalp10-totald10
		totall11=totalp11-totald11
		totall12=totalp12-totald12
		totall=totall1+totall2+totall3+totall4+totall5+totall6+totall7+totall8+totall9+totall10+totall11+totall12
		response.write "<tr><td class=""campolr"" align=""right"" colspan=3><b>" & "Valor Liquido Periodo" & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall1,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall2,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall3,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall4,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall5,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall6,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall7,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall8,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall9,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall10,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall11,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall12,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" align=""right"">" & formatnumber(totall,2,0,0,0) & "</td>"
		response.write "</tr>"
		
		salariobase=0:baseinss=0:basefgts=0:fgtsmes=0:baseirrf=0
'---------------------
		for a=1 to 12
		salariobase=0:baseinss=0:basefgts=0:fgtsmes=0:baseirrf=0
		agencia(a)="-":conta(a)="-"
		mes=a
		sqlq="select ff.SALARIODECALCULO, BASEIRRF, BASEIRRF13, ff.BASEINSS, FF.BASEINSS13, ff.BASEFGTS, FF.BASEFGTS13, INSSCAIXA " & _
		"FROM (select * from corporerm.dbo.PFPERFF where chapa='" & rs2("chapa") & "' AND ANOCOMP=" & ano & " union all " & _
		"select * from corporerm.dbo.PFPERFFCOMPL where chapa='" & rs2("chapa") & "' AND ANOCOMP=" & ano & ") ff " & _
		"WHERE MESCOMP=" & mes & " AND NROPERIODO=" & lastper & " "
		'response.write "<Br>" & sqlq & "<br>"
		rsq.Open sqlq, ,adOpenStatic, adLockReadOnly
		if rsq.recordcount=0 then
			sbase(a)=0:scinss(a)=0:bcfgts(a)=0:vfgts(a)=0:bcirrf(a)=0
		else
			if isnull(rsq("salariodecalculo")) then salariobase=0 else salariobase=rsq("salariodecalculo")

			sqlbase="select max(c.limitesuperior) as baseinss from corporerm.dbo.pcalcvlr c, corporerm.dbo.ptabcalc t " & _
			"where t.iniciovigencia=c.iniciovigencia and t.codigo=c.codtabcalc " & _
			"and c.codtabcalc='01' and '" & dtaccess(dateserial(ano,mes,1)) & "' between t.iniciovigencia and t.finalvigencia "
			rse.Open sqlbase, ,adOpenStatic, adLockReadOnly
			if isnull(rsq("baseinss")) then baseinss=0 else baseinss=rsq("baseinss")
			if isnull(rsq("baseinss13")) then baseinss13=0 else baseinss13=rsq("baseinss13")
			if isnull(rsq("basefgts")) then basefgts=0 else basefgts=rsq("basefgts")
			if isnull(rsq("basefgts13")) then basefgts13=0 else basefgts13=rsq("basefgts13")
			if isnull(rsq("baseirrf")) then baseirrf=0 else baseirrf=rsq("baseirrf")
			if isnull(rsq("baseirrf13")) then baseirrf13=0 else baseirrf13=rsq("baseirrf13")
			if isnull(rsq("insscaixa")) then insscaixa=0 else insscaixa=rsq("insscaixa")
			baseinsst=cdbl(rse("baseinss"))
			baseinssh=cdbl(baseinss)+cdbl(baseinss13)
			if baseinssh>baseinsst then basei=baseinsst else basei=baseinssh
			basei=formatnumber(basei,2)
			basefgts=cdbl(basefgts)+cdbl(basefgts13)
			fgtsmes=int(basefgts*8)/100
			if especial=1 then basefgts=basefgts/divisor
			basefgts=formatnumber(basefgts,2)
			if especial=1 then fgtsmes=fgtsmes/divisor
			fgtsmes=formatnumber(fgtsmes,2)
			baseirrf=cdbl(baseirrf)+cdbl(baseirrf13)
			if especial=1 then baseirrf=baseirrf/divisor
			baseirrf=baseirrf-cdbl(insscaixa)
			rse.close
			sqldep="select valor from corporerm.dbo.pvalfix " & _
			"where '" & dtaccess(dateserial(ano,mes,1)) & "' between iniciovigencia and finalvigencia and codigo='04'"
			rse.Open sqldep, ,adOpenStatic, adLockReadOnly
			valordep=cdbl(rse("valor"))
			rse.close
			sqlqt="select nrodependirrf as ndep " & _
			"from corporerm.dbo.pfhstndp d, (select max(dtmudanca) as mdata from corporerm.dbo.pfhstndp where chapa='" & rs("chapa") & "' and dtmudanca<='" & dtaccess(dateserial(ano,mes,1)) & "') t " & _
			"where chapa='" & rs("chapa") & "' and dtmudanca=t.mdata"
			rse.Open sqlqt, ,adOpenStatic, adLockReadOnly
			if rse.recordcount=0 then ndep=0 else ndep=cdbl(rse("ndep"))
			rse.close
			deducao=valordep * ndep
			baseirrf=baseirrf-deducao
			baseirrf=formatnumber(baseirrf,2)

			sbase(a)=salariobase
			scinss(a)=basei
			bcfgts(a)=basefgts
			vfgts(a)=fgtsmes
			bcirrf(a)=baseirrf
		end if
		rsq.close
		sqlc="select top 1 CODAGENCIAPGTO, CONTAPGTO from corporerm.dbo.PFHSTCPGTO where CHAPA='" & rs2("chapa") & "' and DTMUDANCA<='" & dtaccess(dateserial(ano,mes,1)) & "' order by DTMUDANCA desc "
		'response.write "<br>" & a & "-> " & sqlc
		rsq.Open sqlc, ,adOpenStatic, adLockReadOnly
		if rsq.recordcount=0 then
			sqlc2="select f.CODAGENCIAPAGTO, f.CONTAPAGAMENTO from corporerm.dbo.PFUNC f where CHAPA='" & rs2("chapa") & "'"
			rse.Open sqlc2, ,adOpenStatic, adLockReadOnly
			agencia(a)=rse("codagenciapagto")
			conta(a)=rse("contapagamento")
			rse.close
		else
			agencia(a)=rsq("CODAGENCIAPGTO")
			conta(a)=rsq("CONTAPGTO")
		end if
		rsq.close
	next 'for a=1 to 12

		response.write "<tr><td class=""campoar"" align=""right"" colspan=3>" & "Salário Base" & "</td>"
		for a=1 to 12:response.write "<td class=""campoar"" colspan=2 align=""right"">" & formatnumber(sbase(a),2,0,0,0) & "</td>":next
		response.write "<td class=""campoar"" align=""right"">" & "</td>"

		response.write "<tr><td class=""campolr"" align=""right"" colspan=3>" & "Sal. Contr. INSS" & "</td>"
		for a=1 to 12:response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(scinss(a),2,0,0,0) & "</td>":next
		response.write "<td class=""campolr"" align=""right"">" & "</td>"

		response.write "<tr><td class=""campoar"" align=""right"" colspan=3>" & "Base Cálc. FGTS" & "</td>"
		for a=1 to 12:response.write "<td class=""campoar"" colspan=2 align=""right"">" & formatnumber(bcfgts(a),2,0,0,0) & "</td>":next
		response.write "<td class=""campoar"" align=""right"">" & "</td>"

		response.write "<tr><td class=""campolr"" align=""right"" colspan=3>" & "F.G.T.S. do mês" & "</td>"
		for a=1 to 12:response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(vfgts(a),2,0,0,0) & "</td>":next
		response.write "<td class=""campolr"" align=""right"">" & "</td>"

		response.write "<tr><td class=""campoar"" align=""right"" colspan=3>" & "Base Cálc. IRRF" & "</td>"
		for a=1 to 12:response.write "<td class=""campoar"" colspan=2 align=""right"">" & formatnumber(bcirrf(a),2,0,0,0) & "</td>":next
		response.write "<td class=""campoar"" align=""right"">" & "</td>"
		
		response.write "<tr><td class=""campolr"" align=""right"" colspan=3>" & "Agência" & "</td>"
		for a=1 to 12:response.write "<td class=""campolr"" colspan=2 align=""right"">" & agencia(a) & "</td>":next
		response.write "<td class=""campolr"" align=""right"">" & "</td>"
		
		response.write "<tr><td class=""campoar"" align=""right"" colspan=3>" & "Conta Corrente" & "</td>"
		for a=1 to 12:response.write "<td class=""campoar"" colspan=2 align=""right"">" & conta(a) & "</td>":next
		response.write "<td class=""campoar"" align=""right"">" & "</td>"
		
		totall1=0:totall2=0:totall3=0:totall4=0:totall5=0:totall6=0:totall7=0:totall8=0:totall9=0:totall10=0:totall11=0:totall12=0:totall=0
'-----------------
	end if
end if

if lastper<>rs("nroperiodo") or lastchapa<>rs2("chapa") then
	response.write "<tr><td class=grupor colspan=28>Período: " & rs("nroperiodo") & "</td></tr>"
	lasttipo=""
	totalv1=0:totalv2=0:totalv3=0:totalv4=0:totalv5=0:totalv6=0:totalv7=0:totalv8=0:totalv9=0:totalv10=0:totalv11=0:totalv12=0:totalv=0
	totall1=0:totall2=0:totall3=0:totall4=0:totall5=0:totall6=0:totall7=0:totall8=0:totall9=0:totall10=0:totall11=0:totall12=0:totall=0
	totalp1=0:totalp2=0:totalp3=0:totalp4=0:totalp5=0:totalp6=0:totalp7=0:totalp8=0:totalp9=0:totalp10=0:totalp11=0:totalp12=0:totalp=0
	totald1=0:totald2=0:totald3=0:totald4=0:totald5=0:totald6=0:totald7=0:totald8=0:totald9=0:totald10=0:totald11=0:totald12=0:totald=0
end if

if lasttipo<>rs("provdescbase") then
	select case rs("provdescbase")
		case "P"
			tipoeve="P - Proventos"
		case "D"
			tipoeve="D - Descontos"
		case "B"
			tipoeve="B - Eventos de Base"
	end select
	response.write "<tr><td class=""campoar"" colspan=28>" & tipoeve & "</td></tr>"
end if

totalano=cdbl(rs("v01"))+cdbl(rs("v02"))+cdbl(rs("v03"))+cdbl(rs("v04"))+cdbl(rs("v05"))+cdbl(rs("v06"))+cdbl(rs("v07"))+cdbl(rs("v08"))+cdbl(rs("v09"))+cdbl(rs("v10"))+cdbl(rs("v11"))+cdbl(rs("v12"))
totalv1=totalv1+cdbl(rs("v01")) : totalv2=totalv2+cdbl(rs("v02"))
totalv3=totalv3+cdbl(rs("v03")) : totalv4=totalv4+cdbl(rs("v04"))
totalv5=totalv5+cdbl(rs("v05")) : totalv6=totalv6+cdbl(rs("v06"))
totalv7=totalv7+cdbl(rs("v07")) : totalv8=totalv8+cdbl(rs("v08"))
totalv9=totalv9+cdbl(rs("v09")) : totalv10=totalv10+cdbl(rs("v10"))
totalv11=totalv11+cdbl(rs("v11")) : totalv12=totalv12+cdbl(rs("v12"))

if rs("provdescbase")="P" then
	totalp1=totalp1+cdbl(rs("v01")) : totalp2=totalp2+cdbl(rs("v02"))
	totalp3=totalp3+cdbl(rs("v03")) : totalp4=totalp4+cdbl(rs("v04"))
	totalp5=totalp5+cdbl(rs("v05")) : totalp6=totalp6+cdbl(rs("v06"))
	totalp7=totalp7+cdbl(rs("v07")) : totalp8=totalp8+cdbl(rs("v08"))
	totalp9=totalp9+cdbl(rs("v09")) : totalp10=totalp10+cdbl(rs("v10"))
	totalp11=totalp11+cdbl(rs("v11")) : totalp12=totalp12+cdbl(rs("v12"))
end if
if rs("provdescbase")="D" then
	totald1=totald1+cdbl(rs("v01")) : totald2=totald2+cdbl(rs("v02"))
	totald3=totald3+cdbl(rs("v03")) : totald4=totald4+cdbl(rs("v04"))
	totald5=totald5+cdbl(rs("v05")) : totald6=totald6+cdbl(rs("v06"))
	totald7=totald7+cdbl(rs("v07")) : totald8=totald8+cdbl(rs("v08"))
	totald9=totald9+cdbl(rs("v09")) : totald10=totald10+cdbl(rs("v10"))
	totald11=totald11+cdbl(rs("v11")) : totald12=totald12+cdbl(rs("v12"))
end if

if rs("r01")<>"" and cdbl(rs("r01"))>0 then r01=rs("r01") else r01=""
if rs("r02")<>"" and cdbl(rs("r02"))>0 then r02=rs("r02") else r02=""
if rs("r03")<>"" and cdbl(rs("r03"))>0 then r03=rs("r03") else r03=""
if rs("r04")<>"" and cdbl(rs("r04"))>0 then r04=rs("r04") else r04=""
if rs("r05")<>"" and cdbl(rs("r05"))>0 then r05=rs("r05") else r05=""
if rs("r06")<>"" and cdbl(rs("r06"))>0 then r06=rs("r06") else r06=""
if rs("r07")<>"" and cdbl(rs("r07"))>0 then r07=rs("r07") else r07=""
if rs("r08")<>"" and cdbl(rs("r08"))>0 then r08=rs("r08") else r08=""
if rs("r09")<>"" and cdbl(rs("r09"))>0 then r09=rs("r09") else r09=""
if rs("r10")<>"" and cdbl(rs("r10"))>0 then r10=rs("r10") else r10=""
if rs("r11")<>"" and cdbl(rs("r11"))>0 then r11=rs("r11") else r11=""
if rs("r12")<>"" and cdbl(rs("r12"))>0 then r12=rs("r12") else r12=""

if cdbl(rs("v01"))<>0 then v01=formatnumber(rs("v01"),2,0,0,0) else v01=""
if cdbl(rs("v02"))<>0 then v02=formatnumber(rs("v02"),2,0,0,0) else v02=""
if cdbl(rs("v03"))<>0 then v03=formatnumber(rs("v03"),2,0,0,0) else v03=""
if cdbl(rs("v04"))<>0 then v04=formatnumber(rs("v04"),2,0,0,0) else v04=""
if cdbl(rs("v05"))<>0 then v05=formatnumber(rs("v05"),2,0,0,0) else v05=""
if cdbl(rs("v06"))<>0 then v06=formatnumber(rs("v06"),2,0,0,0) else v06=""
if cdbl(rs("v07"))<>0 then v07=formatnumber(rs("v07"),2,0,0,0) else v07=""
if cdbl(rs("v08"))<>0 then v08=formatnumber(rs("v08"),2,0,0,0) else v08=""
if cdbl(rs("v09"))<>0 then v09=formatnumber(rs("v09"),2,0,0,0) else v09=""
if cdbl(rs("v10"))<>0 then v10=formatnumber(rs("v10"),2,0,0,0) else v10=""
if cdbl(rs("v11"))<>0 then v11=formatnumber(rs("v11"),2,0,0,0) else v11=""
if cdbl(rs("v12"))<>0 then v12=formatnumber(rs("v12"),2,0,0,0) else v12=""
%>
<tr>
	<td class="campor"><%=rs("codevento")%></td>
	<td class="campor" nowrap><%=left(lcase(rs("descricao")),25)%></td>
	<td class="campor"><%=rs("provdescbase")%></td>
	<td class="campor" align="right"><i><%=r01%></i></td>
	<td class="campor" align="right">   <%=v01%></td>
	<td class="campor" align="right"><i><%=r02%></i></td>
	<td class="campor" align="right">   <%=v02%></td>
	<td class="campor" align="right"><i><%=r03%></i></td>
	<td class="campor" align="right">   <%=v03%></td>
	<td class="campor" align="right"><i><%=r04%></i></td>
	<td class="campor" align="right">   <%=v04%></td>
	<td class="campor" align="right"><i><%=r05%></i></td>
	<td class="campor" align="right">   <%=v05%></td>
	<td class="campor" align="right"><i><%=r06%></i></td>
	<td class="campor" align="right">   <%=v06%></td>
	<td class="campor" align="right"><i><%=r07%></i></td>
	<td class="campor" align="right">   <%=v07%></td>
	<td class="campor" align="right"><i><%=r08%></i></td>
	<td class="campor" align="right">   <%=v08%></td>
	<td class="campor" align="right"><i><%=r09%></i></td>
	<td class="campor" align="right">   <%=v09%></td>
	<td class="campor" align="right"><i><%=r10%></i></td>
	<td class="campor" align="right">   <%=v10%></td>
	<td class="campor" align="right"><i><%=r11%></i></td>
	<td class="campor" align="right">   <%=v11%></td>
	<td class="campor" align="right"><i><%=r12%></i></td>
	<td class="campor" align="right">   <%=v12%></td>
	<td class="campor" align="right"><%=formatnumber(totalano,2,0,0,0)%></td>
</tr>
<%
lastper=rs("nroperiodo"):lasttipo=rs("provdescbase"):lastchapa=rs("chapa"):totalano=0
inicio=0
rs.movenext
loop
	if lasttipo<>"B" then
		response.write "<tr><td class=""campotr"" colspan=3 align=""right"">Total " & lasttipo & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv1,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv2,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv3,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv4,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv5,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv6,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv7,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv8,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv9,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv10,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv11,2,0,0,0) & "</td>"
		response.write "<td class=""campotr"" colspan=2 align=""right"">" & formatnumber(totalv12,2,0,0,0) & "</td>"
		totalv=totalv1+totalv2+totalv3+totalv4+totalv5+totalv6+totalv7+totalv8+totalv9+totalv10+totalv11+totalv12
		response.write "<td class=""campotr"" align=""right"">" & formatnumber(totalv,2,0,0,0) & "</td></tr>"
		totalv1=0:totalv2=0:totalv3=0:totalv4=0:totalv5=0:totalv6=0:totalv7=0:totalv8=0:totalv9=0:totalv10=0:totalv11=0:totalv12=0:totalv=0
	end if

		totall1=totalp1-totald1
		totall2=totalp2-totald2
		totall3=totalp3-totald3
		totall4=totalp4-totald4
		totall5=totalp5-totald5
		totall6=totalp6-totald6
		totall7=totalp7-totald7
		totall8=totalp8-totald8
		totall9=totalp9-totald9
		totall10=totalp10-totald10
		totall11=totalp11-totald11
		totall12=totalp12-totald12
		totall=totall1+totall2+totall3+totall4+totall5+totall6+totall7+totall8+totall9+totall10+totall11+totall12
		response.write "<tr><td class=""campolr"" align=""right"" colspan=3><b>" & "Valor Liquido Periodo" & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall1,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall2,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall3,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall4,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall5,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall6,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall7,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall8,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall9,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall10,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall11,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" colspan=2 align=""right""><b>" & formatnumber(totall12,2,0,0,0) & "</td>"
		response.write "<td class=""campolr"" align=""right"">" & formatnumber(totall,2,0,0,0) & "</td>"
		response.write "</tr>"
	totall1=0:totall2=0:totall3=0:totall4=0:totall5=0:totall6=0:totall7=0:totall8=0:totall9=0:totall10=0:totall11=0:totall12=0:totall=0
	
'-----------------------
'---------------------
		for a=1 to 12
		salariobase=0:baseinss=0:basefgts=0:fgtsmes=0:baseirrf=0
		agencia(a)="-":conta(a)="-"
		mes=a
		sqlq="select ff.SALARIODECALCULO, BASEIRRF, BASEIRRF13, ff.BASEINSS, FF.BASEINSS13, ff.BASEFGTS, FF.BASEFGTS13, INSSCAIXA " & _
		"FROM (select * from corporerm.dbo.PFPERFF where chapa='" & rs2("chapa") & "' AND ANOCOMP=" & ano & " union all " & _
		"select * from corporerm.dbo.PFPERFFCOMPL where chapa='" & rs2("chapa") & "' AND ANOCOMP=" & ano & ") ff " & _
		"WHERE MESCOMP=" & mes & " AND NROPERIODO=" & lastper & " "
		'response.write sqlq
		rsq.Open sqlq, ,adOpenStatic, adLockReadOnly
		if rsq.recordcount=0 then
			sbase(a)=0:scinss(a)=0:bcfgts(a)=0:vfgts(a)=0:bcirrf(a)=0
		else
			if isnull(rsq("salariodecalculo")) then salariobase=0 else salariobase=rsq("salariodecalculo")

			sqlbase="select max(c.limitesuperior) as baseinss from corporerm.dbo.pcalcvlr c, corporerm.dbo.ptabcalc t " & _
			"where t.iniciovigencia=c.iniciovigencia and t.codigo=c.codtabcalc " & _
			"and c.codtabcalc='01' and '" & dtaccess(dateserial(ano,mes,1)) & "' between t.iniciovigencia and t.finalvigencia "
			rse.Open sqlbase, ,adOpenStatic, adLockReadOnly
			if isnull(rsq("baseinss")) then baseinss=0 else baseinss=rsq("baseinss")
			if isnull(rsq("baseinss13")) then baseinss13=0 else baseinss13=rsq("baseinss13")
			if isnull(rsq("basefgts")) then basefgts=0 else basefgts=rsq("basefgts")
			if isnull(rsq("basefgts13")) then basefgts13=0 else basefgts13=rsq("basefgts13")
			if isnull(rsq("baseirrf")) then baseirrf=0 else baseirrf=rsq("baseirrf")
			if isnull(rsq("baseirrf13")) then baseirrf13=0 else baseirrf13=rsq("baseirrf13")
			if isnull(rsq("insscaixa")) then insscaixa=0 else insscaixa=rsq("insscaixa")
			baseinsst=cdbl(rse("baseinss"))
			baseinssh=cdbl(baseinss)+cdbl(baseinss13)
			if baseinssh>baseinsst then basei=baseinsst else basei=baseinssh
			basei=formatnumber(basei,2)
			basefgts=cdbl(basefgts)+cdbl(basefgts13)
			fgtsmes=int(basefgts*8)/100
			if especial=1 then basefgts=basefgts/divisor
			basefgts=formatnumber(basefgts,2)
			if especial=1 then fgtsmes=fgtsmes/divisor
			fgtsmes=formatnumber(fgtsmes,2)
			baseirrf=cdbl(baseirrf)+cdbl(baseirrf13)
			if especial=1 then baseirrf=baseirrf/divisor
			baseirrf=baseirrf-cdbl(insscaixa)
			rse.close
			sqldep="select valor from corporerm.dbo.pvalfix " & _
			"where '" & dtaccess(dateserial(ano,mes,1)) & "' between iniciovigencia and finalvigencia and codigo='04'"
			rse.Open sqldep, ,adOpenStatic, adLockReadOnly
			valordep=cdbl(rse("valor"))
			rse.close
			sqlqt="select nrodependirrf as ndep " & _
			"from corporerm.dbo.pfhstndp d, (select max(dtmudanca) as mdata from corporerm.dbo.pfhstndp where chapa='" & rs2("chapa") & "' and dtmudanca<='" & dtaccess(dateserial(ano,mes,1)) & "') t " & _
			"where chapa='" & rs2("chapa") & "' and dtmudanca=t.mdata"
			rse.Open sqlqt, ,adOpenStatic, adLockReadOnly
			if rse.recordcount=0 then ndep=0 else ndep=cdbl(rse("ndep"))
			rse.close
			deducao=valordep * ndep
			baseirrf=baseirrf-deducao
			baseirrf=formatnumber(baseirrf,2)

			sbase(a)=salariobase
			scinss(a)=basei
			bcfgts(a)=basefgts
			vfgts(a)=fgtsmes
			bcirrf(a)=baseirrf
		end if
		rsq.close
		sqlc="select top 1 CODAGENCIAPGTO, CONTAPGTO from corporerm.dbo.PFHSTCPGTO where CHAPA='" & rs2("chapa") & "' and DTMUDANCA<='" & dtaccess(dateserial(ano,mes,1)) & "' order by DTMUDANCA desc "
		'response.write "<br>" & a & "-> " & sqlc
		rsq.Open sqlc, ,adOpenStatic, adLockReadOnly
		if rsq.recordcount=0 then
			sqlc2="select f.CODAGENCIAPAGTO, f.CONTAPAGAMENTO from corporerm.dbo.PFUNC f where CHAPA='" & rs2("chapa") & "'"
			rse.Open sqlc2, ,adOpenStatic, adLockReadOnly
			agencia(a)=rse("codagenciapagto")
			conta(a)=rse("contapagamento")
			rse.close
		else
			agencia(a)=rsq("CODAGENCIAPGTO")
			conta(a)=rsq("CONTAPGTO")
		end if
		rsq.close
	next 'for a=1 to 12

		response.write "<tr><td class=""campoar"" align=""right"" colspan=3>" & "Salário Base" & "</td>"
		for a=1 to 12:response.write "<td class=""campoar"" colspan=2 align=""right"">" & formatnumber(sbase(a),2,0,0,0) & "</td>":next
		response.write "<td class=""campoar"" align=""right"">" & "</td>"

		response.write "<tr><td class=""campolr"" align=""right"" colspan=3>" & "Sal. Contr. INSS" & "</td>"
		for a=1 to 12:response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(scinss(a),2,0,0,0) & "</td>":next
		response.write "<td class=""campolr"" align=""right"">" & "</td>"

		response.write "<tr><td class=""campoar"" align=""right"" colspan=3>" & "Base Cálc. FGTS" & "</td>"
		for a=1 to 12:response.write "<td class=""campoar"" colspan=2 align=""right"">" & formatnumber(bcfgts(a),2,0,0,0) & "</td>":next
		response.write "<td class=""campoar"" align=""right"">" & "</td>"

		response.write "<tr><td class=""campolr"" align=""right"" colspan=3>" & "F.G.T.S. do mês" & "</td>"
		for a=1 to 12:response.write "<td class=""campolr"" colspan=2 align=""right"">" & formatnumber(vfgts(a),2,0,0,0) & "</td>":next
		response.write "<td class=""campolr"" align=""right"">" & "</td>"

		response.write "<tr><td class=""campoar"" align=""right"" colspan=3>" & "Base Cálc. IRRF" & "</td>"
		for a=1 to 12:response.write "<td class=""campoar"" colspan=2 align=""right"">" & formatnumber(bcirrf(a),2,0,0,0) & "</td>":next
		response.write "<td class=""campoar"" align=""right"">" & "</td>"
		
		response.write "<tr><td class=""campolr"" align=""right"" colspan=3>" & "Agência" & "</td>"
		for a=1 to 12:response.write "<td class=""campolr"" colspan=2 align=""right"">" & agencia(a) & "</td>":next
		response.write "<td class=""campolr"" align=""right"">" & "</td>"
		
		response.write "<tr><td class=""campoar"" align=""right"" colspan=3>" & "Conta Corrente" & "</td>"
		for a=1 to 12:response.write "<td class=""campoar"" colspan=2 align=""right"">" & conta(a) & "</td>":next
		response.write "<td class=""campoar"" align=""right"">" & "</td>"
		
		totall1=0:totall2=0:totall3=0:totall4=0:totall5=0:totall6=0:totall7=0:totall8=0:totall9=0:totall10=0:totall11=0:totall12=0:totall=0
'-----------------
'-----------------------	
	
else
	response.write "<tr><td class=campo colspan=28>Sem histórico de ficha financeira</td></tr>"
end if
%>



</table>
<!-- fim holerith -->

<%
rs.close

if rs2.absoluteposition<rs2.recordcount then response.write "<DIV style=""page-break-after:always""></DIV>" '<!-- Aqui quebra a página -->

if request.form("ordem")="chapa" then
	session("ultimohol")=rs2("chapa") 
	session("ultimohtp")="chapa"
else 
	session("ultimohol")=rs2("nome")
	session("ultimohtp")="nome"
end if

lastchapa=rs2("chapa")
inicio=0
rs2.movenext
loop
rs2.close


end if 'request.form=0
%>

<%
set rs=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>