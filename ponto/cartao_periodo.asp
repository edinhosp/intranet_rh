<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a46")="N" or session("a46")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=iso-8859-1">
<title>Espelho de Marcação Eletrônica</title>
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
set rs1=server.createobject ("ADODB.Recordset")
Set rs1.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao

if request.form="" then
%>
<p class=titulo>Espelho de Marcação de Ponto
<form method="POST" action="cartao_periodo.asp">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse">
<tr><td class=titulo>Chapa</td><td class=titulo align="center">De</td><td class=titulo align="center">Até</td></tr>
<tr>
<td class=titulo><input type="text" name="chapa" value="" size="6"></td>
<td class=titulo><input type="text" name="dtinicio" value="<%=dateserial(year(now),month(now)-1,day(now))%>" size="8"></td>
<td class=titulo><input type="text" name="dtfim" value="<%=dateserial(year(now),month(now),day(now)-1)%>" size="8"></td>
</tr>
<tr><td colspan=3 class=titulo>
<input type="submit" value="Clique para Visualizar" class="button" name="B3">
</td></tr>
</table>
<a href="cartao_periodos.asp">Versão simples</a>
</form>
<hr>
<%
else 'request.form <>''
	data1=request.form("dtinicio")
	data2=request.form("dtfim")
	data1=cdate(data1):data2=cdate(data2)
	udia=day(dateserial(ano,mes+1,1)-1)
	chapa=numzero(request.form("chapa"),5)
	sqld="select f.nome, f.codsindicato, c.nome as funcao, s.descricao as setor from corporerm.dbo.pfunc f, corporerm.dbo.psecao s, corporerm.dbo.pfuncao c where f.codsecao=s.codigo and f.codfuncao=c.codigo and f.chapa='" & chapa & "'"
	rs2.Open sqld, ,adOpenStatic, adLockReadOnly
	sindicato=rs2("codsindicato")
	if sindicato="03" then coluna=7 else coluna=5
mano=2004
mmes=6
	
%>
<%linha=linha+1%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="600">
<tr><td class=campo>Espelho de Marcação de Ponto</td></tr></table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="600">
<tr><td class=campo>FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</td></tr></table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="600">
<tr><td class=campo>Período: de <b><%=data1%> até <%=data2%></td></tr></table>
<hr>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class="campor">Chapa</td>
	<td class="campor">Nome</td>
	<td class="campor">Setor</td>
	<td class="campor">Função</td></tr>
<tr><td class="campor"><%=chapa%></td>
	<td class="campor"><b><%=rs2("nome")%></b></td>
	<td class="campor"><%=rs2("Setor")%></td>
	<td class="campor"><%=rs2("funcao")%></td></tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" >
<tr>
	<td class=grupo align="center" colspan=3 style="border-right:2px solid #000000">Datas</td>
	<td class=grupo align="center" colspan=<%=coluna%> style="border-right:2px solid #000000">Carga Horária a cumprir</td>
	<td class=grupo align="center" colspan=7 style="border-right:2px solid #000000">Marcações efetuadas</td>
	<td class=grupo align="center" colspan=8 style="border-right:2px solid #000000">Ocorrências</td>
</tr>
<tr>
	<td class=titulo align="center">Data</td>
	<td class=titulo align="center">Dia</td>
	<td class=titulo align="center" style="border-right:2px solid #000000">Ind</td>
	<td class=titulo align="center" width=35>1</td>
	<td class=titulo align="center" width=35>2</td>
	<td class=titulo align="center" width=35>3</td>
	<td class=titulo align="center" width=35>4</td>
	<%if sindicato="03" then %>
	<td class=titulo align="center" width=35>5</td>
	<td class=titulo align="center" width=35>6</td>
	<%end if%>
	<td class=titulo align="center" width=40 style="border-right:2px solid #000000">=</td>
	<td class=titulo align="center" width=35>1</td>
	<td class=titulo align="center" width=35>2</td>
	<td class=titulo align="center" width=35>3</td>
	<td class=titulo align="center" width=35>4</td>
	<td class=titulo align="center" width=35>5</td>
	<td class=titulo align="center" width=35>6</td>
	<td class=fundo align="center" style="border-right:2px solid #000000">H.Trab.</td>
	<td class=fundo align="center" >&nbsp;Extra&nbsp;</td>
	<td class=fundo align="center" >&nbsp;Atraso&nbsp;</td>
	<td class=fundo align="center" >&nbsp;Falta&nbsp;</td>
	<td class=fundo align="center" >&nbsp;AdN&nbsp;</td>
	<td class=fundo align="center" style="border-right:2px solid #000000">&nbsp;Abono&nbsp;</td>
	<td class=fundo align="center" >&nbsp;Descrição Abono&nbsp;</td>
	<td class=fundo align="center" >&nbsp;Inicio&nbsp;</td>
	<td class=fundo align="center" style="border-right:2px solid #000000">&nbsp;Fim&nbsp;</td>
</tr>
<%
rs2.close
diasloop=datediff("d",data1,data2)+1:'response.write diasloop & "<br>"
diasloop=cint(diasloop)
totalchcumprir=0
totalchcumprida=0

Redim marc(diasloop,6), formato(diasloop,6)

for e=data1 to data2
	idmatriz=e-(data1)
	'marcações do chronus
	sqlcr="select * from ( " & _
	"select chapa, day(data) as dia, data, batida, natureza, status from corporerm.dbo.abatfunam where chapa='" & chapa & "' " & _
	"and data='" & dtaccess(e) & "' " & _
	"UNION ALL " & _
	"select chapa, day(data) as dia, data, batida, natureza, status from corporerm.dbo.abatfun where chapa='" & chapa & "' " & _
	"and data='" & dtaccess(e) & "' " & _
	") z order by batida "
	rs2.Open sqlcr, ,adOpenStatic, adLockReadOnly
	marcacao=1
	if rs2.recordcount>0 then
		rs2.movefirst:do while not rs2.eof
		'response.write "<br>" & rs2("data")
		'------------------------------------------
		dia=rs2("dia"):data=rs2("data")
		natureza=rs2("natureza")
		batida=formatdatetime((rs2("batida")/60)/24,4)
		if dia=diaant then marcacao=marcacao+1 else marcacao=1
		'nat(dia,marcacao)=rs2("natureza")
		if natureza=0 or natureza=4 then natu=0
		if natureza=1 or natureza=5 then natu=1
		resto=marcacao mod 2
		if resto=0 and natu=0 then marcacao=marcacao+1 else marcacao=marcacao
		if resto<>0 and natu=1 then marcacao=marcacao+1 else marcacao=marcacao
		marc(idmatriz,marcacao)=batida:'response.write ">> " & idmatriz & "/" & marcacao & " >> " & marc(idmatriz,marcacao) & "<br>"
		if rs2("status")="D" then formato(idmatriz,marcacao)="<font color='red'>" 'else formato(dia,marcacao)="<font color='black'>"
		diaant=dia
		'------------------------------------------
		rs2.movenext:loop
	else 'recordcount rs2
		for b=1 to 6:marc(idmatriz,b)="":next
	end if 'recordcount rs2
	rs2.close
next

dtponto=data1
sql1="select top 1 codhorario, indiniciohor, dtmudanca from corporerm.dbo.pfhsthor where dtmudanca<='" & dtaccess(dtponto) & "' and chapa='" & chapa & "' order by dtmudanca desc "
'response.write sql1
rs2.Open sql1, ,adOpenStatic, adLockReadOnly
naotem=0
if rs2.recordcount>0 then
indice=rs2("indiniciohor")-0
dtmudanca=rs2("dtmudanca")
codhor=rs2("codhorario")
naotem=1
end if
rs2.close
if naotem=0 then
	sql1="select top 1 codhorario, indiniciohor, dtmudanca from corporerm.dbo.pfhsthor where dtmudanca>='" & dtaccess(dtponto) & "' and chapa='" & chapa & "' order by dtmudanca desc "
	rs2.Open sql1, ,adOpenStatic, adLockReadOnly
	indice=rs2("indiniciohor")-0
	dtmudanca=rs2("dtmudanca")
	codhor=rs2("codhorario")
	rs2.close
end if
sqlb="select max(indice) as loop from corporerm.dbo.abathor where codhorario='" & codhor & "'"
Set rs2=conexao.Execute (sqlb, , adCmdText)
maxindice=rs2("loop"):rs2.close

'response.write "<br>" & indice & "<br>" & dtmudanca &  "<br>" & codhor
dias=datediff("d",dtmudanca,dtponto)
'response.write "<br>" & dias & "<br>" & maxindice
novadata=dtmudanca
for a=1 to dias
	novadata=novadata+1
	'response.write "<br>" & a & "-" & novadata
	indice=indice+1
	if indice>maxindice then indice=1
	'response.write " >> " & a & "-" & indice
next
'response.write "<br>" & indice

'for a=1 to udia
for e=data1 to data2
	'dtponto=dateserial(ano,mes,a)
	dtponto=e
	idmatriz=e-(data1)
	if idmatriz=0 then indice=indice else indice=indice+1
	response.write "<tr>"
	sql1="select top 1 codhorario, indiniciohor, dtmudanca from corporerm.dbo.pfhsthor where dtmudanca='" & dtaccess(dtponto) & "' and chapa='" & chapa & "' "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then 
		'response.write sql1
		indice=rs("indiniciohor")-0
		if codhor<>rs("codhorario") then
			sqlb="select max(indice) as loop from corporerm.dbo.abathor where codhorario='" & rs("codhorario") & "'"
			Set rsd=conexao.Execute (sqlb, , adCmdText)
			maxindice=rsd("loop"):rsd.close
			codhor=rs("codhorario")
		end if
	else 
		indice=indice
	end if
	rs.close
	if indice>maxindice then indice=1

	response.write "<td class=""campo"" align=""center"">" & dtponto & "</td>"
	response.write "<td class=""campo"" align=""center"">" & weekdayname(weekday(dtponto),-1) & "</td>"
	response.write "<td class=""campo"" align=""center"" style='border-right:2px solid #000000'><font color=gray>" & indice & "</td>"
	
	'************* feriado 	
	sql1="select nome from corporerm.dbo.gferiado where diaferiado='" & dtaccess(dtponto) & "' "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then feriado=rs("nome") else feriado=""
	rs.close
	'*************ocorrencias
	sql1="select falta, atraso, abono, base, adicional, htrab, extraexecutado, extraautorizado from corporerm.dbo.aafhtfunam where data='" & dtaccess(dtponto) & "' and chapa='" & chapa & "' "
	sql1=sql1 & "union all "
	sql1=sql1 & "select falta, atraso, abono, base, adicional, htrab, extraexecutado, extraautorizado from corporerm.dbo.aafhtfun where data='" & dtaccess(dtponto) & "' and chapa='" & chapa & "' "

	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then 
		abono=rs("abono")
		falta=rs("falta")
		atraso=rs("atraso")
		base=rs("base")
		htrab=rs("htrab")
		adicional=rs("adicional")
		extraexecutada=rs("extraexecutado")
		extraautorizada=rs("extraautorizado")
		extra=extraautorizada
	else
		abono=0:falta=0:atraso=0
		base=0:htrab=0:adicional=0
		extraexecutada=0:extraautorizada=0:extra=0
	end if
	rs.close
	tbase=tbase+base
	'*************abonos
	sql1="SELECT f.CHAPA, f.DATA, f.CODABONO, a.DESCRICAO, Min(f.HORAINICIO) AS inicio, Max(f.HORAFIM) AS fim " & _
	"FROM corporerm.dbo.AABONFUN f, corporerm.dbo.AABONO a WHERE f.CODABONO=a.CODIGO " & _
	"GROUP BY f.CHAPA, f.DATA, f.CODABONO, a.DESCRICAO " & _
	"HAVING f.data='" & dtaccess(dtponto) & "' and f.chapa='" & chapa & "' "
	rs.Open sql1, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then 
		abdesc=rs("descricao")
		abini=rs("inicio")
		abfim=rs("fim")
	else
		abdesc=""
		abini=0
		abfim=0
	end if
	rs.close

	if feriado<>"" then 
		response.write "<td class=campo align=""left"" colspan=4>&nbsp;<font color=red><b>" & feriado & "</td>"
		response.write "<td class=campo align=""center"" style='border-right:2px solid #000000'>" & "</td>"
	else
	if sindicato="01" then
		sql1="select ent1,sai1,ent2,sai2 from aindice where codhorario='" & codhor & "' and indice=" & indice
		rs.Open sql1, ,adOpenStatic, adLockReadOnly
		if rs.recordcount>0 then
			if isdate(rs("ent1")) then ent1=formatdatetime(rs("ent1"),4) else ent1  ="-"
			if isdate(rs("sai1")) then sai1=formatdatetime(rs("sai1"),4) else sai1="-"
			if isdate(rs("ent2")) then ent2=formatdatetime(rs("ent2"),4) else ent2  ="-"
			if isdate(rs("sai2")) then sai2=formatdatetime(rs("sai2"),4) else sai2="-"
		end if 'rs.recordcount>0
		rs.close	
		response.write "<td class=campo align=""center"">" & ent1 & "</td>"
		response.write "<td class=campo align=""center"">" & sai1 & "</td>"
		response.write "<td class=campo align=""center"">" & ent2 & "</td>"
		response.write "<td class=campo align=""center"">" & sai2 & "</td>"
		response.write "<td class=campo align=""center"" style='border-right:2px solid #000000'>"
		if base>0 then response.write formatdatetime((base/60)/24,4)
		response.write "</td>"
	end if
	
	if sindicato="03" then
	sqlch="select chapa, dia_mes, diasem, ent1, saida1, ent2, saida2, ent3, saida3, totalch " & _
	"from ttapontprof_2 " & _
	"where chapa='" & chapa & "' and dia_mes='" & dtaccess(dtponto) & "' "
	'response.write sqlch
	'grade horaria
	rs.Open sqlch, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
		if isdate(rs("ent1"))   then ent1  =formatdatetime(rs("ent1"),4)   else ent1  ="-"
		if isdate(rs("saida1")) then saida1=formatdatetime(rs("saida1"),4) else saida1="-"
		if isdate(rs("ent2"))   then ent2  =formatdatetime(rs("ent2"),4)   else ent2  ="-"
		if isdate(rs("saida2")) then saida2=formatdatetime(rs("saida2"),4) else saida2="-"
		if isdate(rs("ent3"))   then ent3  =formatdatetime(rs("ent3"),4)   else ent3  ="-"
		if isdate(rs("saida3")) then saida3=formatdatetime(rs("saida3"),4) else saida3="-"
		response.write "<td class=campo align=""center"">" & ent1   & "</td>"
		response.write "<td class=campo align=""center"">" & saida1 & "</td>"
		response.write "<td class=campo align=""center"">" & ent2   & "</td>"
		response.write "<td class=campo align=""center"">" & saida2 & "</td>"
		response.write "<td class=campo align=""center"">" & ent3   & "</td>"
		response.write "<td class=campo align=""center"">" & saida3 & "</td>"
		response.write "<td class=campo align=""center"">" & formatdatetime(rs("totalch"),4) & "</td>"
		totb=cdate(rs("totalch"))
		totalchcumprir=totalchcumprir+rs("totalch")
		tchcumprir=tchcumprir+rs("totalch")
	else 'recordcount rs
		response.write "<td class=titulor align=""center"" colspan=6>&nbsp;</td>"
		response.write "<td class=campo align=""center"">-</td>"
		totb=0
	end if 'recordcount rs
	if weekday(dtponto)=1 then
		tsem1=tchcumprir
		tchcumprir=0
	end if
	rs.close
	end if 'sindicato -03
	end if 'feriado <>''
	
	for b=1 to 6
		batida=marc(idmatriz,b)
		if batida<>"" then ultima=b
		response.write "<td class=campo align=""center"">" & formato(idmatriz,b) & batida & "</font></td>"
	next
	If marc(idmatriz,1)="" and marc(idmatriz,2)="" and marc(idmatriz,3)="" and marc(idmatriz,4)="" and marc(idmatriz,5)="" and marc(idmatriz,6)="" then
		tot1=0:tot2=0:tot3=0
	else
		if marc(idmatriz,2)="" and marc(idmatriz,1)<>"" then tot1=0 else tot1=cdate(marc(idmatriz,2))-cdate(marc(idmatriz,1))
		if marc(idmatriz,4)="" and marc(idmatriz,3)<>"" then tot2=cdate(marc(idmatriz,3))-cdate(marc(idmatriz,2)) else tot2=cdate(marc(idmatriz,4))-cdate(marc(idmatriz,3))
		if marc(idmatriz,6)="" and marc(idmatriz,5)<>"" then tot3=cdate(marc(idmatriz,5))-cdate(marc(idmatriz,4)) else tot3=cdate(marc(idmatriz,6))-cdate(marc(idmatriz,5))
	end if

	thtrab=thtrab+htrab
	totc=tot1+tot2+tot3
	totch=formatdatetime(totc,4)
	if totc=0 then totch="-" else totch=formatdatetime(totc,4)
		
	response.write "<td class=campo align=""center"" style='border-right:2px solid #000000'>"
	if htrab>0 then response.write formatdatetime((htrab/60)/24,4) 
	if htrab=0 then response.write "<font color=gray>" & totch 
	response.write "</font></td>"

	if htrab>0 then tcumprida1=tcumprida1 + htrab:tcumprida2=tcumprida2 + htrab
	if htrab=0 then tcumprida2=tcumprida2 + (totc*24*60)
		
	textra=textra+extra
	response.write "<td class=campo align=""center"">"
	if extra>0 then response.write formatdatetime((extra/60)/24,4) 
	response.write "</td>"
	tatraso=tatraso+atraso
	response.write "<td class=campo align=""center"">"
	if atraso>0 then response.write formatdatetime((atraso/60)/24,4) 
	response.write "</td>"
	tfalta=tfalta+falta
	response.write "<td class=campo align=""center"">"
	if falta>0 then response.write formatdatetime((falta/60)/24,4) 
	response.write "</td>"
	tadicional=tadicional+adicional
	response.write "<td class=campo align=""center"">"
	if adicional>0 then response.write formatdatetime((adicional/60)/24,4) 
	response.write "</td>"
	tabono=tabono+abono
	response.write "<td class=campo align=""center"" style='border-right:2px solid #000000'>"
	if abono>0 then response.write formatdatetime((abono/60)/24,4) 
	response.write "</td>"

	response.write "<td class=campo align=""left"">" & abdesc & "</td>"
	response.write "<td class=campo align=""center"">"
		if abini>0 then response.write formatdatetime((abini/60)/24,4)
	response.write "</td>"
	response.write "<td class=campo align=""center"" style='border-right:2px solid #000000'>"
		if abfim>0 then response.write formatdatetime((abfim/60)/24,4)
	response.write "</td>"
		
	response.write "</tr>"
next
%>
  <tr>
    <td class=titulo align="left" colspan=<%=coluna+2%>>&nbsp;Totais</td>
    <td class=campo align="center" style="border-right:2px solid #000000"><%=formatnumber(tbase/60,2)%></td>
    <td class=titulo align="left" colspan=6>&nbsp;Totais</td>
    <td class=campo align="center" style="border-right:2px solid #000000"><%=formatnumber(tcumprida1/60,2)%></td>
    <td class=campo align="center"><%=formatnumber(textra/60,2)%></td>
    <td class=campo align="center"><%=formatnumber(tatraso/60,2)%></td>
    <td class=campo align="center"><%=formatnumber(tfalta/60,2)%></td>
    <td class=campo align="center"><%=formatnumber(tadicional/60,2)%></td>
    <td class=campo align="center" style="border-right:2px solid #000000"><%=formatnumber(tabono/60,2)%></td>
  </tr>

</table>
<p style="margin-top: 0; margin-bottom: 0"><font size=1><b>Nomeações para Atividades</b></font></p>
<table border="1" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="650">
<tr>
	<td class=titulor>Nomeação</td>
	<td class=titulor>Portaria</td>
	<td class=titulor>Curso</td>
	<td class=titulor>Evento</td>
	<td class=titulor>CHS</td>
	<td class=titulor>CHM</td>
	<td class=titulor>Inicio</td>
	<td class=titulor>Término</td>
</tr>
<%
sqln="SELECT n.NOMEACAO, i.PORTARIA, i.CARGO, i.CODEVE, i.CH, i.MAND_INI, i.MAND_FIM " & _
"FROM n_nomeacoes n, n_indicacoes i WHERE n.id_nomeacao = i.id_nomeacao " & _
"AND i.CHAPA='" & chapa & "' AND '" & dtaccess(dateserial(ano,mes,1)) & "' Between mand_ini And mand_fim"
rs.Open sqln, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class="campor"><%=rs("nomeacao")%></td>
	<td class="campor"><%=rs("portaria")%></td>
	<td class="campor"><%=rs("cargo")%></td>
	<td class="campor"><%=rs("codeve")%></td>
	<td class="campor" align="center"><%=rs("ch")%></td>
	<td class="campor" align="center"><%if rs("ch")<>"" then response.write rs("ch")*4 else response.write "&nbsp;"%></td>
	<td class="campor"><%=rs("mand_ini")%></td>
	<td class="campor"><%=rs("mand_fim")%></td>
</tr>
<%
if isnumeric(rs("ch")) then taes=taes+cdbl(rs("ch"))
if isnumeric(rs("ch")) then taem=taem+(cdbl(rs("ch"))*4)
rs.movenext
loop
%>
<tr>
	<td class="campor" colspan=4>&nbsp;</td>
	<td class="campor" align="center"><%=taes%></td>
	<td class="campor" align="center"><%=taem%></td>
	<td class="campor" colspan=2>&nbsp;</td>
</tr>
<%
else
	response.write "<tr><td class=""campor"" colspan=8>Sem nomeações</td></tr>"
end if
rs.close
%>
</table>
<%
end if ' request.form	
set rs=nothing
set rs2=nothing
set rsd=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>