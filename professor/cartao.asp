<%@ Language=VBScript %>
<!-- #Include file="../adovbs.inc" -->
<!-- #Include file="../funcoes.inc" -->
<%
	Response.buffer=true
	Server.ScriptTimeout = 600
if Session("UsuarioMaster")="" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
if session("a38")<>"T" then response.write "<script language='JavaScript' type='text/javascript'>window.top.location.href='" & Application("Site") & "';</script>"
%>
<html>
<head>
<meta http-equiv="CONTENT-TYPE" content="text/html; charset=windows-1252">
<title>Cartão de Ponto - Prof</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="images/rho.png">
</head>
<body>
<%
dim conexao, conexao2, chapach, rs, rs2
set conexao=server.createobject ("ADODB.Connection")
conexao.Open application("conexao")
sqla="SELECT dc_carga.CURSO FROM dc_carga GROUP BY dc_carga.CURSO;"
set rs=server.createobject ("ADODB.Recordset")
Set rs.ActiveConnection = conexao
set rs2=server.createobject ("ADODB.Recordset")
Set rs2.ActiveConnection = conexao
set rsd=server.createobject ("ADODB.Recordset")
Set rsd.ActiveConnection = conexao
'rs.Open sqla, ,adOpenStatic, adLockReadOnly

if request.form="" then
%>
<p class=titulo>Espelho de Marcação de Ponto
<form method="POST" action="cartao.asp">
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse">
<tr><td class=titulo>Chapa</td><td class=titulo>Mês</td><td class=titulo>Ano</td><td class=titulo></td></tr>
<tr>
<td class=titulo><input type="text" name="chapa" value="" size="5"></td>
<td class=titulo><input type="text" name="mes" value="<%=month(now)%>" size="2"></td>
<td class=titulo><input type="text" name="ano" value="<%=year(now)%>" size="4"></td>
<td class=fundo>&nbsp;Imprime foto?<input type="checkbox" name="foto" value="ON"></td>
</tr>
<tr><td colspan=4 class=titulo>
<input type="submit" value="Clique para Visualizar" class="button" name="B3">
</td></tr>
</table>
</form>
<hr>
<%
else 'request.form <>''
	mes=request.form("mes")
	ano=request.form("ano")
	udia=day(dateserial(ano,mes+1,1)-1)
	chapa=numzero(request.form("chapa"),5)
	sqld="select f.nome, c.nome as funcao, s.descricao as setor from corporerm.dbo.pfunc f, corporerm.dbo.psecao s, corporerm.dbo.pfuncao c where f.codsecao=s.codigo and f.codfuncao=c.codigo and f.chapa='" & chapa & "'"
	rsd.Open sqld, ,adOpenStatic, adLockReadOnly
%>

<%linha=linha+1%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="600">
<tr><td class=campo width=430>Espelho de Marcação de Ponto</td>
	<td width="170" class=campo valign="top" rowspan=3>
<% if request.form("foto")="ON" then %>
		<img border="0" src="../func_foto.asp?chapa=<%=chapa%>"  width="150">
<% end if %>
	</td>
</tr>
<tr><td class=campo>FUNDAÇÃO INSTITUTO DE ENSINO PARA OSASCO</td></tr>
<tr><td class=campo>Período: <%=mes & "/" & ano %></td></tr>
</table>
<hr>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class="campor">Chapa</td>
	<td class="campor">Nome</td>
	<td class="campor">Setor</td>
	<td class="campor">Função</td></tr>
<tr><td class="campor"><%=chapa%></td>
	<td class="campor"><b><%=rsd("nome")%></b></td>
	<td class="campor"><%=rsd("Setor")%></td>
	<td class="campor"><%=rsd("funcao")%></td></tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
  <tr>
    <td class=grupo align="center" colspan=2>Datas</td>
    <td class=grupo align="center" colspan=7>Carga Horária a cumprir</td>
	<td class=grupo>&nbsp;</td>
    <td class=grupo align="center" colspan=7>Marcações efetuadas</td>
    <td class=grupo align="center">Saldo</td>
	<td class=grupo>&nbsp;</td>
  </tr>

  <tr>
    <td class=titulo align="center">Data</td>
    <td class=titulo align="center">Dia</td>
    <td class=titulo align="center" width=35>1</td>
    <td class=titulo align="center" width=35>2</td>
    <td class=titulo align="center" width=35>3</td>
    <td class=titulo align="center" width=35>4</td>
    <td class=titulo align="center" width=35>5</td>
    <td class=titulo align="center" width=35>6</td>
    <td class=titulo align="center" width=40>=</td>
	<td class=grupo>&nbsp;</td>
    <td class=titulo align="center" width=35>1</td>
    <td class=titulo align="center" width=35>2</td>
    <td class=titulo align="center" width=35>3</td>
    <td class=titulo align="center" width=35>4</td>
    <td class=titulo align="center" width=35>5</td>
    <td class=titulo align="center" width=35>6</td>
    <td class=titulo align="center" width=40>=</td>
	<td class=titulo>&nbsp;</td>
	<td class=grupo>&nbsp;</td>
  </tr>
<%
rsd.close
totalchcumprir=0
totalchcumprida=0
dim marc(31,6)
dim formato(31,6)
dim htrab(31)
'dim nat(31,6)
	sqlcr="select chapa, day(data) as dia, data, batida, natureza, status from corporerm.dbo.abatfunam where chapa='" & chapa & "' and month(data)=" & mes & " and year(data)=" & ano & " " & _
	"UNION ALL " & _
	"select chapa, day(data) as dia, data, batida, natureza, status from corporerm.dbo.abatfun where chapa='" & chapa & "' and month(data)=" & mes & " and year(data)=" & ano & " "

	'marcações do chronus
	rs2.Open sqlcr, ,adOpenStatic, adLockReadOnly
	marcacao=1
	if rs2.recordcount>0 then
		rs2.movefirst
		do while not rs2.eof
		dia=rs2("dia")
		natureza=rs2("natureza")
		batida=formatdatetime((rs2("batida")/60)/24,4)
		if dia=diaant then marcacao=marcacao+1 else marcacao=1
		'nat(dia,marcacao)=rs2("natureza")
		if natureza=0 or natureza=4 then natu=0
		if natureza=1 or natureza=5 then natu=1
		resto=marcacao mod 2
		if resto=0 and natu=0 then marcacao=marcacao+1 else marcacao=marcacao
		if resto<>0 and natu=1 then marcacao=marcacao+1 else marcacao=marcacao
		marc(dia,marcacao)=batida
		if rs2("status")="D" then formato(dia,marcacao)="<font color='red'>" 'else formato(dia,marcacao)="<font color='black'>"
		diaant=dia
		rs2.movenext
		loop
	else 'recordcount rs2
		for a=1 to 31
			for b=1 to 6
				marc(a,b)=""
			next
		next
	end if 'recordcount rs2
	rs2.close

'horas trabalhadas pelo chronus	
for t=1 to udia
	datah=dateserial(ano,mes,t)
	sqlh="select htrab from corporerm.dbo.aafhtfunam where chapa='" & chapa & "' and data='" & dtaccess(datah) & "' " & _
	"UNION ALL " & _
	"select htrab from corporerm.dbo.aafhtfun where chapa='" & chapa & "' and data='" & dtaccess(datah) & "' "
	rs2.Open sqlh, ,adOpenStatic, adLockReadOnly
	if rs2.recordcount>0 then
		if rs2("htrab")>0 then htrab(t)=rs2("htrab") else htrab(t)=0
	end if
	rs2.close
next 't
	
for a=1 to udia
	dtponto=dateserial(ano,mes,a)
	response.write "<tr>"
	sqlch="select chapa, dia_mes, diasem, ent1, saida1, ent2, saida2, ent3, saida3, totalch " & _
	"from ttapontprof_2 " & _
	"where chapa='" & chapa & "' and dia_mes='" & dtaccess(dtponto) & "' "
	'grade horaria
	rs.Open sqlch, ,adOpenStatic, adLockReadOnly
	if rs.recordcount>0 then
		if isdate(rs("ent1"))   then ent1  =formatdatetime(rs("ent1"),4)   else ent1  ="-"
		if isdate(rs("saida1")) then saida1=formatdatetime(rs("saida1"),4) else saida1="-"
		if isdate(rs("ent2"))   then ent2  =formatdatetime(rs("ent2"),4)   else ent2  ="-"
		if isdate(rs("saida2")) then saida2=formatdatetime(rs("saida2"),4) else saida2="-"
		if isdate(rs("ent3"))   then ent3  =formatdatetime(rs("ent3"),4)   else ent3  ="-"
		if isdate(rs("saida3")) then saida3=formatdatetime(rs("saida3"),4) else saida3="-"
		response.write "<td class=campo align=""center"">" & rs("dia_mes") & "</td>"
		response.write "<td class=campo align=""center"">" & weekdayname(weekday(dtponto),-1) & "</td>"
		response.write "<td class=campo align=""center"">" & ent1   & "</td>"
		response.write "<td class=campo align=""center"">" & saida1 & "</td>"
		response.write "<td class=campo align=""center"">" & ent2   & "</td>"
		response.write "<td class=campo align=""center"">" & saida2 & "</td>"
		response.write "<td class=campo align=""center"">" & ent3   & "</td>"
		response.write "<td class=campo align=""center"">" & saida3 & "</td>"
		response.write "<td class=campo align=""center"">" & formatdatetime(rs("totalch"),4) & "</td>"
		tdia1=cdate(formatdatetime(rs("totalch"),4))
		tcumprir=tcumprir+tdia1 ' rs("totalch")
		tchcumprir=tcumprir
	else 'recordcount rs
		response.write "<td class=campo align=""center"">" & dtponto & "</td>"
		response.write "<td class=campo align=""center"">" & weekdayname(weekday(dtponto),-1) & "</td>"
		response.write "<td class=titulor align=""center"" colspan=6>&nbsp;</td>"
		response.write "<td class=campo align=""center"">-</td>"
		ent1="":ent2="":ent3="":saida1="":saida2="":saida3=""
		tdia1=0
	end if 'recordcount rs
	if weekday(dtponto)=1 then tsem1=tchcumprir:tchcumprir=0
	response.write "<td class=grupo align=""center"">&nbsp;</td>"
	
	rs.close
		for b=1 to 6
			batida=marc(a,b)
			response.write "<td class=campo align=""center"">" & formato(a,b) & batida & "</font></td>"
		next
		If marc(a,1)="" and marc(a,2)="" and marc(a,3)="" and marc(a,4)="" and marc(a,5)="" and marc(a,6)="" then
			tot1=0:tot2=0:tot3=0
		else
		if (marc(a,2)="" and marc(a,1)<>"") or (marc(a,1)="" and marc(a,2)<>"") then tot1=0 else tot1=cdate(marc(a,2))-cdate(marc(a,1))
		if marc(a,4)="" and marc(a,3)<>"" then tot2=cdate(marc(a,3))-cdate(marc(a,2)) else tot2=cdate(marc(a,4))-cdate(marc(a,3))
		if marc(a,6)="" and marc(a,5)<>"" then tot3=cdate(marc(a,5))-cdate(marc(a,4)) else tot3=cdate(marc(a,6))-cdate(marc(a,5))
		end if
		if htrab(a)=0 then htrabi="-" else htrabi=formatdatetime((htrab(a)/60)/24,4) 
		response.write "<td class=campo align=""center"">" & htrabi & "</td>"
		tcumprida=tcumprida+((htrab(a)/60)/24)
		saldo=(htrab(a)/60)/24 -tdia1
		saldo=cdbl(saldo)*24
		if saldo=0 then saldoh="-" else saldoh=formatnumber(saldo,2)
		if saldo>0.25 then saldoa=saldoa+saldo
		tsaldo=tsaldo+saldo
		response.write "<td class=campo align=""center"">" & saldoh & "</td>"
		response.write "<td class=grupo align=""center"">&nbsp;</td>"
	response.write "</tr>"
next
	if weekday(dtponto)=1 then
		tsem2=tchcumprida
		tchcumprida=0
	end if
tcumprir =int((cdbl(tcumprir) *24)*100+0.5)/100
tcumprida=int((cdbl(tcumprida)*24)*100+0.5)/100
%>
  <tr>
    <td class=titulo align="left" colspan=8>Totais</td>
    <td class=campo align="center"><%=formatnumber(tcumprir,2)%></td>
	<td class=grupo>&nbsp;</td>
    <td class=titulo align="left" colspan=6>Totais</td>
    <td class=campo align="center"><%=formatnumber(tcumprida,2)%></td>
    <td class=campo align="center"><%=formatnumber(tsaldo,2)%></td>
	<td class=grupo align="center">&nbsp;</td>
  </tr>
<tr>
	<td class=campo colspan=16 align="right">Saldo para atividades&nbsp;</td>
	<!--'=formatnumber(totalchcumprida1-totalchcumprir,2)-->
	<td class=campo align="center"><%%><%=formatnumber(saldoa,2)%></td>
	<td class=titulo colspan=2>&nbsp;</td>
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
sqln="SELECT n.NOMEACAO, i.PORTARIA, i.CARGO, i.CODEVE, i.CH, i.MAND_INI, i.MAND_FIM, i.id_nomeacao " & _
"FROM n_nomeacoes n, n_indicacoes i WHERE n.id_nomeacao = i.id_nomeacao " & _
"AND i.CHAPA='" & chapa & "' AND ('" & dtaccess(dateserial(ano,mes,1)) & "' Between mand_ini And mand_fim " & _
"OR '" & dtaccess(dateserial(ano,mes,28)) & "' Between mand_ini And mand_fim )"
rs.Open sqln, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
if rs("id_nomeacao")=12 then fator=4.5 else fator=4
%>
<tr>
	<td class="campor"><%=rs("nomeacao")%></td>
	<td class="campor"><%=rs("portaria")%></td>
	<td class="campor"><%=rs("cargo")%></td>
	<td class="campor"><%=rs("codeve")%></td>
	<td class="campor" align="center"><%=rs("ch")%></td>
	<td class="campor" align="center"><%if rs("ch")<>"" then response.write rs("ch")*fator else response.write "&nbsp;"%></td>
	<td class="campor"><%=rs("mand_ini")%></td>
	<td class="campor"><%=rs("mand_fim")%></td>
</tr>
<%
if isnumeric(rs("ch")) then taes=taes+cdbl(rs("ch"))
if isnumeric(rs("ch")) then taem=taem+(cdbl(rs("ch"))*fator)
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

<p style="margin-top: 0; margin-bottom: 0"><font size=1><b>Aulas na Pós-Graduação</b></font></p>
<table border="1" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="650">
<tr>
	<td class=titulor>Curso</td>
	<td class=titulor>Disciplina</td>
	<td class=titulor>Dia</td>
	<td class=titulor>Horário</td>
	<td class=titulor>Aulas</td>
	<td class=titulor>Inicio</td>
	<td class=titulor>Término</td>
</tr>

<%
sqln="select chapa1, perlet, coddoc curso, materia, inicio, termino, count(ta) as aulas " & _
"from g5ch " & _
"where chapa1='" & chapa & "' " & _
"and ('" & dtaccess(dateserial(ano,mes,1)) & "' between inicio and termino " & _
"or '" & dtaccess(dateserial(ano,mes,28)) & "' between inicio and termino) " & _
"group by chapa1, perlet, coddoc, materia, inicio, termino "
rs.Open sqln, ,adOpenStatic, adLockReadOnly
if rs.recordcount>0 then
rs.movefirst
do while not rs.eof
%>
<tr>
	<td class="campor"><%=rs("curso")%></td>
	<td class="campor"><%=rs("materia")%></td>
	<td class="campor"><%%></td>
	<td class="campor"><%%></td>
	<td class="campor" align="center"><%=rs("aulas")%></td>
	<td class="campor"><%=rs("inicio")%></td>
	<td class="campor"><%=rs("termino")%></td>
</tr>
<%
rs.movenext
loop
%>
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