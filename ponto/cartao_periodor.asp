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

	chapa="01065"
	sqld="select f.nome, f.codsindicato, c.nome as funcao, s.descricao as setor from corporerm.dbo.pfunc f, corporerm.dbo.psecao s, corporerm.dbo.pfuncao c where f.codsecao=s.codigo and f.codfuncao=c.codigo and f.chapa='" & chapa & "'"
	rs2.Open sqld, ,adOpenStatic, adLockReadOnly
	sindicato=rs2("codsindicato")
	if sindicato="03" then coluna=7 else coluna=5
	nome=rs2("nome"):setor=rs2("setor"):funcao=rs2("funcao")
	rs2.close
mano=2004
mmes=6

for a=2005 to 2008
if a=2005 then m0=12 else m0=1
for m=m0 to 12
	data1=dateserial(a,m,16)
	data2=dateserial(a,m+1,15)
	data1=cdate(data1):data2=cdate(data2)
	udia=day(dateserial(ano,mes+1,1)-1)
	tcumprida1=0:tcumprida2=0
	
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
	<td class="campor"><b><%=nome%></b></td>
	<td class="campor"><%=Setor%></td>
	<td class="campor"><%=funcao%></td></tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" >
<tr>
	<td class=grupo align="center" colspan=2 style="border-right:2px solid #000000">Datas</td>
	<td class=grupo align="center" colspan=7 style="border-right:2px solid #000000">Marcações efetuadas</td>
</tr>
<tr>
	<td class=titulo align="center">Data</td>
	<td class=titulo align="center" style="border-right:2px solid #000000">Dia</td>
	<td class=titulo align="center" width=35>1</td>
	<td class=titulo align="center" width=35>2</td>
	<td class=titulo align="center" width=35>3</td>
	<td class=titulo align="center" width=35>4</td>
	<td class=titulo align="center" width=35>5</td>
	<td class=titulo align="center" width=35>6</td>
	<td class=fundo align="center" style="border-right:2px solid #000000">H.Trab.</td>
</tr>
<%

diasloop=datediff("d",data1,data2)+1:'response.write diasloop & "<br>"
diasloop=cint(diasloop)
totalchcumprir=0
totalchcumprida=0

Redim marc(diasloop,6), formato(diasloop,6)

for e=data1 to data2
	idmatriz=e-(data1)
	'marcações do chronus
	sqlcr="select chapa, day(data) as dia, data, batida, natureza, status from corporerm.dbo.abatfunam where chapa='" & chapa & "' " & _
	"and data='" & dtaccess(e) & "' " & _
	"UNION ALL " & _
	"select chapa, day(data) as dia, data, batida, natureza, status from corporerm.dbo.abatfun where chapa='" & chapa & "' " & _
	"and data='" & dtaccess(e) & "' "
	rs2.Open sqlcr, ,adOpenStatic, adLockReadOnly
	marcacao=1
	if rs2.recordcount>0 then
		rs2.movefirst:do while not rs2.eof
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
		marc(idmatriz,marcacao)=batida:'response.write ">> " & idmatriz & " >> " & marc(idmatriz,marcacao) & "<br>"
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

for e=data1 to data2
	'dtponto=dateserial(ano,mes,a)
	dtponto=e
	idmatriz=e-(data1)
	if idmatriz=0 then indice=indice else indice=indice+1
	response.write "<tr>"

	response.write "<td class=campo align=""center"">" & dtponto & "</td>"
	response.write "<td class=campo align=""center"" style='border-right:2px solid #000000'>" & weekdayname(weekday(dtponto),-1) & "</td>"
	
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
		
	response.write "</tr>"
next

if request.form("considerar")="ON" then totalgeral=tcumprida2 else totalgeral=tcumprida1
%>
  <tr>
    <td class=titulo align="left" colspan=8>&nbsp;Totais</td>
    <td class=campo align="center" style="border-right:2px solid #000000"><%=formatnumber(totalgeral/60,2)%></td>
  </tr>
  
</table>

<%

response.write "<DIV style=""page-break-after:always""></DIV>"
next 'mes
next 'ano

%>

<%

set rs=nothing
set rs2=nothing
set rsd=nothing
conexao.close
set conexao=nothing
%>
</body>
</html>