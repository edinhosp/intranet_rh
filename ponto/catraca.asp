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
<title>Consulta de Catraca</title>
<link rel="stylesheet" type="text/css" href="../<%=session("estilo")%>">
<link rel="SHORTCUT ICON" href="../images/rho.png">
</head>
<body>
<%
	dim conexao, conexao2, rs, rs2
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
	'rs.Open sqla, ,adOpenStatic, adLockReadOnly

if request.form="" then
%>
<p class=titulo>Consulta da Catraca
<form method="POST" action="catraca.asp">
<table border="1" cellpadding="3" cellspacing="0" style="border-collapse: collapse">
<tr>
	<td class=grupo>Campus</td>
	<td class=grupo>Chapa</td>
	<td class=grupo>Per�odo</td>
</tr>
<tr>
<td class=titulo>
	<input type="radio" name="campus" value="VY" checked> Vila Yara<br>
	<input type="radio" name="campus" value="NS"> Narciso
</td>
<td class=titulo><input type="text" name="chapa" value="" size="10"></td>
<td class=titulo>
	<input type="text" name="data1" value="<%=int(now())%>" size="8">
	<input type="text" name="data2" value="<%=int(now())%>" size="8">
</td>
</tr>
<tr><td colspan=3 class=titulo>
<input type="submit" value="Clique para Visualizar" class="button" name="B3">
</td></tr>
</table>
</form>
<hr>
<%
else 'request.form <>''
	campus=request.form("campus")
	chapa=numzero(request.form("chapa"),5) & "%"
	data1=request.form("data1")
	data2=request.form("data2")
	if isdate(data1)=true then d1=1 else d1=0
	if isdate(data2)=true then d2=1 else d2=0
	
	response.write "<br>" & request.form
	response.write "<br>" & chapa
	response.write "<br>" & d1 & "-" & d2
	
	if campus="VY" then
		set conexao2=server.createobject ("ADODB.Connection")
		'conexao2.Open application("catracavy")
		'conexao2.open "Driver={SQL Native Client};Server=security;Database=acesso;Integrated Security=SSPI;" 
		'conexao2.open "Provider=sqloledb.1;Data Source=security;Initial Catalog=acesso;uid=athos;password=athos"
		conexao2.open "Provider=SQLOLEDB.1; SERVER=security; DATABASE=acesso; UID=athos; PWD=athos;"
	else
		set conexao2=server.createobject ("ADODB.Connection")
		'conexao2.Open application("catracans")
		'conexao2.open "Driver={SQL Native Client};Server=192.168.100.2;Database=acesso;Integrated Security=SSPI;" 
		'conexao2.open "Provider=sqloledb.1;Data Source=192.168.100.2,1433; Network Library=DBMSSOCN;Initial Catalog=acesso;Integrated Security=SSPI;"
		conexao2.open "Provider=SQLOLEDB.1; SERVER=192.168.100.2,1433; Network Library=DBMSSOCN;; DATABASE=acesso; UID=chef; PWD=n28utu;"
	end if

	
	sqld="select f.nome, f.codsindicato, c.nome as funcao, s.descricao as setor from corporerm.dbo.pfunc f, corporerm.dbo.psecao s, corporerm.dbo.pfuncao c where f.codsecao=s.codigo and f.codfuncao=c.codigo and f.chapa='" & chapa & "'"
	rs2.Open sqld, ,adOpenStatic, adLockReadOnly
	sindicato=rs2("codsindicato")
	if sindicato="03" then coluna=7 else coluna=5
mano=2004
mmes=6
	
%>
<%linha=linha+1%>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="600">
<tr><td class=campo>Espelho de Marca��o de Ponto</td></tr></table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="600">
<tr><td class=campo>FUNDA��O INSTITUTO DE ENSINO PARA OSASCO</td></tr></table>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="600">
<tr><td class=campo>Per�odo: <%=mes & "/" & ano %></td></tr></table>
<hr>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="650">
<tr><td class="campor">Chapa</td>
	<td class="campor">Nome</td>
	<td class="campor">Setor</td>
	<td class="campor">Fun��o</td></tr>
<tr><td class="campor"><%=chapa%></td>
	<td class="campor"><b><%=rs2("nome")%></b></td>
	<td class="campor"><%=rs2("Setor")%></td>
	<td class="campor"><%=rs2("funcao")%></td></tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" >
<tr>
	<td class=grupo align="center" colspan=3 style="border-right:2px solid #000000">Datas</td>
	<td class=grupo align="center" colspan=<%=coluna%> style="border-right:2px solid #000000">Carga Hor�ria a cumprir</td>
	<td class=grupo align="center" colspan=7 style="border-right:2px solid #000000">Marca��es efetuadas</td>
	<td class=grupo align="center" colspan=8 style="border-right:2px solid #000000">Ocorr�ncias</td>
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
	<td class=fundo align="center" >&nbsp;Descri��o Abono&nbsp;</td>
	<td class=fundo align="center" >&nbsp;Inicio&nbsp;</td>
	<td class=fundo align="center" style="border-right:2px solid #000000">&nbsp;Fim&nbsp;</td>
</tr>
<%
rs2.close
totalchcumprir=0
totalchcumprida=0
dim marc(31,6)
dim formato(31,6)
	sqlcr="select chapa, day(data) as dia, data, batida, natureza, status from corporerm.dbo.abatfunam where chapa='" & chapa & "' and month(data)=" & mes & " and year(data)=" & ano & " " & _
	"UNION ALL " & _
	"select chapa, day(data) as dia, data, batida, natureza, status from corporerm.dbo.abatfun where chapa='" & chapa & "' and month(data)=" & mes & " and year(data)=" & ano & " "

	'marca��es do chronus
	rs2.Open sqlcr, ,adOpenStatic, adLockReadOnly
	marcacao=1
	if rs2.recordcount>0 then
		rs2.movefirst
		do while not rs2.eof

		'dia=rs2("dia")
		'batida=formatdatetime((rs2("batida")/60)/24,4)
		'if dia=diaant then marcacao=marcacao+1 else marcacao=1
		'marc(dia,marcacao)=batida
		'if rs2("status")="D" then formato(dia,marcacao)="<font color='red'>" 'else formato(dia,marcacao)="<font color='black'>"
		'diaant=dia
		
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

dtponto=dateserial(ano,mes,1)
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

for a=1 to udia
	dtponto=dateserial(ano,mes,a)
	if a=1 then indice=indice else indice=indice+1
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

	response.write "<td class=campo align=""center"">" & dtponto & "</td>"
	response.write "<td class=campo align=""center"">" & weekdayname(weekday(dtponto),-1) & "</td>"
	response.write "<td class=campo align=""center"" style='border-right:2px solid #000000'><font color=gray>" & indice & "</td>"
	
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
		batida=marc(a,b)
		if batida<>"" then ultima=b
		response.write "<td class=campo align=""center"">" & formato(a,b) & batida & "</font></td>"
	next
	If marc(a,1)="" and marc(a,2)="" and marc(a,3)="" and marc(a,4)="" and marc(a,5)="" and marc(a,6)="" then
		tot1=0:tot2=0:tot3=0
	else
		if marc(a,2)="" and marc(a,1)<>"" then tot1=0 else tot1=cdate(marc(a,2))-cdate(marc(a,1))
		if marc(a,4)="" and marc(a,3)<>"" then tot2=cdate(marc(a,3))-cdate(marc(a,2)) else tot2=cdate(marc(a,4))-cdate(marc(a,3))
		if marc(a,6)="" and marc(a,5)<>"" then tot3=cdate(marc(a,5))-cdate(marc(a,4)) else tot3=cdate(marc(a,6))-cdate(marc(a,5))
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
<p style="margin-top: 0; margin-bottom: 0"><font size=1><b>Nomea��es para Atividades</b></font></p>
<table border="1" cellpadding="1" cellspacing="1" style="border-collapse: collapse" width="650">
<tr>
	<td class=titulor>Nomea��o</td>
	<td class=titulor>Portaria</td>
	<td class=titulor>Curso</td>
	<td class=titulor>Evento</td>
	<td class=titulor>CHS</td>
	<td class=titulor>CHM</td>
	<td class=titulor>Inicio</td>
	<td class=titulor>T�rmino</td>
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
	response.write "<tr><td class=""campor"" colspan=8>Sem nomea��es</td></tr>"
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